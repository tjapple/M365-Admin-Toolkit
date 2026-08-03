#Requires -Version 7.2

<#
.SYNOPSIS
Interactively reviews and changes Microsoft 365 group memberships.

.DESCRIPTION
Uses delegated Microsoft Graph and Exchange Online sessions. Mutations require
explicit confirmation. Use -DryRun or -WhatIf to preview every change.

.EXAMPLE
./groupmanager.ps1 -UserPrincipalName alex.wilber@contoso.com -DryRun

.EXAMPLE
./groupmanager.ps1 -UserPrincipalName alex.wilber@contoso.com -WhatIf

.PARAMETER BypassGroupOwnershipCheck
Allows an Exchange administrator to manage a distribution group they do not
own. This powerful override is disabled by default.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
param(
    [string]$UserPrincipalName,

    [switch]$SkipConnect,

    [switch]$DryRun,

    [switch]$BypassGroupOwnershipCheck
)

$toolkitModulePath = Join-Path $PSScriptRoot "..\src\M365AdminToolkit\M365AdminToolkit.psd1"
Import-Module $toolkitModulePath -Force -ErrorAction Stop
Set-M365ToolkitUiWidth -Width 120

$script:DryRun = [bool]($DryRun -or $WhatIfPreference)
$script:BypassGroupOwnershipCheck = [bool]$BypassGroupOwnershipCheck
$script:CloudGroupSessionState = @{}
$script:GroupManagerExchangeConnected = $false

function Get-GroupManagerGraphScopes {
    return @(
        "User.Read.All",
        "GroupMember.ReadWrite.All"
    )
}

function Connect-GroupManagerGraph {
    Import-RequiredPowerShellModule -ModuleName "Microsoft.Graph.Authentication"

    Write-UiStatus -Status "LOADING..." -Message "Connecting to Microsoft Graph." -Color Cyan
    try {
        Connect-MgGraph -NoWelcome -ContextScope Process -Scopes (Get-GroupManagerGraphScopes)
    }
    catch {
        if ([string]$_.Exception.Message -notmatch "ContextScope") {
            throw
        }
        Write-UiStatus -Status "WARN" -Message "This Microsoft.Graph.Authentication version does not support -ContextScope. Connecting without process-scoped context." -Color Yellow
        Connect-MgGraph -NoWelcome -Scopes (Get-GroupManagerGraphScopes)
    }

    Write-UiStatus -Status "OK" -Message "Connected to Microsoft Graph." -Color Green
}

function Connect-GroupManagerExchangeOnline {
    Import-RequiredPowerShellModule -ModuleName "ExchangeOnlineManagement"

    try {
        if (Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $script:GroupManagerExchangeConnected = $true
            Write-UiStatus -Status "OK" -Message "Reusing the existing Exchange Online session." -Color Green
            return
        }
    }
    catch {
        Write-Verbose "Could not inspect the current Exchange Online session: $($_.Exception.Message)"
    }

    Write-UiStatus -Status "LOADING..." -Message "Connecting to Exchange Online." -Color Cyan
    Connect-ExchangeOnline -ShowBanner:$false
    $script:GroupManagerExchangeConnected = $true
    Write-UiStatus -Status "OK" -Message "Connected to Exchange Online." -Color Green
}

function Ensure-GroupManagerExchangeOnlineConnected {
    if ($script:GroupManagerExchangeConnected) {
        return
    }

    try {
        $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($connection) {
            $script:GroupManagerExchangeConnected = $true
            return
        }
    }
    catch {
        Write-Verbose "Could not verify Exchange Online connection state: $($_.Exception.Message)"
    }

    Connect-GroupManagerExchangeOnline
}

function Connect-GroupManagerServices {
    Connect-GroupManagerGraph
}

function Get-MgGraphExceptionText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $messages = @()
    if ($ErrorRecord.Exception -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.Exception.Message)) {
        $messages += $ErrorRecord.Exception.Message
    }

    if ($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.ErrorDetails.Message)) {
        $detailsText = [string]$ErrorRecord.ErrorDetails.Message
        try {
            $details = $detailsText | ConvertFrom-Json -ErrorAction Stop
            if ($details.error) {
                if (-not [string]::IsNullOrWhiteSpace([string]$details.error.code)) {
                    $messages += ("Graph code: {0}" -f $details.error.code)
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$details.error.message)) {
                    $messages += ("Graph message: {0}" -f $details.error.message)
                }
            }
            else {
                $messages += $detailsText
            }
        }
        catch {
            $messages += $detailsText
        }
    }

    $cleanMessages = @($messages | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
    if ($cleanMessages.Count -eq 0) {
        return "Graph request failed, but no detailed error text was returned."
    }

    return ($cleanMessages -join " | ")
}

function Escape-GraphFilterValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Replace("'", "''")
}

function New-GraphUserLookupFilter {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Lookup
    )

    $escapedLookup = Escape-GraphFilterValue -Value $Lookup
    $nameParts = @($Lookup -split "\s+" | Where-Object { $_ })
    $filter = "userPrincipalName eq '$escapedLookup' or mail eq '$escapedLookup' or mailNickname eq '$escapedLookup' or displayName eq '$escapedLookup'"

    if ($nameParts.Count -ge 2) {
        $escapedFirstName = Escape-GraphFilterValue -Value $nameParts[0]
        $escapedLastName = Escape-GraphFilterValue -Value $nameParts[$nameParts.Count - 1]
        $filter = "$filter or (givenName eq '$escapedFirstName' and surname eq '$escapedLastName')"
    }

    return $filter
}

function Invoke-GraphUserFilterSearch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Filter,

        [Parameter(Mandatory = $true)]
        [string]$Select
    )

    $encodedFilter = [System.Uri]::EscapeDataString($Filter)
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$encodedFilter" + [char]38 + "`$select=$Select"
    $result = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    return @($result.value | Where-Object { $_ })
}

function Get-UniqueGraphObjects {
    param(
        [AllowNull()]
        [object[]]$Objects
    )

    $uniqueObjects = @()
    $seenIds = @{}
    foreach ($object in @($Objects | Where-Object { $_ })) {
        $key = [string]$object.id
        if ([string]::IsNullOrWhiteSpace($key)) {
            $key = [string]$object.userPrincipalName
        }
        if ([string]::IsNullOrWhiteSpace($key)) {
            $key = [string]$object.mail
        }
        if ([string]::IsNullOrWhiteSpace($key)) {
            $key = [string]$object.displayName
        }
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }

        $key = $key.Trim().ToLowerInvariant()
        if ($seenIds.ContainsKey($key)) {
            continue
        }

        $seenIds[$key] = $true
        $uniqueObjects += $object
    }

    return $uniqueObjects
}

function Test-CloudUserExactLookupMatch {
    param(
        [AllowNull()]
        $CloudUser,

        [Parameter(Mandatory = $true)]
        [string]$Lookup
    )

    if ($null -eq $CloudUser -or [string]::IsNullOrWhiteSpace($Lookup)) {
        return $false
    }

    $lookupText = $Lookup.Trim()
    $comparison = [System.StringComparison]::OrdinalIgnoreCase

    foreach ($propertyName in @("displayName", "userPrincipalName", "mail", "mailNickname", "onPremisesSamAccountName", "id")) {
        $value = $CloudUser.$propertyName
        if (-not [string]::IsNullOrWhiteSpace([string]$value) -and [string]::Equals(([string]$value).Trim(), $lookupText, $comparison)) {
            return $true
        }
    }

    $nameParts = @(
        [string]$CloudUser.givenName
        [string]$CloudUser.surname
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
    $fullName = ($nameParts -join " ").Trim()
    if (-not [string]::IsNullOrWhiteSpace($fullName) -and [string]::Equals($fullName, $lookupText, $comparison)) {
        return $true
    }

    foreach ($proxyAddress in @($CloudUser.proxyAddresses | Where-Object { $_ })) {
        $addressText = [string]$proxyAddress
        if (-not [string]::IsNullOrWhiteSpace($addressText) -and [string]::Equals($addressText.Trim(), $lookupText, $comparison)) {
            return $true
        }

        if ($addressText -match "^[^:]+:(.+)$") {
            $addressText = $Matches[1]
        }

        if (-not [string]::IsNullOrWhiteSpace($addressText) -and [string]::Equals($addressText.Trim(), $lookupText, $comparison)) {
            return $true
        }
    }

    return $false
}

function Find-ActiveCloudUsersByLookup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Lookup,

        [Parameter(Mandatory = $true)]
        [string]$Select
    )

    $candidateMatches = @()
    $filter = New-GraphUserLookupFilter -Lookup $Lookup
    $candidateMatches += Invoke-GraphUserFilterSearch -Filter $filter -Select $Select

    $escapedLookup = Escape-GraphFilterValue -Value $Lookup
    try {
        $candidateMatches += Invoke-GraphUserFilterSearch -Filter "onPremisesSamAccountName eq '$escapedLookup'" -Select $Select
    }
    catch {
        Write-Verbose "The optional on-premises account-name filter is unavailable: $($_.Exception.Message)"
    }

    if ($Lookup -like "*@*") {
        $escapedUpperProxy = Escape-GraphFilterValue -Value "SMTP:$Lookup"
        $escapedLowerProxy = Escape-GraphFilterValue -Value "smtp:$Lookup"
        $proxyFilter = "proxyAddresses/any(p:p eq '$escapedUpperProxy') or proxyAddresses/any(p:p eq '$escapedLowerProxy')"
        try {
            $candidateMatches += Invoke-GraphUserFilterSearch -Filter $proxyFilter -Select $Select
        }
        catch {
            Write-UiStatus -Status "INFO" -Message "Proxy address lookup was not available for '$Lookup': $($_.Exception.Message)" -Color Yellow
        }
    }

    return Get-UniqueGraphObjects -Objects $candidateMatches
}

function Show-GroupManagerCloudUserMatches {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Matches
    )

    $rows = @(
        @($Matches | Where-Object { $_ }) |
            Select-Object displayName,userPrincipalName,mail,mailNickname,accountEnabled,id
    )
    Write-UiBox -Title "Cloud User Matches" -Lines (ConvertTo-UiTableLines -Rows $rows -Columns @("displayName", "userPrincipalName", "mail", "mailNickname", "accountEnabled", "id"))
}

function Resolve-GroupManagerCloudUser {
    param(
        [AllowNull()]
        [string]$Lookup
    )

    $candidateLookup = $Lookup
    $select = "id,displayName,userPrincipalName,mail,mailNickname,givenName,surname,accountEnabled,usageLocation,onPremisesSamAccountName,proxyAddresses"
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($candidateLookup)) {
            $candidateLookup = Read-UiInput -Prompt "Target cloud user lookup" -Options @("exact UPN/email/name", "x=exit")
            if ($candidateLookup -eq "x") {
                return $null
            }
            continue
        }

        $trimmedLookup = $candidateLookup.Trim()
        Write-UiStatus -Status "LOOKUP" -Message "Resolving cloud user '$trimmedLookup'." -Color Cyan
        try {
            $candidateMatches = @(Find-ActiveCloudUsersByLookup -Lookup $trimmedLookup -Select $select)
        }
        catch {
            Write-UiStatus -Status "WARN" -Message "Could not search cloud users: $(Get-MgGraphExceptionText -ErrorRecord $_)" -Color Yellow
            $candidateMatches = @()
        }
        $candidateMatches = @($candidateMatches | Where-Object { Test-CloudUserExactLookupMatch -CloudUser $_ -Lookup $trimmedLookup })

        if ($candidateMatches.Count -eq 1) {
            $match = $candidateMatches[0]
            if ([string]::IsNullOrWhiteSpace([string]$match.userPrincipalName)) {
                Write-UiStatus -Status "WARN" -Message "Matched cloud user does not have a UserPrincipalName. Use a different lookup." -Color Yellow
            }
            else {
                Write-UiBox -Title "Selected Cloud User" -Lines @(
                    New-UiBoxLine -Label "DisplayName" -Value $match.displayName
                    New-UiBoxLine -Label "UserPrincipalName" -Value $match.userPrincipalName
                    New-UiBoxLine -Label "Mail" -Value $match.mail
                    New-UiBoxLine -Label "UsageLocation" -Value $match.usageLocation
                    New-UiBoxLine -Label "AccountEnabled" -Value $match.accountEnabled
                )
                return $match
            }
        }
        elseif ($candidateMatches.Count -eq 0) {
            Write-UiStatus -Status "WARN" -Message "No exact cloud user matched '$trimmedLookup'." -Color Yellow
        }
        else {
            Write-UiStatus -Status "TOO MANY" -Message "Multiple exact cloud users matched '$trimmedLookup'. Use a unique exact value such as UPN or object id." -Color Yellow
            Show-GroupManagerCloudUserMatches -Matches $candidateMatches
        }

        $candidateLookup = Read-UiInput -Prompt "Enter a different cloud user lookup" -Options @("exact UPN/email/name", "x=exit")
        if ($candidateLookup -eq "x") {
            return $null
        }
    }
}

function ConvertTo-NumberSelection {
    param(
        [string]$InputText,
        [int]$Max
    )

    if ([string]::IsNullOrWhiteSpace($InputText)) {
        return @()
    }

    $parts = @($InputText -split "[,\s;]+" | Where-Object { $_ })
    foreach ($part in $parts) {
        if ($part -notmatch "^\d+$" -or [int]$part -lt 1 -or [int]$part -gt $Max) {
            return @()
        }
    }

    $numbers = @()
    foreach ($part in $parts) {
        $number = [int]$part
        if ($numbers -notcontains $number) {
            $numbers += $number
        }
    }

    return $numbers
}

function Select-ItemsByNumber {
    param(
        [AllowNull()]
        [object[]]$Items,

        [AllowNull()]
        [int[]]$Numbers
    )

    $cleanItems = @($Items | Where-Object { $_ })
    $selectedItems = @()
    foreach ($number in @($Numbers)) {
        $index = [int]$number - 1
        if ($index -ge 0 -and $index -lt $cleanItems.Count) {
            $selectedItems += $cleanItems[$index]
        }
    }

    return $selectedItems
}

function Get-CloudGroupSessionUserKey {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser
    )

    foreach ($propertyName in @("id", "userPrincipalName", "mail")) {
        $value = [string]$CloudUser.$propertyName
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim().ToLowerInvariant()
        }
    }

    return ""
}

function Get-CloudGroupSessionGroupKey {
    param(
        [Parameter(Mandatory = $true)]
        $Group
    )

    foreach ($propertyName in @("id", "mail", "mailNickname", "displayName")) {
        $value = [string]$Group.$propertyName
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim().ToLowerInvariant()
        }
    }

    return ""
}

function Get-CloudGroupSessionState {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser
    )

    $userKey = Get-CloudGroupSessionUserKey -CloudUser $CloudUser
    if ([string]::IsNullOrWhiteSpace($userKey)) {
        return $null
    }

    if (-not $script:CloudGroupSessionState) {
        $script:CloudGroupSessionState = @{}
    }

    if (-not $script:CloudGroupSessionState.ContainsKey($userKey)) {
        $script:CloudGroupSessionState[$userKey] = @{
            Added   = @{}
            Removed = @{}
        }
    }

    return $script:CloudGroupSessionState[$userKey]
}

function Update-CloudGroupSessionState {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser,

        [Parameter(Mandatory = $true)]
        $Group,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Add", "Remove")]
        [string]$Action
    )

    $state = Get-CloudGroupSessionState -CloudUser $CloudUser
    $groupKey = Get-CloudGroupSessionGroupKey -Group $Group
    if ($null -eq $state -or [string]::IsNullOrWhiteSpace($groupKey)) {
        return
    }

    if ($Action -eq "Add") {
        $state.Removed.Remove($groupKey)
        $state.Added[$groupKey] = $Group
        return
    }

    $state.Added.Remove($groupKey)
    $state.Removed[$groupKey] = $true
}

function Merge-CloudGroupSessionState {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser,

        [AllowNull()]
        [object[]]$Groups
    )

    $state = Get-CloudGroupSessionState -CloudUser $CloudUser
    if ($null -eq $state) {
        return @($Groups | Where-Object { $_ } | Sort-Object displayName)
    }

    $merged = @()
    $seen = @{}
    foreach ($group in @($Groups | Where-Object { $_ })) {
        $groupKey = Get-CloudGroupSessionGroupKey -Group $group
        if ([string]::IsNullOrWhiteSpace($groupKey) -or $state.Removed.ContainsKey($groupKey)) {
            continue
        }

        if (-not $seen.ContainsKey($groupKey)) {
            $merged += $group
            $seen[$groupKey] = $true
        }
    }

    foreach ($groupKey in @($state.Added.Keys)) {
        if (-not $seen.ContainsKey($groupKey) -and -not $state.Removed.ContainsKey($groupKey)) {
            $merged += $state.Added[$groupKey]
            $seen[$groupKey] = $true
        }
    }

    return @($merged | Sort-Object displayName)
}

function Get-CloudUserGroups {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser
    )

    $encodedUser = [System.Uri]::EscapeDataString($CloudUser.id)
    $select = "id,displayName,mailEnabled,securityEnabled,groupTypes,mail,mailNickname,onPremisesSyncEnabled,onPremisesSamAccountName,onPremisesSecurityIdentifier"
    $uri = "https://graph.microsoft.com/v1.0/users/$encodedUser/memberOf/microsoft.graph.group?`$select=$select"
    $groups = @()

    while (-not [string]::IsNullOrWhiteSpace($uri)) {
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $groups += @($result.value | Where-Object { $_ })
        $uri = [string]$result.'@odata.nextLink'
    }

    return @(Merge-CloudGroupSessionState -CloudUser $CloudUser -Groups $groups)
}

function Get-CloudGroupManagementType {
    param(
        [AllowNull()]
        $Group
    )

    if ($null -eq $Group) {
        return "Unknown"
    }

    if (@($Group.groupTypes) -contains "DynamicMembership") {
        return "SkipDynamic"
    }

    if ($Group.onPremisesSyncEnabled -eq $true) {
        return "AD"
    }

    if (@($Group.groupTypes) -contains "Unified") {
        return "Graph"
    }

    if ($Group.mailEnabled -eq $true) {
        return "Exchange"
    }

    return "Graph"
}

function Get-CloudGroupManagementLabel {
    param(
        [AllowNull()]
        $Group
    )

    switch (Get-CloudGroupManagementType -Group $Group) {
        "Graph"       { return "Graph" }
        "Exchange"    { return "Exchange Online" }
        "AD"          { return "Active Directory" }
        "SkipDynamic" { return "Skip - dynamic" }
        default       { return "Unknown" }
    }
}

function Get-CloudGroupTypeLabel {
    param(
        [AllowNull()]
        $Group
    )

    if ($null -eq $Group) {
        return "Unknown"
    }

    $recipientTypeDetails = ""
    if ($Group.PSObject.Properties.Match("RecipientTypeDetails").Count -gt 0) {
        $recipientTypeDetails = [string]$Group.RecipientTypeDetails
    }
    elseif ($Group.PSObject.Properties.Match("recipientTypeDetails").Count -gt 0) {
        $recipientTypeDetails = [string]$Group.recipientTypeDetails
    }

    switch -Regex ($recipientTypeDetails) {
        "^DynamicDistributionGroup$"      { return "Dynamic distribution list" }
        "^MailUniversalSecurityGroup$"   { return "Mail-enabled security group" }
        "^Mail.*Distribution"            { return "Distribution list" }
        "^RoomList$"                     { return "Room list" }
        "^(GroupMailbox|Microsoft365)"   { return "M365 group" }
    }

    $groupTypes = @($Group.groupTypes | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $isUnified = $groupTypes -contains "Unified"
    $isDynamic = $groupTypes -contains "DynamicMembership"
    $isMailEnabled = $Group.mailEnabled -eq $true
    $isSecurityEnabled = $Group.securityEnabled -eq $true

    if ($isUnified -and $isDynamic) {
        return "Dynamic M365 group"
    }
    if ($isUnified) {
        return "M365 group"
    }
    if ($isDynamic -and $isSecurityEnabled) {
        return "Dynamic security group"
    }
    if ($isDynamic) {
        return "Dynamic group"
    }
    if ($isMailEnabled -and $isSecurityEnabled) {
        return "Mail-enabled security group"
    }
    if ($isMailEnabled) {
        return "Distribution list"
    }
    if ($isSecurityEnabled) {
        return "Security group"
    }

    if ($groupTypes.Count -gt 0) {
        return ($groupTypes -join ",")
    }

    return "Group"
}

function New-CloudGroupRows {
    param(
        [AllowNull()]
        [object[]]$Groups
    )

    $rows = @()
    $cleanGroups = @($Groups | Where-Object { $_ })
    for ($i = 0; $i -lt $cleanGroups.Count; $i++) {
        $directorySource = "Cloud-only"
        if ($cleanGroups[$i].onPremisesSyncEnabled -eq $true) {
            $directorySource = "On-prem synced"
        }

        $rows += [pscustomobject]@{
            Number          = $i + 1
            DisplayName     = $cleanGroups[$i].displayName
            Management      = Get-CloudGroupManagementLabel -Group $cleanGroups[$i]
            DirectorySource = $directorySource
            MailEnabled     = $cleanGroups[$i].mailEnabled
            SecurityEnabled = $cleanGroups[$i].securityEnabled
            GroupTypes      = Get-CloudGroupTypeLabel -Group $cleanGroups[$i]
        }
    }

    return $rows
}

function Get-GroupManagerCloudMemberIdentity {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser
    )

    foreach ($propertyName in @("userPrincipalName", "mail", "id")) {
        $value = [string]$CloudUser.$propertyName
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    return ""
}

function Get-ExchangeCloudGroupIdentity {
    param(
        [Parameter(Mandatory = $true)]
        $Group
    )

    foreach ($propertyName in @("mail", "mailNickname", "displayName", "id")) {
        $value = [string]$Group.$propertyName
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    return ""
}

function Find-CloudGroupsBySearchTerm {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm
    )

    $select = "id,displayName,mailEnabled,securityEnabled,groupTypes,mail,mailNickname,onPremisesSyncEnabled,onPremisesSamAccountName,onPremisesSecurityIdentifier"
    $escapedTerm = Escape-GraphFilterValue -Value $SearchTerm
    $escapedSearchTerm = $SearchTerm.Replace('"', '\"')
    $candidateMatches = @()

    $filters = @(
        "startswith(displayName,'$escapedTerm') or startswith(mailNickname,'$escapedTerm') or mail eq '$escapedTerm'",
        "contains(displayName,'$escapedTerm') or contains(mailNickname,'$escapedTerm') or mail eq '$escapedTerm'"
    )

    foreach ($filter in $filters) {
        try {
            $encodedFilter = [System.Uri]::EscapeDataString($filter)
            $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=$encodedFilter" + [char]38 + "`$count=true" + [char]38 + "`$select=$select"
            $result = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{ ConsistencyLevel = "eventual" } -ErrorAction Stop
            $candidateMatches += @($result.value | Where-Object { $_ })
        }
        catch {
            Write-Verbose "An optional Graph group-filter strategy failed: $($_.Exception.Message)"
        }
    }

    try {
        $search = [System.Uri]::EscapeDataString('"displayName:{0}"' -f $escapedSearchTerm)
        $uri = "https://graph.microsoft.com/v1.0/groups?`$search=$search" + [char]38 + "`$count=true" + [char]38 + "`$select=$select"
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{ ConsistencyLevel = "eventual" } -ErrorAction Stop
        $candidateMatches += @($result.value | Where-Object { $_ })
    }
    catch {
        Write-Verbose "The optional Graph full-text group search failed: $($_.Exception.Message)"
    }

    return @(Get-UniqueGraphObjects -Objects $candidateMatches | Sort-Object displayName)
}

function Test-GroupMembershipAlreadyExistsError {
    param(
        [AllowNull()]
        [string]$ErrorMessage
    )

    if ([string]::IsNullOrWhiteSpace($ErrorMessage)) {
        return $false
    }

    return ($ErrorMessage -match "already.*member" -or $ErrorMessage -match "object references already exist")
}

function Test-GroupMembershipNotFoundError {
    param(
        [AllowNull()]
        [string]$ErrorMessage
    )

    if ([string]::IsNullOrWhiteSpace($ErrorMessage)) {
        return $false
    }

    return ($ErrorMessage -match "not.*member" -or $ErrorMessage -match "could not be found" -or $ErrorMessage -match "Request_ResourceNotFound")
}

function Invoke-GroupManagerCloudGroupMembershipChange {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser,

        [Parameter(Mandatory = $true)]
        $Group,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Add", "Remove")]
        [string]$Action
    )

    $managementType = Get-CloudGroupManagementType -Group $Group
    if ($managementType -eq "SkipDynamic") {
        Write-UiStatus -Status "SKIP" -Message "Skipped dynamic cloud group '$($Group.displayName)'." -Color Yellow
        return
    }

    if ($managementType -eq "AD") {
        Write-UiStatus -Status "INFO" -Message "Cannot manage on-prem synced group '$($Group.displayName)' from this standalone group manager. Manage it in Active Directory." -Color Yellow
        return
    }

    if ($managementType -eq "Exchange") {
        $groupIdentity = Get-ExchangeCloudGroupIdentity -Group $Group
        $memberIdentity = Get-GroupManagerCloudMemberIdentity -CloudUser $CloudUser
        if ([string]::IsNullOrWhiteSpace($groupIdentity) -or [string]::IsNullOrWhiteSpace($memberIdentity)) {
            Write-UiStatus -Status "WARN" -Message "Could not resolve Exchange Online group/member identity for '$($Group.displayName)'." -Color Yellow
            return
        }

        $bypassPreview = if ($script:BypassGroupOwnershipCheck) { " -BypassSecurityGroupManagerCheck" } else { "" }
        $commandPreview = if ($Action -eq "Add") {
            "Add-DistributionGroupMember -Identity '$groupIdentity' -Member '$memberIdentity'$bypassPreview"
        }
        else {
            "Remove-DistributionGroupMember -Identity '$groupIdentity' -Member '$memberIdentity'$bypassPreview -Confirm:`$false"
        }

        if ($script:DryRun) {
            Write-UiDryRunCommand -Name "$Action Exchange-managed group membership" -CommandPreview $commandPreview
            Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action $Action
            return
        }

        try {
            Ensure-GroupManagerExchangeOnlineConnected
            $membershipParameters = @{
                Identity    = $groupIdentity
                Member      = $memberIdentity
                ErrorAction = "Stop"
            }
            if ($script:BypassGroupOwnershipCheck) {
                $membershipParameters.BypassSecurityGroupManagerCheck = $true
            }

            if ($Action -eq "Add") {
                Add-DistributionGroupMember @membershipParameters
                Write-UiStatus -Status "OK" -Message "Added '$memberIdentity' to Exchange-managed group '$($Group.displayName)'." -Color Green
                Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action "Add"
            }
            else {
                $membershipParameters.Confirm = $false
                Remove-DistributionGroupMember @membershipParameters
                Write-UiStatus -Status "OK" -Message "Removed '$memberIdentity' from Exchange-managed group '$($Group.displayName)'." -Color Green
                Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action "Remove"
            }
        }
        catch {
            $errorMessage = [string]$_.Exception.Message
            if ($Action -eq "Add" -and (Test-GroupMembershipAlreadyExistsError -ErrorMessage $errorMessage)) {
                Write-UiStatus -Status "SKIP" -Message "'$memberIdentity' is already in Exchange-managed group '$($Group.displayName)'. No change needed." -Color Yellow
                Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action "Add"
                return
            }
            if ($Action -eq "Remove" -and (Test-GroupMembershipNotFoundError -ErrorMessage $errorMessage)) {
                Write-UiStatus -Status "SKIP" -Message "'$memberIdentity' was not found in Exchange-managed group '$($Group.displayName)'. No change needed." -Color Yellow
                Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action "Remove"
                return
            }

            Write-UiStatus -Status "WARN" -Message "Could not $($Action.ToLowerInvariant()) Exchange-managed group '$($Group.displayName)': $errorMessage" -Color Yellow
        }
        return
    }

    try {
        $encodedGroup = [System.Uri]::EscapeDataString($Group.id)
        $encodedUser = [System.Uri]::EscapeDataString($CloudUser.id)
        if ($Action -eq "Add") {
            $body = @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($CloudUser.id)"
            } | ConvertTo-Json -Depth 4
            if ($script:DryRun) {
                Write-UiDryRunCommand -Name "Add Graph-managed group membership" -CommandPreview "Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/groups/$encodedGroup/members/`$ref' -Body <redacted>"
                Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action "Add"
                return
            }
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$encodedGroup/members/`$ref" -Body $body -ContentType "application/json" -ErrorAction Stop
            Write-UiStatus -Status "OK" -Message "Added '$($CloudUser.userPrincipalName)' to Graph-managed group '$($Group.displayName)'." -Color Green
            Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action "Add"
        }
        else {
            if ($script:DryRun) {
                Write-UiDryRunCommand -Name "Remove Graph-managed group membership" -CommandPreview "Invoke-MgGraphRequest -Method DELETE -Uri 'https://graph.microsoft.com/v1.0/groups/$encodedGroup/members/$encodedUser/`$ref'"
                Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action "Remove"
                return
            }
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$encodedGroup/members/$encodedUser/`$ref" -ErrorAction Stop
            Write-UiStatus -Status "OK" -Message "Removed '$($CloudUser.userPrincipalName)' from Graph-managed group '$($Group.displayName)'." -Color Green
            Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action "Remove"
        }
    }
    catch {
        $errorMessage = [string]$_.Exception.Message
        if ($Action -eq "Add" -and (Test-GroupMembershipAlreadyExistsError -ErrorMessage $errorMessage)) {
            Write-UiStatus -Status "SKIP" -Message "'$($CloudUser.userPrincipalName)' is already in Graph-managed group '$($Group.displayName)'. No change needed." -Color Yellow
            Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action "Add"
            return
        }
        if ($Action -eq "Remove" -and (Test-GroupMembershipNotFoundError -ErrorMessage $errorMessage)) {
            Write-UiStatus -Status "SKIP" -Message "'$($CloudUser.userPrincipalName)' was not found in Graph-managed group '$($Group.displayName)'. No change needed." -Color Yellow
            Update-CloudGroupSessionState -CloudUser $CloudUser -Group $Group -Action "Remove"
            return
        }

        Write-UiStatus -Status "WARN" -Message "Could not $($Action.ToLowerInvariant()) Graph-managed group '$($Group.displayName)': $errorMessage" -Color Yellow
    }
}

function Add-CloudGroupsToTargetUser {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser,

        [AllowNull()]
        [object[]]$Groups
    )

    $groupsToAdd = @($Groups | Where-Object { $_ })
    if ($groupsToAdd.Count -eq 0) {
        Write-UiStatus -Status "INFO" -Message "No cloud groups selected to add." -Color Yellow
        return
    }

    foreach ($group in $groupsToAdd) {
        Invoke-GroupManagerCloudGroupMembershipChange -CloudUser $CloudUser -Group $group -Action "Add"
    }
}

function Remove-CloudGroupsFromTargetUser {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser,

        [AllowNull()]
        [object[]]$Groups
    )

    $groupsToRemove = @($Groups | Where-Object { $_ })
    if ($groupsToRemove.Count -eq 0) {
        Write-UiStatus -Status "INFO" -Message "No cloud groups selected to remove." -Color Yellow
        return
    }

    foreach ($group in $groupsToRemove) {
        Invoke-GroupManagerCloudGroupMembershipChange -CloudUser $CloudUser -Group $group -Action "Remove"
    }
}

function Invoke-GroupManagerMembershipReview {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser
    )

    while ($true) {
        try {
            $currentGroups = @(Get-CloudUserGroups -CloudUser $CloudUser)
        }
        catch {
            Write-UiStatus -Status "WARN" -Message "Could not load current cloud groups: $(Get-MgGraphExceptionText -ErrorRecord $_)" -Color Yellow
            $currentGroups = @()
        }

        Write-UiBox -Title "Current Cloud Groups" -Lines (ConvertTo-UiTableLines -Rows (New-CloudGroupRows -Groups $currentGroups) -Columns @("Number", "DisplayName", "Management", "DirectorySource", "GroupTypes"))
        $choice = Read-UiInput -Prompt "Adjust cloud groups?" -Options @("a=add", "r=remove", "c=continue")
        if ($choice -notin @("a", "r", "c")) {
            Write-UiStatus -Status "INFO" -Message "Enter exact lowercase a, r, or c." -Color Yellow
            continue
        }
        if ($choice -eq "c") {
            return
        }
        if ($choice -eq "r") {
            $selection = Read-UiInput -Prompt "Enter cloud group numbers to remove" -Options @("comma separated", "blank=back")
            if ([string]::IsNullOrWhiteSpace($selection)) {
                continue
            }

            $numbers = ConvertTo-NumberSelection -InputText $selection -Max $currentGroups.Count
            if ($numbers.Count -eq 0) {
                Write-UiStatus -Status "INFO" -Message "No valid cloud group numbers selected to remove." -Color Yellow
                continue
            }

            $selectedGroups = @(Select-ItemsByNumber -Items $currentGroups -Numbers $numbers)
            Write-UiBox -Title "Confirm Group Removal" -Lines (ConvertTo-UiTableLines -Rows (New-CloudGroupRows -Groups $selectedGroups) -Columns @("DisplayName", "Management", "GroupTypes")) -Color Yellow
            if (-not (Read-UiYesNo -Prompt "Remove the selected group membership(s)?" -DefaultYes $false)) {
                Write-UiStatus -Status "SKIP" -Message "No group memberships were removed." -Color Yellow
                continue
            }

            Remove-CloudGroupsFromTargetUser -CloudUser $CloudUser -Groups $selectedGroups
            continue
        }
        if ($choice -eq "a") {
            $term = Read-UiInput -Prompt "Search cloud groups by name" -Options @("text", "blank=back")
            if ([string]::IsNullOrWhiteSpace($term)) {
                continue
            }

            try {
                $groupMatches = @(Find-CloudGroupsBySearchTerm -SearchTerm $term)
            }
            catch {
                Write-UiStatus -Status "WARN" -Message "Could not search cloud groups: $(Get-MgGraphExceptionText -ErrorRecord $_)" -Color Yellow
                continue
            }

            if ($groupMatches.Count -eq 0) {
                Write-UiStatus -Status "WARN" -Message "No cloud groups matched '$term'." -Color Yellow
                continue
            }

            Write-UiBox -Title "Cloud Group Search Results" -Lines (ConvertTo-UiTableLines -Rows (New-CloudGroupRows -Groups $groupMatches) -Columns @("Number", "DisplayName", "Management", "DirectorySource", "GroupTypes"))
            $selection = Read-UiInput -Prompt "Enter cloud group numbers to add" -Options @("comma separated", "blank=back")
            if ([string]::IsNullOrWhiteSpace($selection)) {
                continue
            }

            $numbers = ConvertTo-NumberSelection -InputText $selection -Max $groupMatches.Count
            if ($numbers.Count -eq 0) {
                Write-UiStatus -Status "INFO" -Message "No valid cloud group numbers selected to add." -Color Yellow
                continue
            }

            $selectedGroups = @(Select-ItemsByNumber -Items $groupMatches -Numbers $numbers)
            Write-UiBox -Title "Confirm Group Addition" -Lines (ConvertTo-UiTableLines -Rows (New-CloudGroupRows -Groups $selectedGroups) -Columns @("DisplayName", "Management", "GroupTypes")) -Color Yellow
            if (-not (Read-UiYesNo -Prompt "Add the selected group membership(s)?" -DefaultYes $false)) {
                Write-UiStatus -Status "SKIP" -Message "No group memberships were added." -Color Yellow
                continue
            }

            Add-CloudGroupsToTargetUser -CloudUser $CloudUser -Groups $selectedGroups
        }
    }
}

function Invoke-StandaloneGroupManager {
    param(
        [string]$UserPrincipalName,

        [switch]$SkipConnect
    )

    try {
        if (-not $SkipConnect) {
            Connect-GroupManagerServices
        }

        $targetUser = Resolve-GroupManagerCloudUser -Lookup $UserPrincipalName
        if (-not $targetUser) {
            Write-UiStatus -Status "INFO" -Message "No target user selected. Exiting group manager." -Color Yellow
            return
        }

        Invoke-GroupManagerMembershipReview -CloudUser $targetUser
    }
    catch {
        Write-UiStatus -Status "ERROR" -Message "Group manager failed: $($_.Exception.Message)" -Color Red
    }
}

if ($MyInvocation.InvocationName -ne ".") {
    Invoke-StandaloneGroupManager -UserPrincipalName $UserPrincipalName -SkipConnect:$SkipConnect
}
