#Requires -Version 7.2

<#
.SYNOPSIS
Interactively reviews and changes Microsoft 365 user license assignments.

.DESCRIPTION
Uses delegated Microsoft Graph authentication and least-privilege scopes for
user lookup, usage-location updates, SKU inventory, and license assignment.
Mutations require confirmation. Use -DryRun or -WhatIf to preview changes.

.EXAMPLE
./licensemanager.ps1 -UserPrincipalName alex.wilber@contoso.com -DryRun
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
param(
    [string]$UserPrincipalName,

    [switch]$SkipConnect,

    [switch]$DryRun
)

$toolkitModulePath = Join-Path $PSScriptRoot "..\src\M365AdminToolkit\M365AdminToolkit.psd1"
Import-Module $toolkitModulePath -Force -ErrorAction Stop
Set-M365ToolkitUiWidth -Width 120

$script:DryRun = [bool]($DryRun -or $WhatIfPreference)

function Get-LicenseManagerGraphScopes {
    return @(
        "User.Read.All",
        "User.ReadUpdate.All",
        "LicenseAssignment.Read.All",
        "LicenseAssignment.ReadWrite.All"
    )
}

function Connect-LicenseManagerGraph {
    Import-RequiredPowerShellModule -ModuleName "Microsoft.Graph.Authentication"

    Write-UiStatus -Status "LOADING..." -Message "Connecting to Microsoft Graph." -Color Cyan
    try {
        Connect-MgGraph -NoWelcome -ContextScope Process -Scopes (Get-LicenseManagerGraphScopes)
    }
    catch {
        if ([string]$_.Exception.Message -notmatch "ContextScope") {
            throw
        }

        Write-UiStatus -Status "WARN" -Message "This Microsoft.Graph.Authentication version does not support -ContextScope. Connecting without process-scoped context." -Color Yellow
        Connect-MgGraph -NoWelcome -Scopes (Get-LicenseManagerGraphScopes)
    }

    Write-UiStatus -Status "OK" -Message "Connected to Microsoft Graph." -Color Green
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

function Invoke-LicenseManagerGraphCollectionRequest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $items = @()
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $result = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop
        $items += @($result.value | Where-Object { $_ })
        $nextUri = [string]$result.'@odata.nextLink'
    }

    return $items
}

function Escape-GraphFilterValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Replace("'", "''")
}

function Get-GraphUserUri {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdOrUpn,

        [string]$Select
    )

    $encodedUser = [System.Uri]::EscapeDataString($UserIdOrUpn)
    $uri = "https://graph.microsoft.com/v1.0/users/$encodedUser"
    if (-not [string]::IsNullOrWhiteSpace($Select)) {
        $uri = "$uri`?`$select=$Select"
    }

    return $uri
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

function Get-UniqueCloudUsers {
    param(
        [AllowNull()]
        [object[]]$Users
    )

    $uniqueUsers = @()
    $seenIds = @{}
    foreach ($user in @($Users | Where-Object { $_ })) {
        $key = [string]$user.id
        if ([string]::IsNullOrWhiteSpace($key)) {
            $key = [string]$user.userPrincipalName
        }
        if ([string]::IsNullOrWhiteSpace($key) -or $seenIds.ContainsKey($key)) {
            continue
        }

        $seenIds[$key] = $true
        $uniqueUsers += $user
    }

    return $uniqueUsers
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

    return Get-UniqueCloudUsers -Users $candidateMatches
}

function Show-LicenseManagerCloudUserMatches {
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

function Resolve-LicenseManagerCloudUser {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Lookup
    )

    $candidateLookup = $Lookup
    $select = "id,displayName,userPrincipalName,mail,mailNickname,givenName,surname,accountEnabled,usageLocation,assignedLicenses,onPremisesSamAccountName,proxyAddresses"
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($candidateLookup)) {
            $candidateLookup = Read-UiInput -Prompt "Cloud user lookup" -Options @("exact UPN/email/name", "x=exit")
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
            Show-LicenseManagerCloudUserMatches -Matches $candidateMatches
        }

        $candidateLookup = Read-UiInput -Prompt "Enter a different cloud user lookup" -Options @("exact UPN/email/name", "x=exit")
        if ($candidateLookup -eq "x") {
            return $null
        }
    }
}

function Get-SkuCommonNames {
    return @{
        "AAD_PREMIUM" = "Microsoft Entra ID P1"
        "AAD_PREMIUM_P2" = "Microsoft Entra ID P2"
        "ATP_ENTERPRISE" = "Microsoft Defender for Office 365 (Plan 1)"
        "ATP_ENTERPRISE_FACULTY" = "Microsoft Defender for Office 365 (Plan 1) Faculty"
        "ATP_ENTERPRISE_GOV" = "Microsoft Defender for Office 365 (Plan 1) GCC"
        "ATP_ENTERPRISE_STUDENT" = "Microsoft Defender for Office 365 (Plan 1) Student"
        "DESKLESSPACK" = "Office 365 F3"
        "EMS" = "Enterprise Mobility + Security E3"
        "EMSPREMIUM" = "Enterprise Mobility + Security E5"
        "ENTERPRISEPACK" = "Office 365 E3"
        "ENTERPRISEPREMIUM" = "Office 365 E5"
        "EXCHANGESTANDARD" = "Exchange Online (Plan 1)"
        "EXCHANGEENTERPRISE" = "Exchange Online (Plan 2)"
        "FLOW_FREE" = "Microsoft Power Automate Free"
        "M365EDU_A3_FACULTY" = "Microsoft 365 A3 for Faculty"
        "M365EDU_A5_FACULTY" = "Microsoft 365 A5 for Faculty"
        "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
        "O365_BUSINESS_PREMIUM" = "Microsoft 365 Business Standard"
        "PBI_PREMIUM_PER_USER" = "Power BI Premium Per User"
        "PBI_PREMIUM_PER_USER_ADDON" = "Power BI Premium Per User Add-On"
        "POWER_BI_ADDON" = "Power BI for Office 365 Add-On"
        "POWER_BI_INDIVIDUAL_USER" = "Power BI"
        "POWER_BI_PRO" = "Power BI Pro"
        "POWER_BI_STANDARD" = "Microsoft Fabric Free"
        "POWER_BI_STANDARD_FACULTY" = "Microsoft Fabric Free Faculty"
        "POWER_BI_STANDARD_STUDENT" = "Microsoft Fabric Free Student"
        "POWER_BI_STANDARD_GOV" = "Microsoft Fabric Free GCC"
        "POWER_BI_STANDARD_DEPT" = "Microsoft Fabric Free Department"
        "POWER_BI_STANDARD_INF" = "Microsoft Fabric Free"
        "SMB_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
        "SMB_BUSINESS_PREMIUM" = "Microsoft 365 Business Standard - Prepaid Legacy"
        "SPE_E3" = "Microsoft 365 E3"
        "SPE_E5" = "Microsoft 365 E5"
        "SPE_F1" = "Microsoft 365 F3"
        "SPB" = "Microsoft 365 Business Premium"
        "SPZA_IW" = "App Connect IW"
        "STANDARDPACK" = "Office 365 E1"
        "VISIOCLIENT" = "Visio Plan 2"
        "WINDOWS_STORE" = "Windows Store for Business"
    }
}

function New-SkuLookup {
    param(
        [AllowNull()]
        [object[]]$Skus,

        [Parameter(Mandatory = $true)]
        [hashtable]$SkuCommonNames
    )

    $skuById = @{}
    foreach ($sku in @($Skus)) {
        $commonName = $SkuCommonNames[$sku.skuPartNumber]
        if ([string]::IsNullOrWhiteSpace($commonName)) {
            $commonName = $sku.skuPartNumber
        }

        $skuById[[string]$sku.skuId] = [pscustomobject]@{
            CommonName    = $commonName
            SkuPartNumber = $sku.skuPartNumber
        }
    }

    return $skuById
}

function Get-AssignedLicenseRows {
    param(
        [AllowNull()]
        [object[]]$AssignedLicenses,

        [Parameter(Mandatory = $true)]
        [hashtable]$SkuById
    )

    return @(
        @($AssignedLicenses | Where-Object { $_ }) |
            ForEach-Object {
                $license = $SkuById[[string]$_.skuId]
                if ($license) {
                    [pscustomobject]@{
                        License       = $license.CommonName
                        SkuPartNumber = $license.SkuPartNumber
                        SkuId         = $_.skuId
                    }
                }
                else {
                    [pscustomobject]@{
                        License       = "Unknown SKU"
                        SkuPartNumber = ""
                        SkuId         = $_.skuId
                    }
                }
            } |
            Sort-Object License
    )
}

function Get-TenantLicenseRows {
    param(
        [AllowNull()]
        [object[]]$Skus,

        [Parameter(Mandatory = $true)]
        [hashtable]$SkuCommonNames
    )

    return @(
        @($Skus) |
            ForEach-Object {
                $commonName = $SkuCommonNames[$_.skuPartNumber]
                if ([string]::IsNullOrWhiteSpace($commonName)) {
                    $commonName = $_.skuPartNumber
                }

                $graphEnabled = [int]$_.prepaidUnits.enabled
                $warning = [int]$_.prepaidUnits.warning
                $suspended = [int]$_.prepaidUnits.suspended
                $lockedOut = [int]$_.prepaidUnits.lockedOut
                $consumed = [int]$_.consumedUnits
                $effectiveEnabled = $graphEnabled + $warning
                $effectiveAvailable = $effectiveEnabled - $consumed

                [pscustomobject]@{
                    License              = $commonName
                    SkuPartNumber        = $_.skuPartNumber
                    SkuId                = $_.skuId
                    AppliesTo            = $_.appliesTo
                    CapabilityStatus     = $_.capabilityStatus
                    Enabled              = $effectiveEnabled
                    GraphEnabled         = $graphEnabled
                    Warning              = $warning
                    Suspended            = $suspended
                    LockedOut            = $lockedOut
                    Consumed             = $consumed
                    AvailableEnabled     = $effectiveAvailable
                    AvailableWithWarning = $effectiveAvailable
                    Available            = $effectiveAvailable
                }
            } |
            Sort-Object License
    )
}

function ConvertTo-LicenseNumberSelection {
    param(
        [string]$InputText,
        [int]$Max
    )

    if ([string]::IsNullOrWhiteSpace($InputText)) {
        return @()
    }

    $numbers = @()
    $parts = @($InputText -split "[,\s;]+" | Where-Object { $_ })
    if ($parts.Count -eq 0) {
        return @()
    }

    foreach ($part in $parts) {
        if ($part -notmatch "^\d+$") {
            return @()
        }

        $number = [int]$part
        if ($number -lt 1 -or $number -gt $Max) {
            return @()
        }

        if ($numbers -notcontains $number) {
            $numbers += $number
        }
    }

    return $numbers
}

function Get-CloudUserLicenseSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdOrUpn
    )

    $skuCommonNames = Get-SkuCommonNames
    $skuUri = "https://graph.microsoft.com/v1.0/subscribedSkus?`$select=skuId,skuPartNumber,appliesTo,capabilityStatus,prepaidUnits,consumedUnits"
    try {
        $skuRows = @(Invoke-LicenseManagerGraphCollectionRequest -Uri $skuUri)
    }
    catch {
        throw "Could not load tenant subscribed SKUs from Graph. URI: $skuUri. Error: $(Get-MgGraphExceptionText -ErrorRecord $_)"
    }
    $skuById = New-SkuLookup -Skus $skuRows -SkuCommonNames $skuCommonNames

    $userSelect = "id,displayName,userPrincipalName,usageLocation,assignedLicenses"
    $userUri = Get-GraphUserUri -UserIdOrUpn $UserIdOrUpn -Select $userSelect
    try {
        $cloudUser = Invoke-MgGraphRequest -Method GET -Uri $userUri
    }
    catch {
        throw "Could not load assigned licenses for cloud user '$UserIdOrUpn'. URI: $userUri. Error: $(Get-MgGraphExceptionText -ErrorRecord $_)"
    }

    $assignedRows = @(Get-AssignedLicenseRows -AssignedLicenses $cloudUser.assignedLicenses -SkuById $skuById)
    $tenantRows = @(Get-TenantLicenseRows -Skus $skuRows -SkuCommonNames $skuCommonNames)

    return [pscustomobject]@{
        CloudUser     = $cloudUser
        SkuById       = $skuById
        AssignedRows  = $assignedRows
        TenantRows    = $tenantRows
        AvailableRows = @($tenantRows | Where-Object { $_ -and $_.Available -gt 0 })
    }
}

function New-LicenseManagerNumberedRows {
    param(
        [AllowNull()]
        [object[]]$Rows
    )

    $numberedRows = @()
    $cleanRows = @($Rows | Where-Object { $_ })
    for ($i = 0; $i -lt $cleanRows.Count; $i++) {
        $row = $cleanRows[$i]
        $numberedRows += [pscustomobject]@{
            Number               = $i + 1
            License               = $row.License
            SkuPartNumber         = $row.SkuPartNumber
            AppliesTo             = $row.AppliesTo
            Status                = $row.CapabilityStatus
            Enabled               = $row.Enabled
            GraphEnabled          = $row.GraphEnabled
            Warning               = $row.Warning
            Warn                  = $row.Warning
            Suspended             = $row.Suspended
            LockedOut             = $row.LockedOut
            Consumed              = $row.Consumed
            Used                  = $row.Consumed
            Available             = $row.Available
            Avail                 = $row.Available
            AvailableWithWarning  = $row.AvailableWithWarning
            AvailWarn             = $row.AvailableWithWarning
            SkuId                 = $row.SkuId
        }
    }

    return $numberedRows
}

function ConvertTo-LicenseDetailValue {
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return "(blank)"
    }

    return [string]$Value
}

function Write-LicenseInventoryDetails {
    param(
        [Parameter(Mandatory = $true)]
        $LicenseRow
    )

    Write-UiBox -Title "License Details #$($LicenseRow.Number)" -Lines @(
        New-UiBoxLine -Label "License" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.License)
        New-UiBoxLine -Label "SkuPartNumber" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.SkuPartNumber)
        New-UiBoxLine -Label "SkuId" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.SkuId)
        New-UiBoxLine -Label "AppliesTo" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.AppliesTo)
        New-UiBoxLine -Label "Status" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.Status)
        New-UiBoxLine -Label "Enabled" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.Enabled)
        New-UiBoxLine -Label "GraphEnabled" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.GraphEnabled)
        New-UiBoxLine -Label "Warning" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.Warning)
        New-UiBoxLine -Label "Suspended" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.Suspended)
        New-UiBoxLine -Label "LockedOut" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.LockedOut)
        New-UiBoxLine -Label "Consumed" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.Consumed)
        New-UiBoxLine -Label "Available" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.Available)
        New-UiBoxLine -Label "AvailableWithWarning" -Value (ConvertTo-LicenseDetailValue -Value $LicenseRow.AvailableWithWarning)
    )
}

function Write-LicenseManagerSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        $Snapshot
    )

    $usageLocation = [string]$Snapshot.CloudUser.usageLocation
    if ([string]::IsNullOrWhiteSpace($usageLocation)) {
        $usageLocation = "(blank)"
    }

    Write-UiBox -Title "License Target" -Lines @(
        New-UiBoxLine -Label "DisplayName" -Value $Snapshot.CloudUser.displayName
        New-UiBoxLine -Label "UserPrincipalName" -Value $Snapshot.CloudUser.userPrincipalName
        New-UiBoxLine -Label "UsageLocation" -Value $usageLocation
        New-UiBoxLine -Label "Tenant SKU count" -Value @($Snapshot.TenantRows).Count
        New-UiBoxLine -Label "Assignable SKU count" -Value @($Snapshot.AvailableRows).Count
    )

    $tenantRows = @(New-LicenseManagerNumberedRows -Rows $Snapshot.TenantRows)
    if ($tenantRows.Count -eq 0) {
        Write-UiBox -Title "Tenant License Inventory" -Lines @("None")
    }
    else {
        Write-UiBox -Title "Tenant License Inventory" -Lines (ConvertTo-UiTableLines -Rows $tenantRows -Columns @("Number", "License", "Enabled", "Consumed", "Available"))
    }

    $availableRows = @(New-LicenseManagerNumberedRows -Rows $Snapshot.AvailableRows)
    if ($availableRows.Count -eq 0) {
        Write-UiBox -Title "Available Licenses To Add" -Lines @("None")
    }
    else {
        Write-UiBox -Title "Available Licenses To Add" -Lines (ConvertTo-UiTableLines -Rows $availableRows -Columns @("Number", "License", "Enabled", "Available"))
    }

    $assignedRows = @(New-LicenseManagerNumberedRows -Rows $Snapshot.AssignedRows)
    if ($assignedRows.Count -eq 0) {
        Write-UiBox -Title "Assigned Licenses" -Lines @("None")
    }
    else {
        Write-UiBox -Title "Assigned Licenses" -Lines (ConvertTo-UiTableLines -Rows $assignedRows -Columns @("Number", "License"))
    }
}

function Test-LicenseUsageLocation {
    param(
        [AllowNull()]
        [string]$UsageLocation
    )

    return (-not [string]::IsNullOrWhiteSpace($UsageLocation) -and $UsageLocation.Trim() -match "^[A-Za-z]{2}$")
}

function Set-CloudUserUsageLocationForLicensing {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdOrUpn,

        [Parameter(Mandatory = $true)]
        [string]$UsageLocation
    )

    $body = @{
        usageLocation = $UsageLocation.Trim().ToUpperInvariant()
    } | ConvertTo-Json -Depth 4

    return Invoke-MgGraphRequest -Method PATCH -Uri (Get-GraphUserUri -UserIdOrUpn $UserIdOrUpn) -Body $body -ContentType "application/json"
}

function Set-CloudUserObjectUsageLocationValue {
    param(
        [Parameter(Mandatory = $true)]
        $CloudUser,

        [Parameter(Mandatory = $true)]
        [string]$UsageLocation
    )

    try {
        $CloudUser.usageLocation = $UsageLocation
    }
    catch {
        try {
            $CloudUser | Add-Member -MemberType NoteProperty -Name usageLocation -Value $UsageLocation -Force
        }
        catch {
            Write-Verbose "Could not update the in-memory usageLocation property: $($_.Exception.Message)"
        }
    }
}

function Ensure-CloudUserUsageLocationForLicenseAssignment {
    param(
        [Parameter(Mandatory = $true)]
        $Snapshot
    )

    if (Test-LicenseUsageLocation -UsageLocation $Snapshot.CloudUser.usageLocation) {
        return $true
    }

    $location = "US"
    try {
        Invoke-UiCommand `
            -Name "Set cloud user usage location" `
            -CommandPreview "Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/v1.0/users/$($Snapshot.CloudUser.id)' -Body '{ usageLocation = $location }'" `
            -Command {
                Set-CloudUserUsageLocationForLicensing -UserIdOrUpn $Snapshot.CloudUser.id -UsageLocation $location | Out-Null
            }
        Set-CloudUserObjectUsageLocationValue -CloudUser $Snapshot.CloudUser -UsageLocation $location
        if (-not (Test-UiDryRun)) {
            Write-UiStatus -Status "OK" -Message "Set UsageLocation to '$location' for '$($Snapshot.CloudUser.userPrincipalName)'." -Color Green
        }
        return $true
    }
    catch {
        Write-UiStatus -Status "WARN" -Message "Could not set UsageLocation: $(Get-MgGraphExceptionText -ErrorRecord $_)" -Color Yellow
        return $false
    }
}

function Invoke-CloudUserLicenseAssignment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdOrUpn,

        [AllowNull()]
        [object[]]$AddLicenseRows = @(),

        [AllowNull()]
        [object[]]$RemoveLicenseRows = @()
    )

    $addLicenses = @(
        @($AddLicenseRows | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_.SkuId) }) |
            ForEach-Object {
                @{
                    skuId         = $_.SkuId
                    disabledPlans = @()
                }
            }
    )

    $removeLicenses = @(
        @($RemoveLicenseRows | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_.SkuId) }) |
            ForEach-Object { [string]$_.SkuId }
    )

    $body = @{
        addLicenses    = @($addLicenses)
        removeLicenses = @($removeLicenses)
    } | ConvertTo-Json -Depth 8

    $encodedUser = [System.Uri]::EscapeDataString($UserIdOrUpn)
    try {
        return Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$encodedUser/assignLicense" -Body $body -ContentType "application/json"
    }
    catch {
        throw (Get-MgGraphExceptionText -ErrorRecord $_)
    }
}

function Read-LicenseManagerSelection {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [string]$Prompt
    )

    while ($true) {
        $selection = Read-UiInput -Prompt $Prompt -Options @("blank=back")
        if ([string]::IsNullOrWhiteSpace($selection)) {
            return @()
        }

        $numbers = ConvertTo-LicenseNumberSelection -InputText $selection -Max $Rows.Count
        if ($numbers.Count -eq 0) {
            Write-UiStatus -Status "INFO" -Message "Enter valid license numbers from the list." -Color Yellow
            continue
        }

        return @($numbers | ForEach-Object { $Rows[$_ - 1] })
    }
}

function Invoke-InteractiveLicenseManager {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdOrUpn,

        [string]$DisplayName = ""
    )

    if ([string]::IsNullOrWhiteSpace($DisplayName)) {
        $DisplayName = $UserIdOrUpn
    }

    while ($true) {
        Write-UiStatus -Status "LOADING..." -Message "Refreshing license state for '$DisplayName'." -Color Cyan
        try {
            $snapshot = Get-CloudUserLicenseSnapshot -UserIdOrUpn $UserIdOrUpn
        }
        catch {
            Write-UiStatus -Status "WARN" -Message $_.Exception.Message -Color Yellow
            $retry = Read-UiInput -Prompt "License state" -Options @("q=try again", "x=back")
            if ($retry -eq "q") {
                continue
            }
            return
        }

        Write-LicenseManagerSnapshot -Snapshot $snapshot

        $choice = Read-UiInput -Prompt "Adjust licenses?" -Options @("a=add licenses", "r=remove licenses", "d=license details", "q=query license state", "x=exit")
        if ($choice -eq "x") {
            return
        }
        if ($choice -eq "q") {
            continue
        }

        if ($choice -eq "d") {
            $tenantRows = @(New-LicenseManagerNumberedRows -Rows $snapshot.TenantRows)
            if ($tenantRows.Count -eq 0) {
                Write-UiStatus -Status "INFO" -Message "No tenant license rows are available for details." -Color Yellow
                continue
            }

            $selectedRows = @(Read-LicenseManagerSelection -Rows $tenantRows -Prompt "Select license inventory rows for details")
            foreach ($selectedRow in $selectedRows) {
                Write-LicenseInventoryDetails -LicenseRow $selectedRow
            }
            continue
        }

        if ($choice -eq "a") {
            $availableRows = @(New-LicenseManagerNumberedRows -Rows $snapshot.AvailableRows)
            if ($availableRows.Count -eq 0) {
                Write-UiStatus -Status "INFO" -Message "No available tenant licenses can be added." -Color Yellow
                continue
            }

            $selectedRows = @(Read-LicenseManagerSelection -Rows $availableRows -Prompt "Select available licenses to add")
            if ($selectedRows.Count -eq 0) {
                continue
            }

            Write-UiBox -Title "Selected Licenses To Add" -Lines (ConvertTo-UiTableLines -Rows $selectedRows -Columns @("License", "Available"))
            if (-not (Read-UiYesNo -Prompt "Assign selected licenses?" -DefaultYes $false)) {
                Write-UiStatus -Status "SKIP" -Message "Skipped selected license assignment." -Color Yellow
                continue
            }

            if (-not (Ensure-CloudUserUsageLocationForLicenseAssignment -Snapshot $snapshot)) {
                continue
            }

            try {
                $selectedLicenseText = (@($selectedRows | ForEach-Object { $_.License }) -join ", ")
                Invoke-UiCommand `
                    -Name "Assign selected cloud licenses" `
                    -CommandPreview "Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/users/$($snapshot.CloudUser.id)/assignLicense' -Body <add $selectedLicenseText>" `
                    -Command {
                        Invoke-CloudUserLicenseAssignment -UserIdOrUpn $snapshot.CloudUser.id -AddLicenseRows $selectedRows -RemoveLicenseRows @() | Out-Null
                    }
                if (-not (Test-UiDryRun)) {
                    Write-UiStatus -Status "OK" -Message "Assigned $($selectedRows.Count) license(s) to '$($snapshot.CloudUser.userPrincipalName)'." -Color Green
                }
            }
            catch {
                Write-UiStatus -Status "WARN" -Message "Could not assign selected license(s) to '$($snapshot.CloudUser.userPrincipalName)': $($_.Exception.Message)" -Color Yellow
                Write-UiStatus -Status "INFO" -Message "Check whether this SKU can be assigned by Graph, whether required service plans/dependencies apply, whether usage location is set, and whether group-based licensing is involved." -Color Yellow
            }
            continue
        }

        if ($choice -eq "r") {
            $assignedRows = @(New-LicenseManagerNumberedRows -Rows $snapshot.AssignedRows)
            if ($assignedRows.Count -eq 0) {
                Write-UiStatus -Status "INFO" -Message "No assigned licenses can be removed." -Color Yellow
                continue
            }

            $selectedRows = @(Read-LicenseManagerSelection -Rows $assignedRows -Prompt "Select assigned licenses to remove")
            if ($selectedRows.Count -eq 0) {
                continue
            }

            Write-UiBox -Title "Licenses To Remove" -Lines (ConvertTo-UiTableLines -Rows $selectedRows -Columns @("License"))
            $removalConfirmed = $false
            if ($selectedRows.Count -eq $assignedRows.Count) {
                Write-UiStatus -Status "WARN" -Message "This will leave '$($snapshot.CloudUser.userPrincipalName)' unlicensed." -Color Yellow
                if (-not (Read-UiYesNo -Prompt "Are you sure you want to remove all assigned licenses?" -DefaultYes $false)) {
                    Write-UiStatus -Status "SKIP" -Message "Skipped removing all assigned licenses." -Color Yellow
                    continue
                }
                $removalConfirmed = $true
            }

            if (-not $removalConfirmed -and -not (Read-UiYesNo -Prompt "Remove selected licenses?" -DefaultYes $false)) {
                Write-UiStatus -Status "SKIP" -Message "Skipped selected license removal." -Color Yellow
                continue
            }

            try {
                $selectedLicenseText = (@($selectedRows | ForEach-Object { $_.License }) -join ", ")
                Invoke-UiCommand `
                    -Name "Remove selected cloud licenses" `
                    -CommandPreview "Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/users/$($snapshot.CloudUser.id)/assignLicense' -Body <remove $selectedLicenseText>" `
                    -Command {
                        Invoke-CloudUserLicenseAssignment -UserIdOrUpn $snapshot.CloudUser.id -AddLicenseRows @() -RemoveLicenseRows $selectedRows | Out-Null
                    }
                if (-not (Test-UiDryRun)) {
                    Write-UiStatus -Status "OK" -Message "Removed $($selectedRows.Count) license(s) from '$($snapshot.CloudUser.userPrincipalName)'." -Color Green
                }
            }
            catch {
                Write-UiStatus -Status "WARN" -Message "Could not remove selected license(s) from '$($snapshot.CloudUser.userPrincipalName)': $($_.Exception.Message)" -Color Yellow
            }
            continue
        }

        Write-UiStatus -Status "INFO" -Message "Enter exact lowercase a, r, d, q, or x." -Color Yellow
    }
}

function Test-UiDryRun {
    $dryRunVariable = Get-Variable -Name DryRun -Scope Script -ErrorAction SilentlyContinue
    return ($dryRunVariable -and [bool]$dryRunVariable.Value)
}

function Invoke-UiCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Command,

        [string]$CommandPreview = ""
    )

    if (Test-UiDryRun) {
        Write-UiDryRunCommand -Name $Name -CommandPreview $CommandPreview
        return
    }

    & $Command
}

function Invoke-StandaloneLicenseManager {
    param(
        [string]$UserPrincipalName,

        [switch]$SkipConnect
    )

    if (-not $SkipConnect) {
        Connect-LicenseManagerGraph
    }

    $targetLookup = $UserPrincipalName
    while ([string]::IsNullOrWhiteSpace($targetLookup)) {
        $targetLookup = Read-UiInput -Prompt "Cloud user lookup" -Options @("exact UPN/email/name", "x=exit")
        if ($targetLookup -eq "x") {
            return
        }
    }

    $cloudUser = Resolve-LicenseManagerCloudUser -Lookup $targetLookup
    if (-not $cloudUser) {
        return
    }

    Invoke-InteractiveLicenseManager -UserIdOrUpn $cloudUser.id -DisplayName $cloudUser.displayName
}

if ($MyInvocation.InvocationName -ne ".") {
    Invoke-StandaloneLicenseManager -UserPrincipalName $UserPrincipalName -SkipConnect:$SkipConnect
}
