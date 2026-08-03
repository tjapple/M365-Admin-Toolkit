Set-StrictMode -Version Latest

$script:UiWidth = 120
$script:TestedModuleVersions = @{
    "ExchangeOnlineManagement"       = [version]"3.10.0"
    "Microsoft.Graph.Authentication" = [version]"2.38.0"
}

function Set-M365ToolkitUiWidth {
    [CmdletBinding()]
    param(
        [ValidateRange(60, 240)]
        [int]$Width = 120
    )

    $script:UiWidth = $Width
}

function Write-UiStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Status,

        [Parameter(Mandatory)]
        [string]$Message,

        [ConsoleColor]$Color = "Gray"
    )

    Write-Host ("[{0}] {1}" -f $Status, $Message) -ForegroundColor $Color
}

function Read-UiInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,

        [string[]]$Options = @(),

        [string]$Suffix = ">:"
    )

    $promptText = if ([string]::IsNullOrWhiteSpace($Prompt)) { "" } else { $Prompt.Trim() }
    $cleanOptions = @($Options | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($cleanOptions.Count -gt 0) {
        if (-not [string]::IsNullOrWhiteSpace($promptText)) {
            Write-Host ("  {0}" -f $promptText) -ForegroundColor Magenta
        }

        Write-Host -NoNewline ("    {0} {1} " -f ($cleanOptions -join " / "), $Suffix) -ForegroundColor Magenta
    }
    else {
        Write-Host -NoNewline ("  {0} {1} " -f $promptText, $Suffix) -ForegroundColor Magenta
    }

    $answer = $Host.UI.ReadLine()
    if ($null -eq $answer) {
        return ""
    }

    return $answer
}

function Read-UiYesNo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,

        [bool]$DefaultYes = $false
    )

    $suffix = if ($DefaultYes) { "[Y/n]" } else { "[y/N]" }
    while ($true) {
        $answer = Read-UiInput -Prompt $Prompt -Options @($suffix)
        if ([string]::IsNullOrWhiteSpace($answer)) {
            return $DefaultYes
        }

        switch -Regex ($answer.Trim()) {
            "^(y|yes)$" { return $true }
            "^(n|no)$"  { return $false }
            default {
                Write-UiStatus -Status "WARN" -Message "Please enter yes or no." -Color Yellow
            }
        }
    }
}

function Write-UiDryRunCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [string]$CommandPreview = ""
    )

    Write-UiStatus -Status "DRY RUN" -Message "Would execute: $Name" -Color DarkGray
    if ([string]::IsNullOrWhiteSpace($CommandPreview)) {
        return
    }

    foreach ($line in ($CommandPreview -split "`r?`n")) {
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            Write-Host ("  PS> {0}" -f $line) -ForegroundColor DarkGray
        }
    }
}

function Format-UiBoxValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return "(blank)"
    }

    if ($Value -is [array]) {
        $items = @($Value | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($items.Count -eq 0) {
            return "None"
        }

        return ($items -join ", ")
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return "(blank)"
    }

    return $text
}

function New-UiBoxLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Label,

        [AllowNull()]
        $Value
    )

    return ("{0,-25} {1}" -f ($Label + ":"), (Format-UiBoxValue -Value $Value))
}

function Split-UiTextToWidth {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Text,

        [Parameter(Mandatory)]
        [ValidateRange(1, 1000)]
        [int]$Width
    )

    $lines = @()
    $normalizedText = ([string]$Text) -replace "`r`n", "`n"
    $normalizedText = $normalizedText -replace "`r", "`n"
    foreach ($paragraph in ($normalizedText -split "`n")) {
        $remaining = [string]$paragraph
        if ($remaining.Length -eq 0) {
            $lines += ""
            continue
        }

        while ($remaining.Length -gt $Width) {
            $segment = $remaining.Substring(0, $Width)
            $breakIndex = $segment.LastIndexOf(" ")
            if ($breakIndex -le 0) {
                $breakIndex = $Width
            }

            $lines += $remaining.Substring(0, $breakIndex).TrimEnd()
            $remaining = $remaining.Substring($breakIndex).TrimStart()
        }

        $lines += $remaining
    }

    return $lines
}

function New-UiBoxWrappedLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Label,

        [AllowNull()]
        $Value
    )

    $innerWidth = [Math]::Max(20, $script:UiWidth - 6)
    $labelText = "{0,-25} " -f ($Label + ":")
    $continuationText = " " * $labelText.Length
    $valueWidth = [Math]::Max(10, $innerWidth - $labelText.Length)
    $wrappedValues = @(Split-UiTextToWidth -Text (Format-UiBoxValue -Value $Value) -Width $valueWidth)

    if ($wrappedValues.Count -eq 0) {
        return @($labelText)
    }

    $lines = @($labelText + $wrappedValues[0])
    for ($index = 1; $index -lt $wrappedValues.Count; $index++) {
        $lines += ($continuationText + $wrappedValues[$index])
    }

    return $lines
}

function ConvertTo-UiTableLines {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        [string[]]$Columns
    )

    $cleanRows = @($Rows | Where-Object { $_ })
    if ($cleanRows.Count -eq 0) {
        return @("None")
    }

    $widths = @{}
    foreach ($column in $Columns) {
        $widths[$column] = $column.Length
    }

    foreach ($row in $cleanRows) {
        foreach ($column in $Columns) {
            $valueText = Format-UiBoxValue -Value $row.$column
            if ($valueText.Length -gt $widths[$column]) {
                $widths[$column] = $valueText.Length
            }
        }
    }

    $headerParts = @()
    $dividerParts = @()
    foreach ($column in $Columns) {
        $headerParts += $column.PadRight($widths[$column])
        $dividerParts += ("-" * $widths[$column])
    }

    $lines = @(
        ($headerParts -join "  ")
        ($dividerParts -join "  ")
    )

    foreach ($row in $cleanRows) {
        $rowParts = @()
        foreach ($column in $Columns) {
            $rowParts += (Format-UiBoxValue -Value $row.$column).PadRight($widths[$column])
        }

        $lines += ($rowParts -join "  ")
    }

    return $lines
}

function Write-UiBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string[]]$Lines,

        [ConsoleColor]$Color = "Green"
    )

    $innerWidth = [Math]::Max(20, $script:UiWidth - 6)
    $horizontal = "-" * ($innerWidth + 2)

    Write-Host ""
    Write-Host ("  +" + $horizontal + "+") -ForegroundColor $Color
    foreach ($titleLine in (Split-UiTextToWidth -Text (" " + $Title + " ") -Width $innerWidth)) {
        Write-Host ("  | " + $titleLine.PadRight($innerWidth) + " |") -ForegroundColor $Color
    }
    Write-Host ("  +" + $horizontal + "+") -ForegroundColor $Color

    foreach ($line in $Lines) {
        foreach ($wrappedLine in (Split-UiTextToWidth -Text ([string]$line) -Width $innerWidth)) {
            $lineColor = "White"
            if (-not [string]::IsNullOrWhiteSpace($wrappedLine)) {
                $lineLabel = ($wrappedLine -split ":", 2)[0].Trim()
                if ($lineLabel -match "Enabled|Available") {
                    $lineColor = "Yellow"
                }
            }

            Write-Host -NoNewline "  | " -ForegroundColor $Color
            Write-Host -NoNewline $wrappedLine.PadRight($innerWidth) -ForegroundColor $lineColor
            Write-Host " |" -ForegroundColor $Color
        }
    }

    Write-Host ("  +" + $horizontal + "+") -ForegroundColor $Color
    Write-Host ""
}

function Import-RequiredPowerShellModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [version]$MinimumVersion
    )

    if (-not $PSBoundParameters.ContainsKey("MinimumVersion") -and $script:TestedModuleVersions.ContainsKey($ModuleName)) {
        $MinimumVersion = $script:TestedModuleVersions[$ModuleName]
    }

    $available = @(Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending)
    $compatible = if ($MinimumVersion) {
        @($available | Where-Object { $_.Version -ge $MinimumVersion })
    }
    else {
        $available
    }

    if ($compatible.Count -eq 0) {
        $versionText = if ($MinimumVersion) { " -MinimumVersion $MinimumVersion" } else { "" }
        throw "Required module '$ModuleName'$versionText is not installed. Run: Install-Module $ModuleName -Scope CurrentUser$versionText"
    }

    $importParameters = @{
        Name        = $ModuleName
        ErrorAction = "Stop"
    }
    if ($MinimumVersion) {
        $importParameters.MinimumVersion = $MinimumVersion
    }

    Import-Module @importParameters
    Write-UiStatus -Status "OK" -Message "Loaded $ModuleName $((Get-Module $ModuleName | Select-Object -First 1).Version)." -Color Green
}

Export-ModuleMember -Function @(
    "ConvertTo-UiTableLines"
    "Format-UiBoxValue"
    "Import-RequiredPowerShellModule"
    "New-UiBoxLine"
    "New-UiBoxWrappedLine"
    "Read-UiInput"
    "Read-UiYesNo"
    "Set-M365ToolkitUiWidth"
    "Split-UiTextToWidth"
    "Write-UiBox"
    "Write-UiDryRunCommand"
    "Write-UiStatus"
)
