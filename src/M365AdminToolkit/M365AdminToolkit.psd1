@{
    RootModule        = "M365AdminToolkit.psm1"
    ModuleVersion     = "1.0.0"
    GUID              = "5f0d2f9a-bcbd-4f09-8fb4-9ef4d56e1a65"
    Author            = "m365-admin-toolkit contributors"
    CompanyName       = "Community"
    Copyright         = "Copyright (c) 2026 m365-admin-toolkit contributors"
    Description       = "Shared terminal UI and dependency helpers for the M365 Admin Toolkit."
    PowerShellVersion = "7.2"
    FunctionsToExport = @(
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
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @("Microsoft365", "ExchangeOnline", "MicrosoftGraph", "Administration")
            LicenseUri = "https://opensource.org/license/mit"
        }
    }
}
