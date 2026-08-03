Describe "M365AdminToolkit shared module" {
    BeforeAll {
        $repositoryRoot = Split-Path $PSScriptRoot -Parent
        $modulePath = Join-Path $repositoryRoot "src/M365AdminToolkit/M365AdminToolkit.psd1"
        Import-Module $modulePath -Force
    }

    It "has a valid module manifest" {
        { Test-ModuleManifest $modulePath -ErrorAction Stop } | Should -Not -Throw
    }

    It "formats blank and collection values consistently" {
        (Format-UiBoxValue -Value $null) | Should -Be "(blank)"
        (Format-UiBoxValue -Value @()) | Should -Be "None"
        (Format-UiBoxValue -Value @("one", "two")) | Should -Be "one, two"
    }

    It "wraps long text without exceeding the requested width" {
        $lines = @(Split-UiTextToWidth -Text "alpha beta gamma delta epsilon" -Width 10)
        @($lines | Where-Object { $_.Length -gt 10 }).Count | Should -Be 0
        ($lines -join " ") | Should -Be "alpha beta gamma delta epsilon"
    }

    It "creates readable tables" {
        $lines = @(ConvertTo-UiTableLines -Rows @([pscustomobject]@{ Name = "Demo"; State = "Enabled" }) -Columns @("Name", "State"))
        $lines.Count | Should -Be 3
        $lines[0] | Should -Match "Name"
        $lines[2] | Should -Match "Demo"
    }
}
