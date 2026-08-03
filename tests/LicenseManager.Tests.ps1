Describe "License manager safety" {
    BeforeAll {
        $repositoryRoot = Split-Path $PSScriptRoot -Parent
        . (Join-Path $repositoryRoot "scripts/licensemanager.ps1")
    }

    It "requests only the documented delegated Graph scopes" {
        @((Get-LicenseManagerGraphScopes)) -join "," | Should -Be "User.Read.All,User.ReadUpdate.All,LicenseAssignment.Read.All,LicenseAssignment.ReadWrite.All"
    }

    It "does not execute a mutation during dry-run" {
        $script:DryRun = $true
        $script:LicenseMutationExecuted = $false
        Invoke-UiCommand -Name "test mutation" -CommandPreview "Set-DemoValue" -Command {
            $script:LicenseMutationExecuted = $true
        }

        $script:LicenseMutationExecuted | Should -Be $false
    }

    It "escapes apostrophes in Graph filters" {
        (Escape-GraphFilterValue -Value "O'Neil") | Should -Be "O''Neil"
    }
}
