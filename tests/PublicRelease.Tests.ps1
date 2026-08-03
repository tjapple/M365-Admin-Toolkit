Describe "Public release hygiene" {
    BeforeAll {
        $repositoryRoot = Split-Path $PSScriptRoot -Parent
        $powerShellFiles = @(Get-ChildItem $repositoryRoot -Recurse -File -Include *.ps1,*.psm1,*.psd1)
        $script:ReleaseSourceFiles = @($powerShellFiles | Where-Object { $_.Name -ne "PublicRelease.Tests.ps1" })
    }

    It "contains no PowerShell syntax errors" {
        $allErrors = @()
        foreach ($file in $powerShellFiles) {
            $tokens = $null
            $parseErrors = $null
            [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors) | Out-Null
            $allErrors += @($parseErrors)
        }

        $allErrors.Count | Should -Be 0
    }

    It "does not auto-install, force-clobber, or disconnect cloud sessions" {
        $source = ($script:ReleaseSourceFiles | ForEach-Object { Get-Content $_.FullName -Raw }) -join "`n"
        $source | Should -Not -Match "\bInstall-Module\b.*-Force"
        $source | Should -Not -Match "\bAllowClobber\b"
        $source | Should -Not -Match "\bDisconnect-(MgGraph|ExchangeOnline)\b"
    }

    It "contains no obvious private identifiers or credential signatures" {
        $source = (Get-ChildItem $repositoryRoot -Recurse -File | Where-Object { $_.FullName -notmatch "[\\/]\.git[\\/]" -and $_.Name -ne "PublicRelease.Tests.ps1" } | ForEach-Object { Get-Content $_.FullName -Raw -ErrorAction SilentlyContinue }) -join "`n"
        $source | Should -Not -Match "(?i)(gmail\.com|C:\\Users\\|BEGIN [A-Z ]*PRIVATE KEY|AKIA[0-9A-Z]{16}|github_pat_[A-Za-z0-9_]{20,}|gh[pousr]_[A-Za-z0-9]{20,}|xox[baprs]-[A-Za-z0-9-]{10,})"
    }
}
