Describe "Mailbox manager safety" {
    BeforeAll {
        $repositoryRoot = Split-Path $PSScriptRoot -Parent
        $scriptPath = Join-Path $repositoryRoot "scripts/mailboxmanager.ps1"
        . $scriptPath
    }

    It "exposes DryRun and SkipConnect parameters" {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
        $parameterNames = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        $parameterNames -contains "DryRun" | Should -Be $true
        $parameterNames -contains "SkipConnect" | Should -Be $true
    }

    It "does not execute a mutation during dry-run" {
        $script:DryRun = $true
        $script:MailboxMutationExecuted = $false
        Invoke-UiCommand -Name "test mutation" -CommandPreview "Set-DemoMailbox" -Command {
            $script:MailboxMutationExecuted = $true
        }

        $script:MailboxMutationExecuted | Should -Be $false
    }

    It "parses valid mailbox selections" {
        @((ConvertTo-ExchangeNumberSelection -InputText "1; 2 4" -Max 4)) -join "," | Should -Be "1,2,4"
    }
}
