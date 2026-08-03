Describe "Group manager safety" {
    BeforeAll {
        $repositoryRoot = Split-Path $PSScriptRoot -Parent
        . (Join-Path $repositoryRoot "scripts/groupmanager.ps1")
        function Invoke-MgGraphRequest { throw "The dry-run safeguard called Graph unexpectedly." }
    }

    It "requests only the documented delegated Graph scopes" {
        @((Get-GroupManagerGraphScopes)) -join "," | Should -Be "User.Read.All,GroupMember.ReadWrite.All"
    }

    It "parses comma and whitespace separated selections" {
        @((ConvertTo-NumberSelection -InputText "1, 3 4" -Max 4)) -join "," | Should -Be "1,3,4"
    }

    It "rejects an out-of-range selection" {
        @(ConvertTo-NumberSelection -InputText "1,5" -Max 4).Count | Should -Be 0
    }

    It "does not call Graph when dry-run is enabled" {
        $script:DryRun = $true
        $cloudUser = [pscustomobject]@{ id = "demo-user-id"; userPrincipalName = "alex.wilber@contoso.com" }
        $group = [pscustomobject]@{
            id = "demo-group-id"
            displayName = "Demo Security Group"
            groupTypes = @()
            mailEnabled = $false
            securityEnabled = $true
            onPremisesSyncEnabled = $false
        }

        { Invoke-GroupManagerCloudGroupMembershipChange -CloudUser $cloudUser -Group $group -Action Add } | Should -Not -Throw
    }
}
