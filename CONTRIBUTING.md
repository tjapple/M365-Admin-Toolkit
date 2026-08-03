# Contributing

Contributions are welcome when they preserve least privilege, dry-run support,
explicit confirmation, and tenant-data privacy.

1. Create a focused branch.
2. Use only fictional `contoso.com` identities in tests, issues, screenshots,
   and documentation.
3. Add or update Pester tests for behavior changes.
4. Run Pester and PSScriptAnalyzer locally.
5. Describe requested Graph scopes and link Microsoft documentation whenever a
   change introduces or broadens a permission.
6. Never commit exports, transcripts, tenant identifiers, credentials, or
   copied production error messages.

Pull requests that add a mutation must support `-DryRun`/`-WhatIf`, show a
specific preview, and require a confirmation that defaults to No.
