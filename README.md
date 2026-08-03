# M365 Admin Toolkit

A safety-focused PowerShell toolkit for common Microsoft 365 administration
workflows. It provides interactive terminal interfaces for group membership,
license assignment, and Exchange Online mailbox management while keeping
credentials out of source code and making mutations explicit.

This is an independent community project and is not affiliated with or
endorsed by Microsoft.

## What it includes

| Tool | Capabilities | Service |
| --- | --- | --- |
| `groupmanager.ps1` | Search users and groups; review, add, and remove memberships; distinguish cloud, dynamic, synced, and Exchange-managed groups | Microsoft Graph and Exchange Online |
| `licensemanager.ps1` | Review tenant SKU inventory and user assignments; set usage location; add or remove licenses | Microsoft Graph |
| `mailboxmanager.ps1` | Review mailbox state; manage delegates, mailbox properties, archives, and inbox rules | Exchange Online |

Shared terminal rendering and dependency validation live in the
`M365AdminToolkit` module so each entry-point script stays focused on its
administrative workflow.

## Safety defaults

- `-DryRun` and `-WhatIf` preview mutations without executing them.
- Every interactive mutation requires an explicit confirmation whose default
  is **No**.
- Existing Graph and Exchange Online sessions are not forcibly disconnected.
- Dependencies are validated but never installed automatically.
- Microsoft Graph requests use delegated authentication and documented,
  task-specific scopes.
- The Exchange `BypassSecurityGroupManagerCheck` override is disabled unless
  `-BypassGroupOwnershipCheck` is supplied explicitly.
- No access tokens, passwords, tenant IDs, domains, or account identifiers are
  stored by the repository.

Start with dry-run mode in a test tenant. These tools can still make
high-impact production changes when dry-run mode is disabled.

## Requirements

- Windows with PowerShell 7.2 or later
- An account with the appropriate Microsoft Entra and/or Exchange Online role
- `Microsoft.Graph.Authentication` 2.38.0 or later for Graph-based tools
- `ExchangeOnlineManagement` 3.10.0 or later for Exchange-based tools

Install dependencies explicitly:

```powershell
Install-Module Microsoft.Graph.Authentication -MinimumVersion 2.38.0 -Scope CurrentUser
Install-Module ExchangeOnlineManagement -MinimumVersion 3.10.0 -Scope CurrentUser
```

The versions above are the tested baselines, not embedded credentials or
tenant-specific requirements.

## Usage

Clone the repository, open PowerShell in its root, and preview a workflow with
fictional Contoso identities:

```powershell
./scripts/groupmanager.ps1 -UserPrincipalName alex.wilber@contoso.com -DryRun
./scripts/licensemanager.ps1 -UserPrincipalName alex.wilber@contoso.com -WhatIf
./scripts/mailboxmanager.ps1 -UserPrincipalName alex.wilber@contoso.com -DryRun
```

If a compatible authenticated session already exists, use `-SkipConnect` to
reuse it during development or testing. The tools never persist the session.

For distribution groups that the signed-in Exchange administrator does not
own, the group manager can opt into the ownership override:

```powershell
./scripts/groupmanager.ps1 -BypassGroupOwnershipCheck -DryRun
```

Review the preview, remove `-DryRun`, and confirm each proposed mutation only
after validating the selected user and resource.

## Permissions

The Graph tools intentionally avoid `Directory.ReadWrite.All` and
`Group.ReadWrite.All`. The exact scopes, supporting Microsoft documentation,
and required administrator roles are documented in
[docs/PERMISSIONS.md](docs/PERMISSIONS.md).

Exchange Online uses role-based access control from the signed-in session. The
toolkit does not attempt to elevate or assign roles.

## Repository structure

```text
scripts/                    Interactive entry points
src/M365AdminToolkit/       Shared PowerShell module
tests/                      Pester safety and unit tests
docs/                       Permission and security design notes
examples/                   Sanitized demonstration output
.github/workflows/          CI and secret-scanning automation
```

## Testing

```powershell
Install-Module Pester -MinimumVersion 5.6.1 -Scope CurrentUser
Install-Module PSScriptAnalyzer -MinimumVersion 1.22.0 -Scope CurrentUser

Invoke-Pester -Path ./tests
Invoke-ScriptAnalyzer -Path ./scripts,./src,./tests -Recurse -Settings ./PSScriptAnalyzerSettings.psd1
```

CI parses every PowerShell file, runs Pester, runs PSScriptAnalyzer, and scans
the repository for secrets. Cloud connections are mocked or skipped; CI never
needs tenant credentials.

## Demonstration data

Only publish screenshots and recordings created with fictional identities.
The sample in [examples/demo-transcript.txt](examples/demo-transcript.txt) uses
Microsoft's conventional `contoso.com` example domain. Do not record a real
tenant ID, UPN, mailbox address, group name, delegate, SKU inventory, or error
message from a production environment.

## License and security

Released under the [MIT License](LICENSE). See [SECURITY.md](SECURITY.md) for
responsible vulnerability reporting and [docs/SECURITY_MODEL.md](docs/SECURITY_MODEL.md)
for the toolkit's trust boundaries and operational safeguards.
