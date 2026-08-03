# Security model

## Trust boundaries

- The operator controls the workstation, installed PowerShell modules, and
  interactive sign-in.
- Microsoft Graph and Exchange Online enforce authorization and tenant roles.
- The toolkit holds session state only in the current PowerShell process.
- The repository contains no unattended authentication flow and accepts no
  client secret, certificate, password, or access token parameter.

## Safeguards

1. Read operations load current state before presenting a mutation.
2. Selection screens display the target user and affected resources.
3. Confirmation prompts default to No.
4. `-DryRun` and `-WhatIf` display command previews and skip mutation calls.
5. The group ownership bypass is disabled by default.
6. Module dependencies are checked against tested minimum versions and are not
   installed or upgraded at runtime.
7. Existing cloud sessions are reused rather than silently destroyed.

## Data handling

The tools display tenant data in the terminal but do not write it to disk.
PowerShell host logging, shell history, endpoint monitoring, or a transcript
started by the operator can still capture that output. Store any such records
according to the tenant's data-handling policy and never commit them.

Command previews intentionally replace request bodies with placeholders where
the payload could contain an object identifier.

## Known limitations

- Interactive confirmation cannot protect against an operator selecting the
  wrong similarly named object. Verify immutable identifiers where available.
- Microsoft 365 APIs are eventually consistent; refreshed state can lag after
  a successful mutation.
- Dry-run validates control flow, not whether the current account will be
  authorized for the real command.
- Third-party PowerShell Gallery and GitHub Actions dependencies remain part of
  the software supply chain and should be reviewed when versions change.
