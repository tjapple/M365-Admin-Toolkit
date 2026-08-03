# Permissions

The toolkit uses delegated authentication: actions run as the signed-in
administrator and are also constrained by that person's Microsoft Entra or
Exchange Online role. Granting an OAuth scope does not grant the operator an
administrator role.

## Group manager

| Scope | Why it is requested |
| --- | --- |
| `User.Read.All` | Search users and read the directory properties displayed by the tool |
| `GroupMember.ReadWrite.All` | List basic group properties and read, add, or remove user memberships |

`GroupMember.ReadWrite.All` is the documented least-privileged delegated
permission for adding a user as a group member. The signed-in user must also
hold an appropriate role or group ownership rights.

Sources: [Add members](https://learn.microsoft.com/graph/api/group-post-members),
[Microsoft Graph permissions reference](https://learn.microsoft.com/graph/permissions-reference).

Exchange-managed distribution groups are changed through Exchange Online
PowerShell and its RBAC model. Ownership bypass is opt-in and still requires an
authorized Exchange role.

## License manager

| Scope | Why it is requested |
| --- | --- |
| `User.Read.All` | Search users and display assignment state |
| `User.ReadUpdate.All` | Set `usageLocation` when it is required before licensing |
| `LicenseAssignment.Read.All` | List subscribed SKUs and availability |
| `LicenseAssignment.ReadWrite.All` | Add or remove user license assignments |

Sources: [Update user](https://learn.microsoft.com/graph/api/user-update),
[List subscribed SKUs](https://learn.microsoft.com/graph/api/subscribedsku-list),
[Assign license](https://learn.microsoft.com/graph/api/user-assignlicense).

Microsoft documents License Administrator and User Administrator among the
supported delegated roles for license assignment. Choose the least privileged
role that matches your operating model.

## Mailbox manager

The mailbox manager requests no Graph scopes. It connects interactively to
Exchange Online and relies on Exchange RBAC for every command. Required roles
depend on the selected action; the toolkit neither grants roles nor bypasses
RBAC.

## Review cadence

Microsoft can add narrower permissions or change API requirements. Re-check
the linked documentation before each tagged release and update both the code
and this file together.
