# How the connector calls Exchange Online

Microsoft Graph cannot change mailbox permissions or distribution list membership, and the connector sandbox cannot
run PowerShell. The connector therefore sends each Exchange cmdlet as JSON to Exchange Online's cmdlet endpoint: the
endpoint the official `ExchangeOnlineManagement` PowerShell module itself calls for every cmdlet since version 3.0.
Microsoft does not document it for direct use. Set automatic replies is the exception; it calls Microsoft Graph.

## Endpoint

```text
POST {exchange host}/adminapi/beta/{tenant id}/InvokeCommand
Authorization: Bearer <app-only token for the Exchange host>
Content-Type: application/json
Accept: application/json
X-ResponseFormat: json
X-AnchorMailbox: APP:SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}@{tenant id}
```

| Cloud | Exchange host (token resource and endpoint) | Login host |
|---|---|---|
| Global | `https://outlook.office365.com` | `https://login.microsoftonline.com` |
| US Government GCC High | `https://outlook.office365.us` | `https://login.microsoftonline.us` |
| US Government DoD | `https://outlook-dod.office365.us` | `https://login.microsoftonline.us` |
| China | `https://partner.outlook.cn` | `https://login.chinacloudapi.cn` |

The token is a client-credentials token for `{exchange host}/.default`; the app needs the application permission
**Exchange.ManageAsApp** with admin consent, and an Exchange role for its service principal. The Admin API's
`Exchange.AdminAPI.ManageAsApp` permission does not open this endpoint. The connector waits up to 60 seconds (5 to
connect).

## Request body

```json
{
  "CmdletInput": {
    "CmdletName": "Set-Mailbox",
    "Parameters": {
      "Identity": "bob@contoso.com",
      "EmailAddresses": { "add": ["smtp:robert@contoso.com"], "@odata.type": "#Exchange.GenericHashTable" },
      "AccessRights": ["FullAccess"],
      "AccessRights@odata.type": "#Collection(String)"
    }
  }
}
```

A hash parameter (add and remove lists) is marked `#Exchange.GenericHashTable`; a list of strings gets a sibling
`<name>@odata.type: #Collection(String)`. A `null` parameter clears the value (`ForwardingSmtpAddress: null`).

## Responses

**Success, 200**: OData, the cmdlet's output objects in `value` (absent or empty for cmdlets that output nothing).
The connector drops every key holding `@` (`@odata.context`, `Name@odata.type`, ...) before the actions read them. A
200 whose body is not JSON fails the step, so a proxy or sign-in page is never read as an empty answer.

```json
{ "@odata.context": "...", "value": [ { "Identity": "support", "User": "jane@contoso.com", "AccessRights": ["FullAccess"] } ] }
```

**Failure**: an HTTP error status with an OData error. The first detail reads `Ex<id>|<namespace>.<Exception>|<text>`;
the connector matches the exception's short name.

| Exception | Status | Read as |
|---|---|---|
| `ManagementObjectNotFoundException` | 404 | no such recipient (an address is free, a forwarding target is missing) |
| `MemberAlreadyExistsException` | 400 | the user already is a member |
| `MemberNotFoundException` | 404 | the user is not a member |
| `UserNotFoundInPermissionEntryException` | 404 | the user has no entry on the folder |
| `RecipientTaskException` | 400 | a refusal, such as a hold on a mailbox without the licence; fails the step |

```json
{ "error": { "code": "NotFound", "message": "Error executing cmdlet",
  "details": [ { "code": "Context", "target": "",
    "message": "Ex6F9304|Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException|The operation couldn't be performed because object 'x' couldn't be found on '...'." } ] } }
```

| Status | Connector |
|---|---|
| 401 | Fails, naming Exchange.ManageAsApp, admin consent, Tenant ID and Cloud |
| 403 | Fails, naming the Exchange role of the service principal |
| 429 | Pauses the step for `Retry-After` (60 s without it) |
| 503 | Pauses the step for 60 s |
| 502, 504 | A read pauses for 60 s. A change fails: its answer was lost, so it may have been applied |
| No answer in time (read timeout) | A read pauses for 30 s. A change fails, as for 502 and 504 |
| No connection (refused, DNS, connect timeout) | Pauses the step for 30 s: nothing reached Exchange |
| Sign-in answers 429 or 5xx, or does not answer | Pauses the step for 30 s; refused credentials fail at once |
| Other | Fails with the status, the exception and the message, unless the action treats the exception as "already so" |

A change whose answer is lost fails with "the change may have been applied" instead of running again, because a second
run would see the change and report `changed: false` for it. Microsoft Graph (automatic replies) follows the same
rules. Changes made after the admin grants a permission or a role reach the connector within an hour, when its cached
token expires.

## Shapes in Exchange's answers

- `AccessRights` is a list whose items may join several rights: `["FullAccess, ReadPermission"]`.
- `GrantSendOnBehalfTo` and `ForwardingAddress` name each recipient by its `Name`, which in newer tenants is its
  Entra object ID. The connector matches a recipient by name, identity, primary address or object ID, never by display
  name, which two people can share.
- `EmailAddresses` reads `["smtp:alias@...", "SMTP:primary@..."]`; upper-case `SMTP` is the primary address.
- `ForwardingSmtpAddress` reads `"smtp:user@..."`. Exchange ignores it while `ForwardingAddress` is set, so the
  connector sets `ForwardingAddress` and clears `ForwardingSmtpAddress`.
- `SharingPermissionFlags` is a string, `"Delegate, CanViewPrivateItems"`, or null. `Set-MailboxFolderPermission`
  without the parameter keeps a delegate a delegate; only `None` takes it away.
- `TotalItemSize` reads `"11.74 KB (12,020 bytes)"`; the byte count is the exact part.
- `LitigationHoldDuration` reads `"30.00:00:00"` or `"Unlimited"`.
- Booleans may come as `true` or as the string `"True"`.

## Exchange behaviour the actions rely on

- Switching litigation hold on records the caller as `LitigationHoldOwner` and ignores an owner in the same call; an
  owner set while the hold is on is kept. The connector sets the owner in a second call.
- A mailbox's type reads as before for a few seconds after `Set-Mailbox -Type`.
- Granting calendar Delegate also adds the user to the mailbox's `GrantSendOnBehalfTo`; removing the folder permission
  removes that entry too.
- Calendar folder permission changes take 8 to 12 seconds; other changes a few seconds.

## Still to confirm

- The GCC High, DoD and China hosts.
- A mailbox synchronized from on-premises Active Directory.
