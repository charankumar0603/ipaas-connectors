# How the connector calls Exchange Online

Microsoft Graph cannot change mailbox permissions or distribution list membership, and the connector sandbox cannot
run PowerShell. The connector therefore sends each Exchange cmdlet as JSON to Exchange Online's cmdlet endpoint: the
endpoint the official `ExchangeOnlineManagement` PowerShell module itself calls for every cmdlet since version 3.0.
Microsoft does not document it for direct use. Set automatic replies is the exception; it calls Microsoft Graph.

The connector's specs stub Exchange with exactly these shapes.

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
`<name>@odata.type: #Collection(String)`. Only the cmdlets in `EXCHANGE_CMDLETS` are ever sent.

## Responses

**Success, 200**: OData, the cmdlet's output objects in `value` (absent or empty for cmdlets that output nothing).
The connector drops every key holding `@` (`@odata.context`, `Name@odata.type`, ...) before the actions read them.

```json
{ "@odata.context": "...", "value": [ { "Identity": "support", "User": "jane@contoso.com", "AccessRights": ["FullAccess"] } ] }
```

**Failure**: an HTTP error status with an OData error. The first detail names the exception and its message,
separated by `|`; the connector matches the exception's short name (`MemberAlreadyExistsException`,
`MemberNotFoundException`, `ManagementObjectNotFoundException`, `UserNotFoundInPermissionEntryException`).

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
| 502, 503, 504 | Pauses the step for 60 s |
| No answer (timeout, refused or dropped connection) | Pauses the step for 30 s |
| Sign-in answers 429 or 5xx, or does not answer | Pauses the step for 30 s; refused credentials fail at once |
| Other | Fails with the status, the exception and the message, unless the action treats the exception as "already so" |

Every action reads the current state before it changes anything, so a step that runs again after a pause is safe
even when its change had gone through. Microsoft Graph (automatic replies) is paused the same way on 429, 502, 503,
504 and no answer.

## Confirmed against a live tenant

Checked on 2026-11-26 against a Microsoft 365 Business Premium tenant (Global cloud), every action run through iPaaS
and every change read back with a separate client:

- App-only tokens with Exchange.ManageAsApp and the Exchange Administrator role are accepted on `beta`. The
  `@odata.type` markers are accepted for hashtables (`@{Add=...}`, `@{Remove=...}`) and string lists.
- Statuses and exceptions: not found 404 `ManagementObjectNotFoundException`; already a member 400
  `MemberAlreadyExistsException`; not a member 404 `MemberNotFoundException`; no calendar entry 404
  `UserNotFoundInPermissionEntryException`; hold without the licence 400 `RecipientTaskException`. The detail reads
  `Ex<id>|<namespace>.<Exception>|<message>`.
- Shapes:
  - `AccessRights` is a list whose items may join several rights: `["FullAccess, ReadPermission"]`.
  - `GrantSendOnBehalfTo` lists Entra object IDs.
  - `EmailAddresses` reads `["smtp:alias@...", "SMTP:primary@..."]`.
  - `SharingPermissionFlags` is a string, `"Delegate, CanViewPrivateItems"`, or null.
  - `TotalItemSize` reads `"11.74 KB (12,020 bytes)"`.
  - `LitigationHoldDuration` reads `"30.00:00:00"` or `"Unlimited"`.
  - `ForwardingSmtpAddress` reads `"smtp:user@..."`.
- Switching litigation hold on records the caller as `LitigationHoldOwner` and ignores an owner in the same call, so
  the connector sets the owner in a second call.
- A mailbox's type reads as before for a few seconds after `Set-Mailbox -Type`.
- Times: a token 0.25-0.6 s (cached for its lifetime), reads 0.15-1.3 s, most writes 0.4-4.5 s, calendar folder
  permission writes 8-12 s. No throttling in about 80 runs.

## Still to confirm

- The GCC High, DoD and China hosts.
- A mailbox synchronized from on-premises Active Directory (the test tenant is cloud-only).
