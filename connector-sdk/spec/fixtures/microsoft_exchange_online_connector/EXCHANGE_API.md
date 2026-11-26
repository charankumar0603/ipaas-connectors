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
| Other | Fails with the status, the exception and the message, unless the action treats the exception as "already so" |

## Still to confirm against a live tenant

- Which statuses Exchange sends for each failure (not found, already a member, no role), and that every exception
  name above appears in `details`.
- Whether the `@odata.type` markers are required, and accepted, on `beta` for hashtables and string lists.
- The JSON shapes of `GrantSendOnBehalfTo`, `EmailAddresses`, `AccessRights`, `SharingPermissionFlags`,
  `TotalItemSize` and `LitigationHoldDuration` in REST output.
- The DoD and China hosts.
