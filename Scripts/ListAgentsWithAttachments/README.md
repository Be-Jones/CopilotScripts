# Get-AgentKnowledgeSources

Inventories Microsoft 365 Copilot agents in your tenant and reports **which agents have
files uploaded to them as knowledge sources**.

Built for access reviews: when a user uploads a file into an agent, Copilot copies it into
system-managed storage. That copy no longer inherits the original file's SharePoint
permissions, so it is worth knowing which agents hold one.

Read-only. The script never modifies the tenant.

---

## Requirements

| | |
|---|---|
| **PowerShell** | 7.0 or later |
| **Module** | `Microsoft.Graph.Authentication` |
| **Role** | AI Administrator or Global Administrator |
| **License** | **Microsoft Agent 365** — the Package Management API requires it |
| **Cloud** | Global service only — not GCC High, DoD, or 21Vianet |

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

> **PowerShell 7 is required, not optional.** The script uses ternary (`? :`) and
> null-coalescing (`??`) operators. In Windows PowerShell 5.1 it fails at parse time with
> `Unexpected token '?'` — which looks like a corrupt file but just means wrong shell.
> Check with `$PSVersionTable.PSVersion`.

### Graph permissions

| Scope | Needed for | Optional? |
|---|---|---|
| `CopilotPackages.Read.All` | The agent registry | Required |
| `User.Read.All` | Owner name and UPN | Required |

Interactive runs prompt for consent on first use. Unattended runs need both as
**Application** permissions with admin consent — see *Scheduled / unattended runs*.

---

## Quick start

```powershell
Unblock-File .\Get-AgentKnowledgeSources.ps1     # if downloaded
.\Get-AgentKnowledgeSources.ps1 -AgentBuilderOnly -CsvPath .\agents.csv
```

Sign in when prompted. Two files are produced:

| File | Contents |
|---|---|
| `agents.csv` | One row **per knowledge source** — an agent with web search plus uploaded files appears more than once |
| `agents.summary.csv` | One row **per agent** — this is the report to read |

---

## Output columns (`agents.summary.csv`)

| Column | Meaning |
|---|---|
| `Agent` | Display name |
| `Platform` | e.g. `Microsoft 365 Copilot Agent Builder` |
| `Owner` / `OwnerUpn` | Resolved from the directory, not the package's publisher string |
| `LastModifiedDateTime` | UTC, sortable (`2026-08-13 00:13:39Z`) |
| `HasAttachments` | `True` / `False` / **empty = UNKNOWN** (see below) |
| `AttachmentCount` | Count of **uploaded files only** |
| `LinkedCount` | Linked SharePoint locations — *not* counted as attachments |
| `Capabilities` | All knowledge capabilities on the agent |
| `AttachedItems` | Stable `unique_id` GUID of each uploaded file (see *Knowledge file names*) |
| `AdminCenterUrl` | Opens the admin center filtered to that agent |
| `PackageId` | Stable identifier for run-over-run comparison |

### Reading `HasAttachments`

- `True` — has uploaded files
- `False` — confirmed none
- **empty / null** — **UNKNOWN.** The registry did not return capability data for this
  agent. This is *not* the same as "no attachments" and must not be read as clean.
  `-OnlyWithAttachments` deliberately keeps these rows so they cannot vanish from a review.

### What counts as an attachment

Only files **uploaded into the agent** (`x-is_embedded` in the manifest). Linked sites,
libraries and URLs are excluded — that content stays where it is and keeps its own
SharePoint permissions. Use `-IncludeLinked` to list linked locations without counting them.

---

## Knowledge file names

The supported Graph API returns file **GUIDs**, not names, so `AttachedItems` shows values
like `unique_id:bc334af7-8d76-4279-b42e-c936d3616f7e`.

To see the actual file names, open the agent in the Microsoft 365 admin center — the
`AdminCenterUrl` column in the CSV links straight to it, filtered to that agent. Select the
agent and its knowledge files are listed in the detail panel.

This is a deliberate design choice. The names are only available from an undocumented
internal endpoint that Microsoft can change without notice, and retrieving them requires
browser-session steps that are not reasonable to ask a customer to perform. The script
therefore sticks to the supported API.

**The GUIDs are stable**, so run-over-run comparison still detects an agent gaining or
losing an attachment. Resolve names in the admin center only when you need to know *which*
file changed.

## Scheduled / unattended runs

The script can authenticate as an application, with no interactive sign-in.

> **Not yet verified.** This path is implemented but has not been run end to end.
> Do one manual run with the flags below and confirm the CSVs look right before
> putting it on a schedule.

### 1. Register an app

In **Entra admin center → App registrations → New registration**. Name it, leave redirect
URI blank, register. Note the **Application (client) ID** and **Directory (tenant) ID**.

### 2. Grant application permissions

**API permissions → Add a permission → Microsoft Graph → Application permissions**:

| Permission | For |
|---|---|
| `CopilotPackages.Read.All` | The agent registry |
| `User.Read.All` | Owner name and UPN |

Then **Grant admin consent**. Both must be **Application** permissions — delegated versions
will not work for an unattended run.

### 3. Add a client secret

**Certificates & secrets → New client secret.** Copy the **Value** immediately — it is only
shown once. Note its expiry: when the secret lapses the scheduled run starts failing, so
diarise a rotation.

### 4. Run it

```powershell
$secret = Read-Host 'Client secret' -AsSecureString

.\Get-AgentKnowledgeSources.ps1 -AgentBuilderOnly -CsvPath .\agents.csv `
    -TenantId contoso.onmicrosoft.com -ClientId <app-id> -ClientSecret $secret
```

For a scheduled task, store the secret somewhere the job can read it without embedding it in
the script — Azure Key Vault, a Windows credential, or an encrypted file created with
`ConvertFrom-SecureString`, read back with `ConvertTo-SecureString`. Never put the plain
value in the script or in the task's arguments.

The Microsoft Agent 365 license requirement still applies to the tenant.

## Parameters

| Flag | Effect |
|---|---|
| `-AgentBuilderOnly` | Filters query to only agents built in Agent Builder. Omit for all Copilot agents, including Copilot Studio |
| `-OnlyWithAttachments` | Filters output to only agents with uploaded files. UNKNOWN rows are deliberately retained |
| `-IncludeLinked` | Also list linked SharePoint locations. Never counted as attachments |
| `-CsvPath <path>` | Write both CSVs. The summary is written alongside as `<name>summary.csv` |
| `-ShowRawDefinition` | Dump the full package detail response per agent, for troubleshooting |
| `-Verbose` | Show each call and decision |

### Authentication

Interactive by default. For unattended runs:

| Flag | Effect |
|---|---|
| `-TenantId <id>` | Tenant GUID or domain. Required with `-ClientId` |
| `-ClientId <id>` | App registration id for app-only auth |
| `-ClientSecret <secure>` | Client secret as a SecureString |

The scope flags have been exercised against a live tenant; the authentication flags
have not — see *Scheduled / unattended runs*.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `Unexpected token '?'` | Running Windows PowerShell 5.1. Use PowerShell 7 |
| `402` / "unlicensed" | Package Management API needs a Microsoft Agent 365 license |
| `403` on the registry | Interactive: needs AI Administrator or Global Administrator plus consent. App-only: permissions must be **Application**, not Delegated, with admin consent granted |
| `-ClientId requires -ClientSecret` | App-only needs all three of `-TenantId`, `-ClientId`, `-ClientSecret` |
| `Client-credentials token request failed` | Wrong client id / secret / tenant, or the secret has expired |
| `404` on the registry | Global cloud only — unavailable in GCC High / DoD / 21Vianet |
| Owner columns empty | `User.Read.All` missing. Run `Disconnect-MgGraph`, re-run and accept consent |
| `A parameter cannot be found` | Older copy of the script — browsers save duplicates as `(1)` rather than overwriting |
| Names show as GUIDs | Expected — open the agent via `AdminCenterUrl` to see names |
| Agents missing from output | Check the platform values printed with `-Verbose`; try without `-AgentBuilderOnly` |

### Why direct SharePoint lookups fail

Agent Builder stores uploaded files in **SharePoint Embedded containers**, which require
two permission layers: a Graph permission *and* a container-type permission granted by the
container type's owning application. That container type is Microsoft's, so ordinary admin
tooling cannot satisfy the second layer.

Confirmed in testing: with `Sites.Read.All`, `Files.Read.All` **and**
`FileStorageContainer.Selected` all granted, every lookup still returned `403`. The script
therefore does not attempt them and does not request those scopes.

---

## How it works

1. `GET /v1.0/copilot/admin/catalog/packages` — lists agents, filtered to `supportedHosts` containing `Copilot`
2. `GET /v1.0/copilot/admin/catalog/packages/{id}` — per-agent detail
3. Parses `elementDetails[].elements[].definition` (a JSON string) into the declarative agent manifest
4. Reads `capabilities[]`, flagging `items_by_sharepoint_ids` entries with `x-is_embedded: true`

Both Graph endpoints are documented:
<https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/package/overview>

Manifest schema:
<https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/declarative-agent-manifest-1.5>

Note: Agent Builder agents return element type `DeclarativeCopilots`, not the
`declarativeAgent` shown in the documentation. The script matches both.

---

## Known limitations

- File **names** are not returned by the supported API; look them up in the admin center
- Agents whose capability data is not returned are reported as UNKNOWN, not as clean
- The `createdIn` filter in `AdminCenterUrl` is only applied with `-AgentBuilderOnly`; the portal's label for other platforms was not confirmed
- Requires a Microsoft Agent 365 license, which is separate from Copilot licensing
