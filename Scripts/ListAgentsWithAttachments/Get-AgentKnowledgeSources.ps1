#Requires -Version 7.0

<#
.SYNOPSIS
    Inventories Copilot agents in the tenant agent registry and reports which ones
    have knowledge sources, flagging file/attachment-backed sources.

.DESCRIPTION
    Uses the Microsoft 365 Copilot Package Management API (the documented agent
    registry that backs the admin center agent inventory).

    WHAT COUNTS AS AN ATTACHMENT
    Only files UPLOADED into the agent count. In the manifest these appear under
    capabilities[].items_by_sharepoint_ids with 'x-is_embedded': true - Agent
    Builder copies them into a system-managed library named PublishedAgent_<guid>.
    Linked sites, libraries and items_by_url are excluded: that content stays in
    place and keeps its own SharePoint permissions. Use -IncludeLinked to see it.

    Endpoints used - both documented:
      GET /v1.0/copilot/admin/catalog/packages        (List packages)
      GET /v1.0/copilot/admin/catalog/packages/{id}   (Get package details)
      https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/package/overview

    For each package the script reads elementDetails[] entries whose elementType
    is 'declarativeAgent', parses the JSON in the 'definition' property, and
    extracts the manifest 'capabilities' array.

    Capability names and properties follow the declarative agent manifest schema:
      https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/declarative-agent-manifest-1.5

    CONFIRMED AGAINST A LIVE TENANT:
    The registry DOES return the full declarative agent manifest, including the
    'capabilities' array, inside elementDetails[].elements[].definition (a JSON
    string). Microsoft's published example shows a truncated definition; that is
    an abbreviated doc sample, not the real payload shape.

    Note the element type for Agent Builder agents is 'DeclarativeCopilots',
    not the 'declarativeAgent' shown in the docs. Both are matched.

    Where capability data is genuinely absent, the agent is reported as UNKNOWN
    (HasKnowledge = $null) - never as a clean 'no knowledge sources'.

.PARAMETER TenantId
    Tenant GUID or domain (e.g. contoso.onmicrosoft.com). REQUIRED for app-only
    authentication; optional for interactive sign-in.

.PARAMETER ClientId
    App registration (client) id for unattended app-only authentication.
    Requires -TenantId and -ClientSecret.
    The app needs CopilotPackages.Read.All and User.Read.All as APPLICATION
    permissions with admin consent granted. See README for setup.

.PARAMETER ClientSecret
    Client secret as a SecureString, for app-only authentication. Pair with
    -ClientId and -TenantId. Note this is a stored credential with an expiry -
    plan to rotate it before it lapses or the scheduled run starts failing.

.PARAMETER AgentBuilderOnly
    Restrict to agents created with Microsoft 365 Copilot Agent Builder.

.PARAMETER OnlyWithAttachments
    Return only agents that have uploaded file attachments as knowledge.

.PARAMETER IncludeLinked
    Also list linked SharePoint/OneDrive locations (sites, libraries, items_by_url)
    alongside uploaded files. Linked content is NEVER counted as an attachment -
    it keeps its own SharePoint permissions - this switch only makes it visible.

.PARAMETER ShowRawDefinition
    Emit the raw declarativeAgent definition JSON for inspection. Use this first
    to confirm what your tenant returns.

.PARAMETER CsvPath
    Optional path to write the report as CSV.

.EXAMPLE
    .\Get-AgentKnowledgeSources.ps1 -AgentBuilderOnly -CsvPath .\agents.csv

    Interactive sign-in. Reports every Agent Builder agent and writes
    agents.csv (one row per knowledge source) plus agents.summary.csv
    (one row per agent).

.EXAMPLE
    .\Get-AgentKnowledgeSources.ps1 -AgentBuilderOnly -OnlyWithAttachments -IncludeLinked -CsvPath .\agents.csv

    Only agents that have uploaded files, with linked SharePoint locations also
    listed (linked content is never counted as an attachment).

.EXAMPLE
    $secret = Read-Host 'Client secret' -AsSecureString
    .\Get-AgentKnowledgeSources.ps1 -AgentBuilderOnly -CsvPath .\agents.csv `
        -TenantId contoso.onmicrosoft.com -ClientId <app-id> -ClientSecret $secret

    Unattended run for a scheduled task. No interactive sign-in.

.EXAMPLE
    .\Get-AgentKnowledgeSources.ps1 -ShowRawDefinition -Verbose

    Troubleshooting: dumps the full package detail response for each agent.

.LINK
    https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/package/overview

.LINK
    https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/declarative-agent-manifest-1.5

.NOTES
    Version  : 1.0
    Requires : PowerShell 7.0 or later (uses ternary and null-coalescing operators;
               Windows PowerShell 5.1 fails with a parse error)
               Microsoft.Graph.Authentication module
    Scopes   : CopilotPackages.Read.All  (agent registry - required)
               User.Read.All             (owner name and UPN)

    WHY FILE NAMES MAY NOT RESOLVE (documented, not a script defect):
    Agent Builder stores uploaded knowledge in SharePoint Embedded containers
    (the PublishedAgent_<guid> libraries). SharePoint Embedded uses TWO permission
    layers and BOTH are required:

      1. A Microsoft Graph permission - FileStorageContainer.Selected. Delegated
         does not need admin consent.
      2. Container-type level application permission, granted by the container
         type's OWNING application through container type registration.

    Microsoft states plainly that Graph consent alone does not grant access to
    containers; the app must also be granted permission to the container type.
    The Agent Builder container type is owned by Microsoft, and Microsoft Graph
    PowerShell is not its owning application - so layer 2 cannot be satisfied
    here, and name lookups will usually return 403 no matter what is consented.
    This was confirmed in testing: with Sites.Read.All, Files.Read.All AND
    FileStorageContainer.Selected all granted, every lookup still returned 403.
    That is why these scopes are no longer requested by default.

    ATTACHMENT DETECTION AND COUNTS ARE UNAFFECTED. They are read from the agent
    manifest via the registry API, not from SharePoint.

    KNOWLEDGE FILE NAMES ARE NOT RETURNED. The report identifies each uploaded
    file by its stable unique_id GUID. To see file names, open the agent in the
    Microsoft 365 admin center - the AdminCenterUrl column links straight to it.
    GUIDs are stable, so run-over-run comparison still detects an agent gaining
    or losing an attachment.

    Auth     : Interactive by default. For unattended runs pass -ClientId with
               -TenantId and -ClientSecret; the
               app registration needs BOTH scopes above as APPLICATION permissions
               with admin consent.
    Role     : AI Administrator or Global Administrator (interactive runs)
    LICENSE  : The Package Management API requires a Microsoft Agent 365 license.
               Without it these endpoints return 402/403 regardless of role.
    Cloud    : Global service only - not GCC High, DoD, or 21Vianet.
    Read-only: this script never modifies the tenant.
#>

[CmdletBinding()]
param(
    [switch] $AgentBuilderOnly,
    [string] $TenantId,
    [string] $ClientId,
    [securestring] $ClientSecret,
    [switch] $OnlyWithAttachments,
    [switch] $IncludeLinked,
    [switch] $ShowRawDefinition,
    [string] $CsvPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Agent element type names vary by surface. Agent Builder agents come back as
# 'DeclarativeCopilots' (confirmed against a live tenant response), while the
# published API examples show 'declarativeAgent'. Match both, plus custom engine.
$script:AgentElementTypePattern = '(?i)declarative(Agent|Copilots?)|customEngineAgent'
$script:PlatformValuesSeen = @()
$script:ApiVersion = 'v1.0'
$script:BaseUri    = "https://graph.microsoft.com/$($script:ApiVersion)/copilot/admin/catalog/packages"

#region ---------- connection ----------

function Get-SafeProperty {
    # Set-StrictMode -Version Latest throws on ANY property access against $null
    # or a missing property. Graph context objects vary by auth mode (an app-only
    # context has no Account; a disconnected session has no context at all), so
    # every context read goes through here.
    param($Object, [string] $Name)

    if ($null -eq $Object) { return $null }
    if ($Object.PSObject.Properties.Name -notcontains $Name) { return $null }
    return $Object.$Name
}

function Connect-Registry {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft.Graph.Authentication is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $required = @('CopilotPackages.Read.All', 'User.Read.All')

    # ---- validate the app-only combination before touching the network ----
    $appOnly = [bool]$ClientId
    if ($appOnly) {
        if (-not $TenantId) {
            throw "-TenantId is required with -ClientId. Use the tenant GUID or domain, e.g. contoso.onmicrosoft.com."
        }
        if (-not $ClientSecret) {
            throw "-ClientId requires -ClientSecret."
        }
    }
    elseif ($ClientSecret) {
        throw "-ClientSecret requires -ClientId and -TenantId."
    }

    if ($appOnly) {
        # App-only: permissions come from admin-consented APPLICATION roles on the
        # app registration, so no -Scopes list is passed and no consent prompt
        # appears. Both required permissions must be granted as Application (not
        # Delegated) permissions or the calls fail with 403.
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

        # Connect-MgGraph has no client-secret parameter; acquire a token via the
        # client-credentials flow and hand it over.
        Write-Verbose "Acquiring app-only token via client credentials."
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
                    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret))
        try {
            $tok = Invoke-RestMethod -Method POST -ErrorAction Stop `
                -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                -ContentType 'application/x-www-form-urlencoded' `
                -Body @{
                    client_id     = $ClientId
                    client_secret = $plain
                    scope         = 'https://graph.microsoft.com/.default'
                    grant_type    = 'client_credentials'
                }
        }
        catch {
            throw "Client-credentials token request failed: $($_.Exception.Message). Check the client id, secret and tenant id, that the secret has not expired, and that admin consent has been granted for CopilotPackages.Read.All and User.Read.All as APPLICATION permissions."
        }
        finally { $plain = $null }

        Connect-MgGraph -AccessToken ($tok.access_token | ConvertTo-SecureString -AsPlainText -Force) -NoWelcome | Out-Null
        $tok = $null

        $ctx = Get-MgContext
        Write-Verbose "App-only connection as client $((Get-SafeProperty -Object $ctx -Name 'ClientId'))"
        return
    }

    # ---- delegated (interactive) ------------------------------------------
    $ctx = $null
    try { $ctx = Get-MgContext } catch { }

    # Reconnect when any required scope is absent. Graph PowerShell reuses a
    # cached session, so a session created before a scope was needed will not
    # have it and the dependent lookup fails with no prompt.
    $held   = @(Get-SafeProperty -Object $ctx -Name 'Scopes')
    $absent = @($required | Where-Object { $held -notcontains $_ })

    if (-not $ctx -or $absent.Count) {
        if ($ctx -and $absent.Count) {
            Write-Verbose "Existing session is missing: $($absent -join ', '). Reconnecting."
        }
        $connect = @{ Scopes = $required }
        if ($TenantId) { $connect.TenantId = $TenantId }
        Connect-MgGraph @connect | Out-Null
        $ctx = Get-MgContext
    }

    $who = (Get-SafeProperty -Object $ctx -Name 'Account') ?? '(unknown)'
    Write-Verbose "Connected as $who"
    Write-Verbose "Granted scopes: $(@(Get-SafeProperty -Object $ctx -Name 'Scopes') -join ', ')"
}

function Invoke-GraphPaged {
    param([Parameter(Mandatory)][string] $Uri)

    $next = $Uri
    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        if ($page.PSObject.Properties.Name -contains 'value') { $page.value } else { $page }
        $next = ($page.PSObject.Properties.Name -contains '@odata.nextLink') ? $page.'@odata.nextLink' : $null
    }
}

function Test-RegistryAccess {
    # Fail fast with an actionable message instead of a raw Graph error.
    try {
        Invoke-MgGraphRequest -Method GET -Uri "$($script:BaseUri)?`$top=1" -OutputType PSObject | Out-Null
        return
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match '402|payment|license') {
            throw "The agent registry rejected the request as unlicensed. The Package Management API requires a Microsoft Agent 365 license on the signing-in account. Details: $msg"
        }
        if ($msg -match '403|Forbidden|Authorization') {
            throw "Access denied to the agent registry. Confirm the account holds AI Administrator or Global Administrator, that CopilotPackages.Read.All was consented, and that a Microsoft Agent 365 license is assigned. Details: $msg"
        }
        if ($msg -match '404|NotFound') {
            throw "The agent registry endpoint was not found at $($script:BaseUri). This API is Global-service only (not GCC High/DoD/21Vianet) and may not yet be enabled in this tenant. Details: $msg"
        }
        throw
    }
}

#endregion

#region ---------- registry enumeration ----------

function Get-AgentPackages {
    # Host filter only. The 'platform' filter is deliberately NOT sent to the
    # server: the docs name the filter values as 'Copilot Studio' and
    # 'Microsoft 365 Copilot Agent Builder', but Microsoft's own example
    # responses show platform values of 'web' and 'teams'. Sending a guessed
    # literal risks a silent zero-result response, which is indistinguishable
    # from "you have no Agent Builder agents". Platform is matched client-side
    # below, and every distinct value seen is reported so you can verify it.
    $query   = '?$filter=' + [uri]::EscapeDataString("supportedHosts/any(h:h eq 'Copilot')")
    $fetched = @()

    try {
        $fetched = @(Invoke-GraphPaged -Uri ($script:BaseUri + $query))
        if (-not $fetched.Count) {
            Write-Verbose "Host filter returned nothing; retrying unfiltered."
            $fetched = @(Invoke-GraphPaged -Uri $script:BaseUri) | Where-Object {
                (@(Get-SafeProperty -Object $_ -Name 'supportedHosts') -join ',') -match '(?i)copilot'
            }
        }
    }
    catch {
        Write-Verbose "Server-side filter rejected ($($_.Exception.Message)); retrieving unfiltered."
        $fetched = @(Invoke-GraphPaged -Uri $script:BaseUri) | Where-Object {
            (@(Get-SafeProperty -Object $_ -Name 'supportedHosts') -join ',') -match '(?i)copilot'
        }
    }

    $fetched = @($fetched)

    # Surface the real platform values so -AgentBuilderOnly can be sanity-checked
    # against this tenant rather than trusted blindly.
    $platforms = @($fetched | ForEach-Object {
        ($_.PSObject.Properties.Name -contains 'platform') ? $_.platform : '(no platform property)'
    } | Select-Object -Unique)
    $script:PlatformValuesSeen = $platforms
    Write-Verbose "Distinct platform values returned: $($platforms -join ' | ')"

    if ($AgentBuilderOnly) {
        # Loose match tolerates 'Microsoft 365 Copilot Agent Builder',
        # 'AgentBuilder', 'agent_builder', and similar spellings.
        $matched = @($fetched | Where-Object {
            ($_.PSObject.Properties.Name -contains 'platform') -and $_.platform -match '(?i)agent[\s_-]*builder'
        })
        if (-not $matched.Count) {
            Write-Warning (@"
-AgentBuilderOnly matched 0 of $($fetched.Count) agent(s).
Platform values actually returned by this tenant: $($platforms -join ' | ')
Rather than report an empty result that looks like "no Agent Builder agents exist",
ALL $($fetched.Count) Copilot agent(s) are being returned. Inspect the Platform column
to see which value denotes Agent Builder in your tenant.
"@)
            return $fetched
        }
        return $matched
    }
    return $fetched
}

function Get-PackageDetail {
    param([Parameter(Mandatory)][string] $PackageId)
    Invoke-MgGraphRequest -Method GET -Uri "$($script:BaseUri)/$PackageId" -OutputType PSObject
}

#endregion

$script:OwnerLookupFailures = 0
# Count of capabilities that reference ONLY linked SharePoint content (no uploads).
$script:LinkedOnlyCount = 0

#region ---------- owner + date helpers ------------------------------------

# ownerId on a package detail is a directory object id. Confirmed:
# GET /v1.0/users/{id} returns displayName, userPrincipalName and mail for the
# ownerId the registry returns.
$script:UserCache = @{}

function Resolve-AgentOwner {
    param([string] $OwnerId)

    if (-not $OwnerId) { return $null }
    if ($script:UserCache.ContainsKey($OwnerId)) { return $script:UserCache[$OwnerId] }

    $result = [pscustomobject]@{ DisplayName = $null; Upn = $null; Id = $OwnerId }
    try {
        $u = Invoke-MgGraphRequest -Method GET -OutputType PSObject `
            -Uri "https://graph.microsoft.com/v1.0/users/$OwnerId`?`$select=displayName,userPrincipalName,mail"
        $result.DisplayName = $u.displayName
        $result.Upn         = $u.userPrincipalName ?? $u.mail
    }
    catch {
        # Deleted user, service principal, or missing User.Read.All. Report the
        # raw id rather than dropping the owner entirely.
        Write-Verbose "Owner lookup failed for $OwnerId : $($_.Exception.Message)"
        $script:OwnerLookupFailures++
    }
    $script:UserCache[$OwnerId] = $result
    return $result
}

function Format-AgentDate {
    param($Value)
    if (-not $Value) { return $null }
    # Normalise to sortable UTC. Keeps CSV text sortable and unambiguous, rather
    # than letting Excel reinterpret a raw string in local format.
    try   { return ([datetime]$Value).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') + 'Z' }
    catch { return [string]$Value }
}

#endregion

#region ---------- admin center deep links ---------------------------------

# Admin center agent list URL, confirmed from the browser address bar:
#     https://admin.cloud.microsoft/?#/agents/all
# Note the '/?#/' - a query marker BEFORE the hash route.
#
# Selecting an agent opens a detail panel without changing the route, so there is
# no true per-agent URL. A confirmed-working equivalent is the list view filtered
# and searched down to one agent:
#     .../?#/agents/all?createdIn=Agent+Builder+in+Microsoft+365+Copilot&search=contract+reviewer
# Spaces are '+'. Search appears to be case-insensitive.
$script:AgentListUrl = 'https://admin.cloud.microsoft/?#/agents/all'

# The 'createdIn' value is the portal's own label and does NOT match the
# 'platform' string the Graph API returns ('Microsoft 365 Copilot Agent Builder').
# Only the Agent Builder literal has been observed, so the filter is emitted only
# when the run is scoped to Agent Builder; otherwise search alone is used, rather
# than guessing a label for Copilot Studio or other platforms.
$script:AgentBuilderCreatedIn = 'Agent+Builder+in+Microsoft+365+Copilot'

function ConvertTo-AdminSearchTerm {
    param([string] $Text)
    if (-not $Text) { return $null }
    # Form-style encoding: percent-encode, then spaces as '+'.
    return ([uri]::EscapeDataString($Text) -replace '%20', '+')
}

function Get-AdminCenterUrl {
    param([string] $AgentName)

    $query = @()
    if ($AgentBuilderOnly) { $query += "createdIn=$($script:AgentBuilderCreatedIn)" }

    $term = ConvertTo-AdminSearchTerm -Text $AgentName
    if ($term) { $query += "search=$term" }

    if ($query.Count) { return $script:AgentListUrl + '?' + ($query -join '&') }
    return $script:AgentListUrl
}

#endregion

#region ---------- knowledge source analysis ----------

# Capability names per the declarative agent manifest schema. Those that resolve
# to concrete stored content are treated as attachment-style knowledge.
# Knowledge items are identified by their stable unique_id GUID. Resolving those
# GUIDs to file names requires access to the SharePoint Embedded container that
# Agent Builder stores uploads in, which ordinary admin permissions do not grant
# (every attempt returned 403). File names are read from the admin center UI
# instead - see AdminCenterUrl in the output.
function Format-KnowledgeItem {
    param($ItemRef)
    $id = Get-SafeProperty -Object $ItemRef -Name 'unique_id'
    if ($id) { return "unique_id:$id" }
    return ($ItemRef | ConvertTo-Json -Compress -Depth 5)
}

$script:ToolCapabilities       = @('CodeInterpreter', 'GraphicArt')

function Get-AgentDefinitions {
    param([Parameter(Mandatory)] $Detail)

    if ($Detail.PSObject.Properties.Name -notcontains 'elementDetails') { return @() }

    $out = @()
    foreach ($group in @($Detail.elementDetails)) {
        if ($group.elementType -notmatch $script:AgentElementTypePattern) { continue }
        foreach ($el in @($group.elements)) {
            if (-not $el -or $el.PSObject.Properties.Name -notcontains 'definition') { continue }
            $raw = $el.definition
            $parsed = $null
            # 'definition' is documented as a JSON *string*; tolerate an object too.
            if ($raw -is [string]) {
                try { $parsed = $raw | ConvertFrom-Json } catch { Write-Verbose "Unparseable definition on element $($el.id)." }
            }
            else { $parsed = $raw }

            $out += [pscustomobject]@{
                ElementId   = $el.id
                ElementType = $group.elementType
                Raw         = $raw
                Definition  = $parsed
            }
        }
    }
    return $out
}

function Get-KnowledgeSources {
    param([Parameter(Mandatory)] $Definition)

    # Returns $null when capabilities are absent - meaning "unknown", not "none".
    if (-not $Definition) { return $null }
    if ($Definition.PSObject.Properties.Name -notcontains 'capabilities') { return $null }

    $rows = @()
    foreach ($cap in @($Definition.capabilities)) {
        if (-not $cap -or $cap.PSObject.Properties.Name -notcontains 'name') { continue }
        $capName = $cap.name
        $items   = @()
        $kind    = 'Other'
        # True only for uploaded file attachments; set inside the branches below.
        $isAttachmentCapability = $false
        # Count of UPLOADED files only. $items may also carry linked entries when
        # -IncludeLinked is set, so item count must never be used as the
        # attachment count - that conflated the two and over-reported.
        $attachmentItemCount = 0

        switch -Regex ($capName) {
            'OneDriveAndSharePoint' {
                # Only files UPLOADED into the agent ('x-is_embedded') count as
                # attachments. Linked sites/libraries and items_by_url keep their
                # own SharePoint permissions, so they are not an exposure of the
                # same kind and are excluded unless -IncludeLinked is passed.
                $embeddedItems = @()
                $linkedItems   = @()

                if ($cap.PSObject.Properties.Name -contains 'items_by_url') {
                    foreach ($i in @($cap.items_by_url)) {
                        if ($i.PSObject.Properties.Name -contains 'url') { $linkedItems += $i.url }
                    }
                }
                if ($cap.PSObject.Properties.Name -contains 'items_by_sharepoint_ids') {
                    foreach ($i in @($cap.items_by_sharepoint_ids)) {
                        $isEmbedded = ($i.PSObject.Properties.Name -contains 'x-is_embedded') -and $i.'x-is_embedded'
                        if ($isEmbedded) { $embeddedItems += (Format-KnowledgeItem -ItemRef $i) }
                        else             { $linkedItems   += (Format-KnowledgeItem -ItemRef $i) }
                    }
                }

                if ($IncludeLinked) {
                    $items = @($embeddedItems) + @($linkedItems | ForEach-Object { "$_ (linked)" })
                }
                else {
                    $items = @($embeddedItems)
                }

                if (-not $embeddedItems.Count -and $linkedItems.Count) { $script:LinkedOnlyCount++ }

                $kind = if ($embeddedItems.Count) { "Uploaded file attachment(s): $($embeddedItems.Count)" }
                        elseif ($linkedItems.Count) { "Linked SharePoint content only - not an uploaded attachment ($($linkedItems.Count) item(s))" }
                        else { 'SharePoint-OneDrive, tenant-wide (nothing pinned)' }

                # Attachment flag and count key off uploaded files ONLY.
                $isAttachmentCapability = $embeddedItems.Count -gt 0
                $attachmentItemCount    = $embeddedItems.Count
            }
            'GraphConnectors' {
                if ($cap.PSObject.Properties.Name -contains 'connections') {
                    foreach ($i in @($cap.connections)) { $items += $i.connection_id }
                }
                $kind = 'Copilot connector'
            }
            'Dataverse' {
                if ($cap.PSObject.Properties.Name -contains 'knowledge_sources') {
                    foreach ($i in @($cap.knowledge_sources)) { $items += ($i.host_name ?? ($i | ConvertTo-Json -Compress -Depth 5)) }
                }
                $kind = 'Dataverse knowledge'
                $isAttachmentCapability = $items.Count -gt 0
                $attachmentItemCount    = $items.Count
            }
            'WebSearch'      { $kind = 'Web search' }
            'TeamsMessages'  { $kind = 'Teams messages' }
            'People'         { $kind = 'People' }
            'CodeInterpreter'{ $kind = 'Tool (not knowledge)' }
            'GraphicArt'     { $kind = 'Tool (not knowledge)' }
        }

        $rows += [pscustomobject]@{
            Capability   = $capName
            SourceKind   = $kind
            ItemCount       = $items.Count
            AttachmentCount = $attachmentItemCount
            Items           = ($items -join '; ')
            IsAttachment    = $isAttachmentCapability
            IsTool       = ($script:ToolCapabilities -contains $capName)
        }
    }
    return $rows
}

#endregion

#region ---------- main ----------

Connect-Registry
Test-RegistryAccess

$packages = @(Get-AgentPackages) | Group-Object -Property id | ForEach-Object { $_.Group[0] }
$packages = @($packages)
Write-Verbose "Agent packages returned: $($packages.Count)"

if (-not $packages.Count) {
    Write-Warning "The registry returned no agent packages. If you expect agents here, re-run with -Verbose to see what the registry returned."
    return
}

$report      = @()
$noCapability = @()

foreach ($pkg in $packages) {
    try   { $detail = Get-PackageDetail -PackageId $pkg.id }
    catch { Write-Warning "Could not read details for '$($pkg.displayName)': $($_.Exception.Message)"; continue }

    # createdDateTime and ownerId appear on the package DETAIL response; the list
    # response carries lastModifiedDateTime. Read detail first, fall back to list.
    $prop = {
        param($Name)
        if (($detail.PSObject.Properties.Name -contains $Name) -and ($null -ne $detail.$Name)) { return $detail.$Name }
        if ($pkg.PSObject.Properties.Name    -contains $Name) { return $pkg.$Name }
        return $null
    }

    $owner = Resolve-AgentOwner -OwnerId (& $prop 'ownerId')

    $common = @{
        Agent        = $pkg.displayName
        PackageId    = $pkg.id
        Platform     = ($pkg.PSObject.Properties.Name -contains 'platform')  ? $pkg.platform  : $null
        Owner        = $owner ? ($owner.DisplayName ?? "(unresolved: $($owner.Id))") : $null
        OwnerUpn     = $owner ? $owner.Upn : $null
        LastModifiedDateTime = Format-AgentDate (& $prop 'lastModifiedDateTime')
        AvailableTo  = ($pkg.PSObject.Properties.Name -contains 'availableTo') ? $pkg.availableTo : $null
        AdminCenterUrl = (Get-AdminCenterUrl -AgentName $pkg.displayName)
    }

    $defs = @(Get-AgentDefinitions -Detail $detail)

    # Dump the ENTIRE detail response, not just parsed elements. The previous
    # version only printed when elements were found, so the one case that needed
    # inspection - nothing parsed - printed nothing at all.
    if ($ShowRawDefinition) {
        Write-Host "`n=== $($pkg.displayName) :: full package detail response ===" -ForegroundColor DarkCyan
        Write-Host ($detail | ConvertTo-Json -Depth 12)
        Write-Host "=== end $($pkg.displayName) ===" -ForegroundColor DarkCyan
    }

    if (-not $defs.Count) {
        # The list payload carries elementTypes. If it says this IS a declarative
        # agent but the detail response carried no parseable element, that is a
        # RETRIEVAL GAP, not evidence the agent has no knowledge. Never report it
        # as a clean negative.
        $declaredTypes = ($pkg.PSObject.Properties.Name -contains 'elementTypes') ? (@($pkg.elementTypes) -join ',') : ''
        $looksLikeAgent = $declaredTypes -match $script:AgentElementTypePattern
        $hasElementDetails = ($detail.PSObject.Properties.Name -contains 'elementDetails')

        if ($looksLikeAgent -or -not $hasElementDetails) {
            $why = -not $hasElementDetails `
                ? 'detail response contained no elementDetails property' `
                : "elementDetails present but held no parseable agent element (elementTypes reported: $declaredTypes)"
            $noCapability += $pkg.displayName
            $report += [pscustomobject]($common + @{
                Capability = '(not returned)'
                SourceKind = "UNKNOWN - $why"
                ItemCount = 0; AttachmentCount = 0; Items = ''; HasKnowledge = $null; HasAttachments = $null })
        }
        else {
            $report += [pscustomobject]($common + @{
                Capability = '(no agent element)'
                SourceKind = "Not an agent package (elementTypes: $declaredTypes)"
                ItemCount = 0; AttachmentCount = 0; Items = ''; HasKnowledge = $false; HasAttachments = $false })
        }
        continue
    }

    foreach ($d in $defs) {
        $sources = Get-KnowledgeSources -Definition $d.Definition

        if ($null -eq $sources) {
            # Capabilities were not returned - report as UNKNOWN, never as "none".
            $noCapability += $pkg.displayName
            $report += [pscustomobject]($common + @{
                Capability = '(not returned)'; SourceKind = 'UNKNOWN - API did not return capability detail'
                ItemCount = 0; AttachmentCount = 0; Items = ''; HasKnowledge = $null; HasAttachments = $null })
            continue
        }

        $knowledge = @($sources | Where-Object { -not $_.IsTool })
        if (-not $knowledge.Count) {
            $report += [pscustomobject]($common + @{
                Capability = '(none)'; SourceKind = 'No knowledge sources configured'
                ItemCount = 0; AttachmentCount = 0; Items = ''; HasKnowledge = $false; HasAttachments = $false })
            continue
        }

        foreach ($s in $knowledge) {
            $report += [pscustomobject]($common + @{
                Capability     = $s.Capability
                SourceKind     = $s.SourceKind
                ItemCount       = $s.ItemCount
                AttachmentCount = $s.AttachmentCount
                Items          = $s.Items
                HasKnowledge   = $true
                HasAttachments = $s.IsAttachment
            })
        }
    }
}

# Keep the unfiltered set: the per-agent rollup is built from it so that an
# agent's full capability list survives filtering of the detail rows.
$fullReport = @($report)

if ($OnlyWithAttachments) {
    # Retain UNKNOWN rows. Filtering on '-eq $true' alone would silently drop
    # agents whose capability data could not be read, and an agent that vanishes
    # from an attachment report reads as "has none" - the exact wrong conclusion.
    $report = @($report | Where-Object { $_.HasAttachments -eq $true -or $null -eq $_.HasAttachments })

    $droppedUnknown = @($report | Where-Object { $null -eq $_.HasAttachments })
    if ($droppedUnknown.Count) {
        Write-Warning "$($droppedUnknown.Count) row(s) with UNKNOWN capability data are included in the attachment report. They are NOT confirmed attachments - review them individually."
    }
}

if ($CsvPath) {
    $report | Select-Object Agent, Platform, Owner, OwnerUpn, LastModifiedDateTime, AvailableTo, Capability, SourceKind, ItemCount, AttachmentCount, Items, HasKnowledge, HasAttachments, AdminCenterUrl, PackageId |
        Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report written to $CsvPath" -ForegroundColor Green
}

# --- per-agent rollup ------------------------------------------------------
# The detail report is one row PER KNOWLEDGE SOURCE, so an agent with two
# capabilities legitimately appears twice. Roll up to one row per agent so the
# attachment answer is unambiguous.
$summary = $fullReport | Group-Object -Property PackageId | ForEach-Object {
    $rows  = $_.Group
    $first = $rows[0]
    $attachRows    = @($rows | Where-Object { $_.HasAttachments -eq $true })
    $anyAttach     = $attachRows.Count -gt 0
    $anyUnknown    = @($rows | Where-Object { $null -eq $_.HasKnowledge }).Count -gt 0
    # Show uploaded files only here; '(linked)' entries are excluded so the
    # summary column matches AttachmentCount.
    $attachedItems = (@($attachRows | ForEach-Object { $_.Items }) -split '; ' |
        Where-Object { $_ -and $_ -notmatch '\(linked\)$' }) -join '; '

    # Count uploaded files first. This MUST precede the admin-center comparison
    # below, which reads $attachCount - computing it afterwards left the
    # comparison reading a stale value from the previous agent.
    # Sum defensively: under Set-StrictMode -Version Latest, reading .Sum off a
    # pipeline that produced no output throws PropertyNotFound.
    $attachCount = 0
    if ($attachRows.Count) {
        $measured = $attachRows | Measure-Object -Property AttachmentCount -Sum
        if ($null -ne $measured -and $null -ne $measured.Sum) { $attachCount = [int]$measured.Sum }
    }

    [pscustomobject]@{
        Agent          = $first.Agent
        Platform       = $first.Platform
        Owner          = $first.Owner
        OwnerUpn       = $first.OwnerUpn
        LastModifiedDateTime = $first.LastModifiedDateTime
        HasAttachments = $anyAttach ? $true : ($anyUnknown ? $null : $false)
        AttachmentCount= $attachCount
        LinkedCount    = @($rows | ForEach-Object { $_.Items } | ForEach-Object { $_ -split '; ' } | Where-Object { $_ -match '\(linked\)$' }).Count
        Capabilities   = (@($rows | ForEach-Object { $_.Capability }) -join ', ')
        AttachedItems  = $attachedItems
        AdminCenterUrl = (Get-AdminCenterUrl -AgentName $first.Agent)
        PackageId      = $first.PackageId
    }
}

if ($OnlyWithAttachments) {
    $summary = @($summary | Where-Object { $_.HasAttachments -eq $true -or $null -eq $_.HasAttachments })
}

if ($CsvPath) {
    $summaryPath = [IO.Path]::ChangeExtension($CsvPath, $null) + 'summary.csv'
    $summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
    Write-Host "Per-agent summary written to $summaryPath" -ForegroundColor Green
}

Write-Host "`nPer-agent summary (one row per agent):" -ForegroundColor Cyan
$summary | Select-Object Agent, Owner, LastModifiedDateTime, HasAttachments, AttachmentCount, AttachedItems | Format-Table -AutoSize -Wrap

Write-Host "`nPlatform values seen in this tenant (use these to confirm the Agent Builder filter):" -ForegroundColor Cyan
foreach ($p in @($script:PlatformValuesSeen)) { Write-Host "  $p" -ForegroundColor DarkGray }

Write-Host "`nAgents with attachment-style knowledge sources:" -ForegroundColor Cyan
$attach = @($report | Where-Object { $_.HasAttachments -eq $true })
if ($attach.Count) {
    $attach | Select-Object Agent, Capability, SourceKind, ItemCount, Items | Format-Table -AutoSize -Wrap
}
else {
    Write-Host "  (none found)" -ForegroundColor DarkGray
}

if ($noCapability.Count) {
    Write-Warning (@"
Capability detail was NOT returned for $($noCapability.Count) agent(s): $((($noCapability | Select-Object -Unique | Select-Object -First 10) -join ', ')).
These are reported as UNKNOWN, not as "no knowledge sources" - do not read them as clean.
Microsoft's published example for this endpoint shows a truncated agent definition, so the
registry may not expose manifest capabilities for these agents.

NEXT STEP: re-run with -ShowRawDefinition to dump the full package detail response, and
If an agent you KNOW has attached knowledge appears here, the registry is not surfacing its
knowledge sources for that agent.
"@)
}

if ($script:LinkedOnlyCount -and -not $IncludeLinked) {
    Write-Host "`nNote: $($script:LinkedOnlyCount) agent capability/capabilities reference linked SharePoint content" -ForegroundColor DarkGray
    Write-Host "with no uploaded files. Excluded by design (permissions are inherited). Use -IncludeLinked to list them." -ForegroundColor DarkGray
}

if ($script:OwnerLookupFailures) {
    Write-Warning @"
$($script:OwnerLookupFailures) owner lookup(s) failed, so Owner shows a raw directory id and
OwnerUpn is empty.

Most common cause: User.Read.All is not on the current token. Graph PowerShell reuses a cached
session, so a session created before this permission was requested will not have it.
  FIX: Disconnect-MgGraph, then re-run and accept the consent prompt.
       (App-only runs need User.Read.All as an APPLICATION permission with admin consent.)

Other causes: the owner account was deleted, or the owner is a non-user identity.
"@
}

$report
#endregion
