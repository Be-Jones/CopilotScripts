<#
.SYNOPSIS
    Blocks all Copilot agents in the M365 Agent Registry via the Microsoft Graph API.

.DESCRIPTION
    Uses the Microsoft Graph Package Management API (beta) to enumerate every agent
    registered in the Microsoft 365 Copilot Agent Registry and block any that are
    not already blocked. Agents that should be protected from blocking must be
    listed in a JSON reference file. Requires Powershell 7.

    API base  : https://graph.microsoft.com/beta/copilot/admin/catalog/packages
    Scope     : CopilotPackages.ReadWrite.All  (Delegated — work/school account)
    License   : Microsoft Agent 365 required on the tenant

    The block endpoint returns 204 No Content on success. Agents that are already
    blocked are skipped. Results are summarised at the end.

.PARAMETER TenantId
    Optional. The Entra tenant ID or verified domain to authenticate against.
    If omitted, the home tenant of the signed-in account is used.

.PARAMETER AllowList
    Required. Path to a JSON file containing agents that must NOT be blocked.
    The script matches allowed agents by package id. Display names are informational.
    A list of agents can be exported directly from https://admin.microsoft365.com/#/agents/all.
    From the .csv export, use the Displayname and Title ID columns in the allowedAgents.json file.

.PARAMETER Interactive
    Optional. Writes human-readable progress and table output for manual runs.
    Automated runs emit a structured summary object and warnings/errors only.

.PARAMETER BlockAgentBuilderAgents
    Optional. Includes Microsoft 365 Copilot Agent Builder agents in block attempts.
    By default, Agent Builder agents are excluded from blocking.

.EXAMPLE
    .\Block-AllCopilotAgents.ps1 -AllowList .\allowedAgents.json

.EXAMPLE
    .\Block-AllCopilotAgents.ps1 -TenantId contoso.onmicrosoft.com -AllowList .\allowedAgents.json -Interactive

.EXAMPLE
    .\Block-AllCopilotAgents.ps1 -AllowList .\allowedAgents.json -BlockAgentBuilderAgents

.NOTES
    API reference:
      https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/package/overview
    Required admin role: AI Administrator or Global Administrator
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter()]
    [string] $TenantId,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string] $AllowList,

    [Parameter()]
    [switch] $Interactive,

    [Parameter()]
    [switch] $BlockAgentBuilderAgents,

    [Parameter()]
    [ValidateRange(0, 86400)]
    [int] $RetryAfterSeconds = 30,

    [Parameter()]
    [ValidateRange(0, 20)]
    [int] $MaxRetries = 3
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Constants ──────────────────────────────────────────────────────────────────
$GRAPH_BETA       = 'https://graph.microsoft.com/beta'
$PACKAGES_URL     = "$GRAPH_BETA/copilot/admin/catalog/packages"
$REQUIRED_SCOPE   = 'CopilotPackages.ReadWrite.All'

# ── Helper: interactive progress output ───────────────────────────────────────
function Write-InteractiveMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Message,

        [Parameter()]
        [System.ConsoleColor] $ForegroundColor = [System.ConsoleColor]::White
    )

    if ($Interactive) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
}

# ── Helper: allowed agents reference file ─────────────────────────────────────
function Get-AllowedAgentIds {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        Write-Error -Category ObjectNotFound -ErrorAction Stop -Message "Allowed agents file not found: $Path"
    }

    try {
        $allowedAgentsDocument = Get-Content -Path $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Error -Category ParserError -ErrorAction Stop -Message "Allowed agents file is not valid JSON: $Path. $($_.Exception.Message)"
    }

    $allowedAgentsProperty = $allowedAgentsDocument.PSObject.Properties['allowedAgents']
    if ($null -eq $allowedAgentsProperty) {
        Write-Error -Category InvalidData -ErrorAction Stop -Message "Allowed agents file must contain an 'allowedAgents' array. See allowedAgents.json for the expected format."
    }

    if ($null -ne $allowedAgentsProperty.Value -and $allowedAgentsProperty.Value -isnot [array]) {
        Write-Error -Category InvalidData -ErrorAction Stop -Message "Allowed agents file property 'allowedAgents' must be an array."
    }

    $allowedAgentIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $seenDisplayNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($allowedAgent in @($allowedAgentsProperty.Value)) {
        if ($null -eq $allowedAgent) { continue }

        $idProperty = $allowedAgent.PSObject.Properties['id']
        if ($null -eq $idProperty -or [string]::IsNullOrWhiteSpace([string]$idProperty.Value)) {
            Write-Error -Category InvalidData -ErrorAction Stop -Message "Each allowedAgents entry must include a non-empty 'id'."
        }

        $allowedAgentId = [string]$idProperty.Value
        if (-not $allowedAgentIds.Add($allowedAgentId)) {
            Write-Error -Category InvalidData -ErrorAction Stop -Message "Allowed agents file contains a duplicate id: $allowedAgentId"
        }

        $displayNameProperty = $allowedAgent.PSObject.Properties['displayName']
        if ($null -ne $displayNameProperty -and -not [string]::IsNullOrWhiteSpace([string]$displayNameProperty.Value)) {
            [void]$seenDisplayNames.Add([string]$displayNameProperty.Value)
        }
    }

    return [pscustomobject]@{
        Ids          = $allowedAgentIds
        DisplayNames = @($seenDisplayNames)
        Count        = $allowedAgentIds.Count
    }
}

# ── Module check ───────────────────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
    Write-Error -Category ResourceUnavailable -ErrorAction Stop -Message "Required PowerShell module 'Microsoft.Graph.Authentication' is not installed. Install it before running this script: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser. Installation help: https://www.powershellgallery.com/packages/Microsoft.Graph.Authentication"
}
Import-Module -Name 'Microsoft.Graph.Authentication' -ErrorAction Stop

# ── Helper: invoke with retry on 429 / 503 ────────────────────────────────────
function Invoke-GraphWithRetry {
    [CmdletBinding()]
    param(
        [string] $Method,
        [string] $Uri,
        [int]    $Retries = $MaxRetries
    )

    $attempt = 0
    do {
        try {
            return Invoke-MgGraphRequest -Method $Method -Uri $Uri
        }
        catch {
            $status = Get-GraphErrorStatusCode -ErrorRecord $_
            if (($status -in 429, 503) -and $attempt -lt $Retries) {
                $wait = Get-RetryDelaySeconds -Response $_.Exception.Response
                Write-Warning "[retry] HTTP $status — waiting ${wait}s before retry $($attempt + 1)/$Retries ..."
                Start-Sleep -Seconds $wait
                $attempt++
            }
            else {
                throw
            }
        }
    } while ($attempt -le $Retries)
}

function Get-RetryDelaySeconds {
    [CmdletBinding()]
    param(
        [Parameter()]
        $Response
    )

    $wait = $RetryAfterSeconds
    $retryHeader = $null
    if ($null -ne $Response -and $null -ne $Response.Headers) {
        $retryHeader = $Response.Headers['Retry-After']
    }

    if ($retryHeader) {
        $retryHeaderValue = @($retryHeader)[0]
        $retryAfterSecondsValue = 0

        if ([int]::TryParse([string]$retryHeaderValue, [ref]$retryAfterSecondsValue)) {
            $wait = $retryAfterSecondsValue
        }
        else {
            $retryAfterDate = [datetimeoffset]::MinValue
            if ([datetimeoffset]::TryParse([string]$retryHeaderValue, [ref]$retryAfterDate)) {
                $wait = [int][math]::Ceiling(($retryAfterDate - [datetimeoffset]::UtcNow).TotalSeconds)
            }
        }
    }

    return [math]::Max(0, [math]::Min($wait, 86400))
}

function Get-GraphErrorStatusCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord] $ErrorRecord
    )

    $responseProperty = $ErrorRecord.Exception.PSObject.Properties['Response']
    if ($null -eq $responseProperty) {
        return 0
    }

    $response = $responseProperty.Value
    if ($null -eq $response) {
        return 0
    }

    $statusCodeProperty = $response.PSObject.Properties['StatusCode']
    if ($null -eq $statusCodeProperty -or $null -eq $statusCodeProperty.Value) {
        return 0
    }

    $statusCode = $statusCodeProperty.Value
    $statusCodeValueProperty = $statusCode.PSObject.Properties['value__']
    if ($null -ne $statusCodeValueProperty) {
        return [int]$statusCodeValueProperty.Value
    }

    return [int]$statusCode
}

function Test-UnpublishedAgentBlockError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord] $ErrorRecord
    )

    $status = Get-GraphErrorStatusCode -ErrorRecord $ErrorRecord
    $message = @(
        $ErrorRecord.Exception.Message
        $ErrorRecord.ToString()
        ($ErrorRecord | Out-String)
    ) -join "`n"

    return ($status -eq 424 -or $message -match '(?i)(FailedDependency|Failed Dependency|draft|not published|unpublished)')
}

# ── Build list URL ─────────────────────────────────────────────────────────────
# Filter to packages that surface in Copilot. The elementTypes field is often
# unpopulated, so the script does not rely on it for the default target set.
$listUrl = "${PACKAGES_URL}?`$filter=supportedHosts/any(h:h eq 'Copilot')"

try {
    $allowedAgents = Get-AllowedAgentIds -Path $AllowList
    Write-InteractiveMessage -Message "[allow] Loaded $($allowedAgents.Count) protected agent id(s) from $AllowList" -ForegroundColor Cyan

    # ── Connect ────────────────────────────────────────────────────────────────
    $connectParams = @{
        Scopes    = @($REQUIRED_SCOPE)
        NoWelcome = $true
    }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }

    Write-InteractiveMessage -Message "[auth] Connecting to Microsoft Graph (scope: $REQUIRED_SCOPE) ..." -ForegroundColor Cyan
    Connect-MgGraph @connectParams | Out-Null
    Write-InteractiveMessage -Message "[auth] Connected as: $((Get-MgContext).Account)" -ForegroundColor Green

    # ── Enumerate all agents (page through results) ────────────────────────────
    Write-InteractiveMessage -Message '[fetch] Enumerating agents from the Agent Registry ...' -ForegroundColor Cyan

    $agents  = [System.Collections.Generic.List[pscustomobject]]::new()
    $pageUrl = $listUrl

    do {
        $page = Invoke-GraphWithRetry -Method GET -Uri $pageUrl
        foreach ($item in $page.value) {
            $agents.Add([pscustomobject]@{
                Id          = $item.id
                DisplayName = $item.displayName ?? '(no name)'
                Type        = $item.type
                IsBlocked   = [bool]($item.isBlocked)
                Publisher   = $item.publisher ?? ''
                Platform    = $item.platform ?? ''
                ElementTypes = ($item.elementTypes -join ', ')
            })
        }
        $pageUrl = $page.PSObject.Properties['@odata.nextLink']?.Value
        Write-Verbose "[fetch] Page complete — $($agents.Count) agents so far ..."
    } while ($pageUrl)

    Write-InteractiveMessage -Message "[fetch] Found $($agents.Count) agent(s) total." -ForegroundColor Cyan

    if ($Interactive -and $agents.Count -gt 0) {
        $agents | Format-Table -AutoSize -Property Id, DisplayName, Type, Publisher, Platform, ElementTypes, IsBlocked
    }

    # ── Block each unblocked agent ─────────────────────────────────────────────
    $stats = @{ AlreadyBlocked = 0; Excluded = 0; AgentBuilderExcluded = 0; Blocked = 0; DraftUnpublished = 0; Failed = 0 }
    $failures = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($agent in $agents) {
        if ($agent.IsBlocked) {
            Write-Verbose "[skip] Already blocked: $($agent.DisplayName)"
            $stats.AlreadyBlocked++
            continue
        }

        if ($allowedAgents.Ids.Contains($agent.Id)) {
            Write-InteractiveMessage -Message "[exempt]  $($agent.DisplayName)" -ForegroundColor DarkYellow
            $stats.Excluded++
            continue
        }

        if (-not $BlockAgentBuilderAgents -and $agent.Platform -eq 'Microsoft 365 Copilot Agent Builder') {
            Write-InteractiveMessage -Message "[exempt]  $($agent.DisplayName) (Agent Builder)" -ForegroundColor DarkYellow
            $stats.AgentBuilderExcluded++
            continue
        }

        $blockUrl = "$PACKAGES_URL/$($agent.Id)/block"
        try {
            Invoke-GraphWithRetry -Method POST -Uri $blockUrl | Out-Null
            Write-InteractiveMessage -Message "[blocked] $($agent.DisplayName)" -ForegroundColor Green
            $stats.Blocked++
        }
        catch {
            $errMsg = $_.Exception.Message
            if ((Test-UnpublishedAgentBlockError -ErrorRecord $_) -or $errMsg -match '(?i)(FailedDependency|Failed Dependency|draft|not published|unpublished)') {
                Write-Warning "[draft]   $($agent.DisplayName) — cannot block until the agent is published."
                $stats.DraftUnpublished++
            }
            else {
                Write-Warning "[failed]  $($agent.DisplayName) — $errMsg"
                $failures.Add([pscustomobject]@{
                    Id    = $agent.Id
                    Name  = $agent.DisplayName
                    Error = $errMsg
                })
                $stats.Failed++
            }
        }
    }

    $summary = [pscustomobject]@{
        TotalAgentsFound    = $agents.Count
        AllowedAgentsFile   = $AllowList
        AllowedAgentsLoaded = $allowedAgents.Count
        BlockAgentBuilderAgents = [bool]$BlockAgentBuilderAgents
        AlreadyBlocked      = $stats.AlreadyBlocked
        ExemptProtected     = $stats.Excluded
        AgentBuilderExcluded = $stats.AgentBuilderExcluded
        SuccessfullyBlocked = $stats.Blocked
        DraftUnpublished    = $stats.DraftUnpublished
        Failed              = $stats.Failed
        Failures            = @($failures)
    }

    if ($Interactive) {
        $divider = '─' * 52
        Write-Host ''
        Write-Host $divider -ForegroundColor Cyan
        Write-Host ' Block-AllCopilotAgents — Summary' -ForegroundColor Cyan
        Write-Host $divider -ForegroundColor Cyan
        Write-Host "  Total agents found   : $($summary.TotalAgentsFound)"
        Write-Host "  Allowed agents loaded: $($summary.AllowedAgentsLoaded)"
        Write-Host "  Already blocked      : $($summary.AlreadyBlocked)" -ForegroundColor DarkGray
        Write-Host "  Exempt (protected)   : $($summary.ExemptProtected)" -ForegroundColor DarkYellow
        Write-Host "  Agent Builder exempt : $($summary.AgentBuilderExcluded)" -ForegroundColor DarkYellow
        Write-Host "  Draft / unpublished  : $($summary.DraftUnpublished)" -ForegroundColor Yellow

        Write-Host "  Successfully blocked : $($summary.SuccessfullyBlocked)" -ForegroundColor Green
        if ($summary.Failed -gt 0) {
            Write-Host "  Failed               : $($summary.Failed)" -ForegroundColor Red
            Write-Host ''
            Write-Host '  Failed agents:' -ForegroundColor Red
            $failures | Format-Table -AutoSize -Property Name, Id, Error
        }
        Write-Host $divider -ForegroundColor Cyan
    }

    $summary
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
