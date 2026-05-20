<#
.SYNOPSIS
    Blocks all Copilot agents in the M365 Agent Registry via the Microsoft Graph API.

.DESCRIPTION
    Uses the Microsoft Graph Package Management API (beta) to enumerate every agent
    registered in the Microsoft 365 Copilot Agent Registry and block any that are
    not already blocked. Agents that should be excluded from blocking are listed 
    starting in line 59. Requires Powershell 7.

    API base  : https://graph.microsoft.com/beta/copilot/admin/catalog/packages
    Scope     : CopilotPackages.ReadWrite.All  (Delegated — work/school account)
    License   : Microsoft Agent 365 required on the tenant

    The block endpoint returns 204 No Content on success. Agents that are already
    blocked are skipped. Results are summarised at the end.

.PARAMETER TenantId
    Optional. The Entra tenant ID or verified domain to authenticate against.
    If omitted, the home tenant of the signed-in account is used.

.PARAMETER AgentTypes
    One or more element type strings to filter by.
    Defaults to @('DeclarativeAgent','CustomEngineAgent') which targets agents only.
    Pass @('*') to include every package type (bots, add-ins, etc.) that supports Copilot.

.PARAMETER WhatIf
    Dry-run mode — lists the agents that would be blocked without actually blocking them.

.EXAMPLE
    .\Block-AllCopilotAgents.ps1

.EXAMPLE
    .\Block-AllCopilotAgents.ps1 -TenantId contoso.onmicrosoft.com -WhatIf

.NOTES
    API reference:
      https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/package/overview
    Required admin role: AI Administrator or Global Administrator
#>

#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [int] $RetryAfterSeconds = 30,

    [Parameter()]
    [int] $MaxRetries = 3
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Agents that must never be blocked ────────────────────────────────────────
$EXCLUDED_AGENTS = @(
    'Excel (Agent)',
    'PowerPoint (Agent)',
    'Word (Agent)',
    'Cowork (Frontier)',
    'Researcher',
    'Analyst'
)

# ── Constants ──────────────────────────────────────────────────────────────────
$GRAPH_BETA       = 'https://graph.microsoft.com/beta'
$PACKAGES_URL     = "$GRAPH_BETA/copilot/admin/catalog/packages"
$REQUIRED_SCOPE   = 'CopilotPackages.ReadWrite.All'

# ── Module check ───────────────────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
    Write-Host "[setup] Installing Microsoft.Graph.Authentication ..." -ForegroundColor Cyan
    Install-Module -Name 'Microsoft.Graph.Authentication' `
                   -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
}
Import-Module -Name 'Microsoft.Graph.Authentication' -ErrorAction Stop

# ── Connect ────────────────────────────────────────────────────────────────────
$connectParams = @{
    Scopes    = @($REQUIRED_SCOPE)
    NoWelcome = $true
}
if ($TenantId) { $connectParams['TenantId'] = $TenantId }

Write-Host "[auth] Connecting to Microsoft Graph (scope: $REQUIRED_SCOPE) ..." -ForegroundColor Cyan
Connect-MgGraph @connectParams
Write-Host "[auth] Connected as: $((Get-MgContext).Account)" -ForegroundColor Green

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
            $status = $_.Exception.Response?.StatusCode?.value__ ?? 0
            if (($status -in 429, 503) -and $attempt -lt $Retries) {
                $wait = $RetryAfterSeconds
                # honour Retry-After header when present
                $retryHeader = $_.Exception.Response?.Headers?['Retry-After']
                if ($retryHeader) { $wait = [int]$retryHeader }
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

# ── Build list URL ─────────────────────────────────────────────────────────────
# Filter to packages that surface in Copilot — this is what the API docs define
# as "all agents". The elementTypes field is often unpopulated so we do not
# filter on it; instead we post-filter in PowerShell using -AgentTypes if needed.
$listUrl = "${PACKAGES_URL}?`$filter=supportedHosts/any(h:h eq 'Copilot')"

# ── Enumerate all agents (page through results) ────────────────────────────────
Write-Host "[fetch] Enumerating agents from the Agent Registry ..." -ForegroundColor Cyan

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
            ElementTypes = ($item.elementTypes -join ', ')
        })
    }
    $pageUrl = $page.PSObject.Properties['@odata.nextLink']?.Value
    Write-Verbose "[fetch] Page complete — $($agents.Count) agents so far ..."
} while ($pageUrl)

Write-Host "[fetch] Found $($agents.Count) agent(s) total." -ForegroundColor Cyan

if ($agents.Count -eq 0) {
    Write-Host "[done] No agents found. Nothing to do." -ForegroundColor Yellow
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    exit 0
}

# ── Display what was found ─────────────────────────────────────────────────────
$agents | Format-Table -AutoSize -Property Id,DisplayName, Type, Publisher, ElementTypes, IsBlocked

# ── Block each unblocked agent ─────────────────────────────────────────────────
$stats = @{ AlreadyBlocked = 0; Excluded = 0; Blocked = 0; WouldBlock = 0; Failed = 0 }
$failures = [System.Collections.Generic.List[pscustomobject]]::new()

foreach ($agent in $agents) {
    if ($agent.IsBlocked) {
        Write-Verbose "[skip] Already blocked: $($agent.DisplayName)"
        $stats.AlreadyBlocked++
        continue
    }

    if ($EXCLUDED_AGENTS -contains $agent.DisplayName) {
        Write-Host "[exempt]  $($agent.DisplayName)" -ForegroundColor DarkYellow
        $stats.Excluded++
        continue
    }

    $blockUrl = "$PACKAGES_URL/$($agent.Id)/block"
    $label    = "$($agent.DisplayName) [$($agent.Id)]"

    if ($PSCmdlet.ShouldProcess($label, 'Block Copilot agent')) {
        try {
            Invoke-GraphWithRetry -Method POST -Uri $blockUrl | Out-Null
            Write-Host "[blocked] $($agent.DisplayName)" -ForegroundColor Green
            $stats.Blocked++
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-Warning "[failed]  $($agent.DisplayName) — $errMsg"
            $failures.Add([pscustomobject]@{
                Id    = $agent.Id
                Name  = $agent.DisplayName
                Error = $errMsg
            })
            $stats.Failed++
        }
    }
    else {
        # -WhatIf path
        Write-Host "[whatif]  Would block: $($agent.DisplayName)" -ForegroundColor Yellow
        $stats.WouldBlock++
    }
}

# ── Summary ────────────────────────────────────────────────────────────────────
$divider = '─' * 52
Write-Host ''
Write-Host $divider -ForegroundColor Cyan
Write-Host ' Block-AllCopilotAgents — Summary' -ForegroundColor Cyan
Write-Host $divider -ForegroundColor Cyan
Write-Host "  Total agents found   : $($agents.Count)"
Write-Host "  Already blocked      : $($stats.AlreadyBlocked)" -ForegroundColor DarkGray
Write-Host "  Exempt (protected)   : $($stats.Excluded)" -ForegroundColor DarkYellow

if ($WhatIfPreference) {
    Write-Host "  Would be blocked     : $($stats.WouldBlock)" -ForegroundColor Yellow
}
else {
    Write-Host "  Successfully blocked : $($stats.Blocked)" -ForegroundColor Green
    if ($stats.Failed -gt 0) {
        Write-Host "  Failed               : $($stats.Failed)" -ForegroundColor Red
        Write-Host ''
        Write-Host '  Failed agents:' -ForegroundColor Red
        $failures | Format-Table -AutoSize -Property Name, Id, Error
    }
}
Write-Host $divider -ForegroundColor Cyan

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
