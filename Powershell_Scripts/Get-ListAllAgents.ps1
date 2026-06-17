<#
.SYNOPSIS
    Lists Copilot agents from the M365 Agent Registry and emits allowlist-ready JSON.

.DESCRIPTION
    Uses the Microsoft Graph Package Management API (beta) to enumerate packages
    from the Microsoft 365 Copilot Agent Registry. By default, the script lists
    packages that surface in Copilot, which is the API-documented way to list all
    agents. The JSON output matches the allowedAgents.json shape consumed by
    Block-AllCopilotAgents.ps1.

    API base  : https://graph.microsoft.com/beta/copilot/admin/catalog/packages
    Scope     : CopilotPackages.Read.All  (Delegated - work/school account)
    License   : Microsoft Agent 365 required on the tenant

.PARAMETER TenantId
    Optional. The Entra tenant ID or verified domain to authenticate against.
    If omitted, the home tenant of the signed-in account is used.

.PARAMETER SupportedHost
    Optional. Filters packages by supported host. Defaults to Copilot.

.PARAMETER ElementType
    Optional. Filters packages by element type, such as DeclarativeAgent or
    CustomEngineAgent.

.PARAMETER LastModifiedAfter
    Optional. Filters packages last modified after the specified date/time.

.PARAMETER LastModifiedBefore
    Optional. Filters packages last modified before the specified date/time.

.PARAMETER Filter
    Optional. Raw OData $filter expression. If provided, it is used instead of
    SupportedHost, ElementType, LastModifiedAfter, and LastModifiedBefore.

.PARAMETER OutputPath
    Optional. Writes the allowlist-ready JSON to this path. The JSON is always
    emitted to the pipeline as well.

.PARAMETER Interactive
    Optional. Writes human-readable progress for manual runs. Automated runs emit
    JSON and warnings/errors only.

.EXAMPLE
    .\Get-ListAllAgents.ps1

.EXAMPLE
    .\Get-ListAllAgents.ps1 -OutputPath .\allowedAgents.generated.json

.EXAMPLE
    .\Get-ListAllAgents.ps1 -ElementType DeclarativeAgent -Interactive

.EXAMPLE
    .\Get-ListAllAgents.ps1 -Filter "supportedHosts/any(h:h eq 'Copilot') and lastModifiedDateTime gt 2026-01-01T00:00:00.0000000Z"

.NOTES
    API reference:
      https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/package/copilotpackages-list
    Required admin role: AI Administrator or Global Administrator
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [ValidateSet('Copilot', 'Outlook', 'Teams', 'M365', 'Word', 'Excel', 'PowerPoint')]
    [string] $SupportedHost = 'Copilot',

    [Parameter()]
    [ValidateSet('Bots', 'DeclarativeAgent', 'CustomEngineAgent', 'OfficeAddIns')]
    [string] $ElementType,

    [Parameter()]
    [Nullable[datetimeoffset]] $LastModifiedAfter,

    [Parameter()]
    [Nullable[datetimeoffset]] $LastModifiedBefore,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $Filter,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $OutputPath,

    [Parameter()]
    [switch] $Interactive,

    [Parameter()]
    [ValidateRange(0, 86400)]
    [int] $RetryAfterSeconds = 30,

    [Parameter()]
    [ValidateRange(0, 20)]
    [int] $MaxRetries = 3
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- Constants -----------------------------------------------------------------
$GRAPH_BETA     = 'https://graph.microsoft.com/beta'
$PACKAGES_URL   = "$GRAPH_BETA/copilot/admin/catalog/packages"
$REQUIRED_SCOPE = 'CopilotPackages.Read.All'

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

function ConvertTo-ODataDateTime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [datetimeoffset] $Value
    )

    return $Value.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ', [System.Globalization.CultureInfo]::InvariantCulture)
}

function New-CopilotPackagesFilter {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string] $SupportedHost,

        [Parameter()]
        [string] $ElementType,

        [Parameter()]
        [Nullable[datetimeoffset]] $LastModifiedAfter,

        [Parameter()]
        [Nullable[datetimeoffset]] $LastModifiedBefore,

        [Parameter()]
        [string] $RawFilter
    )

    if (-not [string]::IsNullOrWhiteSpace($RawFilter)) {
        return $RawFilter
    }

    $filterParts = [System.Collections.Generic.List[string]]::new()

    if (-not [string]::IsNullOrWhiteSpace($SupportedHost)) {
        $filterParts.Add("supportedHosts/any(h:h eq '$SupportedHost')")
    }

    if (-not [string]::IsNullOrWhiteSpace($ElementType)) {
        $filterParts.Add("elementTypes/any(h:h eq '$ElementType')")
    }

    if ($LastModifiedAfter.HasValue) {
        $filterParts.Add("lastModifiedDateTime gt $(ConvertTo-ODataDateTime -Value $LastModifiedAfter.Value)")
    }

    if ($LastModifiedBefore.HasValue) {
        $filterParts.Add("lastModifiedDateTime lt $(ConvertTo-ODataDateTime -Value $LastModifiedBefore.Value)")
    }

    return ($filterParts -join ' and ')
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

function Invoke-GraphWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Method,

        [Parameter(Mandatory)]
        [string] $Uri,

        [Parameter()]
        [int] $Retries = $MaxRetries
    )

    $attempt = 0
    do {
        try {
            return Invoke-MgGraphRequest -Method $Method -Uri $Uri
        }
        catch {
            $status = $_.Exception.Response?.StatusCode?.value__ ?? 0
            if (($status -in 429, 503) -and $attempt -lt $Retries) {
                $wait = Get-RetryDelaySeconds -Response $_.Exception.Response
                Write-Warning "[retry] HTTP $status - waiting ${wait}s before retry $($attempt + 1)/$Retries ..."
                Start-Sleep -Seconds $wait
                $attempt++
            }
            else {
                throw
            }
        }
    } while ($attempt -le $Retries)
}

if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
    Write-Error -Category ResourceUnavailable -ErrorAction Stop -Message "Required PowerShell module 'Microsoft.Graph.Authentication' is not installed. Install it before running this script: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser. Installation help: https://www.powershellgallery.com/packages/Microsoft.Graph.Authentication"
}
Import-Module -Name 'Microsoft.Graph.Authentication' -ErrorAction Stop

$odataFilter = New-CopilotPackagesFilter `
    -SupportedHost $SupportedHost `
    -ElementType $ElementType `
    -LastModifiedAfter $LastModifiedAfter `
    -LastModifiedBefore $LastModifiedBefore `
    -RawFilter $Filter

$listUrl = $PACKAGES_URL
if (-not [string]::IsNullOrWhiteSpace($odataFilter)) {
    $listUrl = "${PACKAGES_URL}?`$filter=$([System.Uri]::EscapeDataString($odataFilter))"
}

try {
    $connectParams = @{
        Scopes    = @($REQUIRED_SCOPE)
        NoWelcome = $true
    }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }

    Write-InteractiveMessage -Message "[auth] Connecting to Microsoft Graph (scope: $REQUIRED_SCOPE) ..." -ForegroundColor Cyan
    Connect-MgGraph @connectParams | Out-Null
    Write-InteractiveMessage -Message "[auth] Connected as: $((Get-MgContext).Account)" -ForegroundColor Green
    Write-InteractiveMessage -Message "[fetch] Enumerating Copilot packages ..." -ForegroundColor Cyan

    $packages = [System.Collections.Generic.List[pscustomobject]]::new()
    $pageUrl = $listUrl

    do {
        $page = Invoke-GraphWithRetry -Method GET -Uri $pageUrl
        foreach ($item in $page.value) {
            $packages.Add([pscustomobject]@{
                id                   = $item.id
                displayName          = $item.displayName ?? '(no name)'
                description          = $item.longDescription
                type                 = $item.type
                publisher            = $item.publisher ?? ''
                isBlocked            = [bool]($item.isBlocked)
                supportedHosts       = @($item.supportedHosts)
                elementTypes         = @($item.elementTypes)
                lastModifiedDateTime = $item.lastModifiedDateTime
            })
        }

        $pageUrl = $page.PSObject.Properties['@odata.nextLink']?.Value
        Write-Verbose "[fetch] Page complete - $($packages.Count) package(s) so far ..."
    } while ($pageUrl)

    Write-InteractiveMessage -Message "[fetch] Found $($packages.Count) package(s)." -ForegroundColor Cyan

    $allowListDocument = [pscustomobject]@{
        allowedAgents = @($packages | Sort-Object -Property displayName, id)
    }

    $json = $allowListDocument | ConvertTo-Json -Depth 8

    if ($OutputPath) {
        $outputDirectory = Split-Path -Path $OutputPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -Path $outputDirectory -PathType Container)) {
            Write-Error -Category ObjectNotFound -ErrorAction Stop -Message "Output directory not found: $outputDirectory"
        }

        Set-Content -Path $OutputPath -Value $json -Encoding utf8 -ErrorAction Stop
        Write-InteractiveMessage -Message "[done] Wrote allowlist-ready JSON to $OutputPath" -ForegroundColor Green
    }

    $json
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
