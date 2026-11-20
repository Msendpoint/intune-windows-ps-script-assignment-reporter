### FILE: README.md
# Intune Windows PowerShell Script Assignment Reporter

## What this repo does
This project inventories Intune **Windows PowerShell Script** assignments using **Microsoft Graph** and produces a CSV report that is suitable for audits and change reviews.

It resolves:
- Built-in include targets: **All Devices** / **All Users** (service release 2611 behavior)
- Included Entra groups and excluded Entra groups (group exclusions)

## Output
A CSV exported to the provided `ExportPath` containing one row per script assignment, with readable group names where possible.

## Components
- `src/Reporter.ps1`: Production-ready PowerShell script logic.
- `index.php`: M365 SaaS Engine-style endpoint skeleton that calls the PowerShell logic.

## Prerequisites
- Admin workstation
- PowerShell 7.x (Windows PowerShell 5.1 supported)
- Microsoft Graph PowerShell SDK v2 modules:
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.DeviceManagement`
  - `Microsoft.Graph.Groups`

## Required Graph Permissions
- `DeviceManagementConfiguration.Read.All`
- `Group.Read.All`

## Run (PowerShell)
```powershell
.
# Example
.\src\Reporter.ps1 -ExportPath .\Intune-WindowsPowerShellScriptAssignments.csv
```

## Run (SaaS)
Use the repository entrypoint `index.php` per your platform conventions.


### FILE: index.php
<?php
/**
 * M365 SaaS Engine endpoint skeleton
 *
 * This file demonstrates how to call Microsoft Graph (via the provided $ms->graphCall helper)
 * and then render UI cards (via render_premium_card).
 *
 * IMPORTANT:
 * - Actual execution of the inventory logic is implemented in PowerShell (src/Reporter.ps1)
 * - This endpoint can be adapted to trigger job execution in your SaaS environment.
 */

// You said: $ms is already initialized.
// You said: $accessToken is in $_SESSION['ms_access_token'].

// Defensive checks
session_start();
$accessToken = $_SESSION['ms_access_token'] ?? null;

if (!$accessToken) {
    http_response_code(401);
    echo json_encode(['error' => 'Missing ms_access_token in session']);
    exit;
}

$title = 'Intune Assignment Reporter';

// Example: quick health check call (count scripts). The full report generation is done in PowerShell.
$endpoint = 'https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?$select=id,displayName';

try {
    $resp = $ms->graphCall($endpoint, $accessToken, 'GET', null);
    $count = isset($resp['value']) ? count($resp['value']) : 0;

    echo render_premium_card(
        'Discovered Intune Windows PowerShell scripts',
        (string)$count,
        null,
        'up',
        '📊',
        null
    );

    // In a real SaaS environment, you would enqueue / run src/Reporter.ps1 and return download links.
    // For now, return guidance.
    echo json_encode([
        'ok' => true,
        'scripts_discovered' => $count,
        'next' => 'Run src/Reporter.ps1 to export the full CSV report with include/exclude resolution.'
    ]);
} catch (Throwable $e) {
    http_response_code(500);
    echo json_encode(['error' => $e->getMessage()]);
}


### FILE: src/Reporter.ps1
#!/usr/bin/env pwsh
<#
.SYNOPSIS
Inventories Intune Windows PowerShell script assignments, resolving built-in include targets (All Devices/All Users) and group include/exclusion scopes.

.DESCRIPTION
This admin workstation script queries Microsoft Graph for Intune Windows PowerShell scripts (including their assignments), then builds an audit-ready CSV report.

The script understands Intune assignment targets including:
- Built-in include targets: All Devices / All Users (newer behavior aligned to Intune targeting model)
- Included groups (groupAssignmentTarget)
- Excluded groups (exclusionGroupAssignmentTarget)

It outputs a clean report with readable group display names when possible. It also implements throttling retry logic for Graph requests.

.EXAMPLE
.
# Export report to CSV
.\src\Reporter.ps1 -ExportPath .\Intune-WindowsPowerShellScriptAssignments.csv

.NOTES
Author:      Souhaiel Morhag
Company:     MSEndpoint.com
Blog:        https://msendpoint.com
Academy:     https://app.msendpoint.com/academy
LinkedIn:    https://linkedin.com/in/souhaiel-morhag
GitHub:      https://github.com/Msendpoint
License:     MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ExportPath = '.\Intune-WindowsPowerShellScriptAssignments.csv',

    [Parameter(Mandatory = $false)]
    [switch]$IncludeRawAssignments,

    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO','WARN','ERROR','OK')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    switch ($Level) {
        'INFO'  { Write-Host "[$timestamp] [INFO ] $Message" -ForegroundColor Cyan }
        'WARN'  { Write-Host "[$timestamp] [WARN ] $Message" -ForegroundColor Yellow }
        'ERROR' { Write-Host "[$timestamp] [ERROR] $Message" -ForegroundColor Red }
        'OK'     { Write-Host "[$timestamp] [ OK  ] $Message" -ForegroundColor Green }
    }
}

function Test-MgModules {
    [CmdletBinding()]
    param()

    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.DeviceManagement',
        'Microsoft.Graph.Groups'
    )

    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "Required module '$module' is not installed. Install-Module $module -Scope CurrentUser"
        }
    }
}

function Connect-GraphIfNeeded {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$Scopes = @(
            'DeviceManagementConfiguration.Read.All',
            'Group.Read.All'
        )
    )

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Log -Message 'Connecting to Microsoft Graph...' -Level INFO
        Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
    }
    else {
        Write-Log -Message "Reusing existing Microsoft Graph context for tenant: $($context.TenantId)" -Level INFO
    }
}

function Invoke-GraphGetWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetryCount = 5
    )

    $attempt = 0
    do {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        }
        catch {
            $attempt++
            $response = $_.Exception.Response
            $statusCode = $null
            $retryAfter = 5

            if ($response) {
                try { $statusCode = [int]$response.StatusCode } catch { $statusCode = $null }

                try {
                    if ($response.Headers -and $response.Headers.Contains('Retry-After')) {
                        $retryAfter = [int]$response.Headers.GetValues('Retry-After')[0]
                    }
                }
                catch {
                    $retryAfter = 5
                }
            }

            if ((($statusCode -eq 429) -or ($statusCode -eq 503)) -and $attempt -lt $MaxRetryCount) {
                Write-Log -Message "Graph throttling/service busy on '$Uri'. Retry $attempt of $MaxRetryCount in $retryAfter second(s)." -Level WARN
                Start-Sleep -Seconds $retryAfter
            }
            else {
                throw
            }
        }
    } while ($attempt -lt $MaxRetryCount)

    throw "Graph request failed after $MaxRetryCount attempts for URI: $Uri"
}

# Cache group display names to reduce Graph calls
$script:GroupNameCache = @{}

function Get-GroupDisplayNameSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )

    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        return ''
    }

    if ($script:GroupNameCache.ContainsKey($GroupId)) {
        return $script:GroupNameCache[$GroupId]
    }

    try {
        $group = Get-MgGroup -GroupId $GroupId -Property Id, DisplayName -ErrorAction Stop
        $script:GroupNameCache[$GroupId] = $group.DisplayName
        return $group.DisplayName
    }
    catch {
        Write-Log -Message "Failed to resolve group display name for GroupId '$GroupId'. Using raw ID." -Level WARN
        $script:GroupNameCache[$GroupId] = $GroupId
        return $GroupId
    }
}

function Resolve-AssignmentTarget {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$AssignmentTarget
    )

    # The Graph SDK uses an odata.type marker under property '@odata.type'
    $odataType = $AssignmentTarget.'@odata.type'

    switch ($odataType) {
        '#microsoft.graph.allDevicesAssignmentTarget' {
            return [pscustomobject]@{
                TargetType   = 'Include'
                ScopeType    = 'BuiltIn'
                ScopeName    = 'All Devices'
                ScopeId      = 'BuiltIn-AllDevices'
                RawODataType = $odataType
            }
        }
        '#microsoft.graph.allLicensedUsersAssignmentTarget' {
            return [pscustomobject]@{
                TargetType   = 'Include'
                ScopeType    = 'BuiltIn'
                ScopeName    = 'All Users'
                ScopeId      = 'BuiltIn-AllUsers'
                RawODataType = $odataType
            }
        }
        '#microsoft.graph.groupAssignmentTarget' {
            $groupId = $AssignmentTarget.groupId
            return [pscustomobject]@{
                TargetType   = 'Include'
                ScopeType    = 'Group'
                ScopeName    = (Get-GroupDisplayNameSafe -GroupId $groupId)
                ScopeId      = $groupId
                RawODataType = $odataType
            }
        }
        '#microsoft.graph.exclusionGroupAssignmentTarget' {
            $groupId = $AssignmentTarget.groupId
            return [pscustomobject]@{
                TargetType   = 'Exclude'
                ScopeType    = 'Group'
                ScopeName    = (Get-GroupDisplayNameSafe -GroupId $groupId)
                ScopeId      = $groupId
                RawODataType = $odataType
            }
        }