#!/usr/bin/env pwsh
# -----------------------------------------------------------------------------
# Grants the OAuth2 scopes and Azure RBAC role assignments a hired digital
# worker instance (Agent Identity / "AI" SP) needs to actually function.
#
# When you hire an agent in Microsoft Teams, A365 creates a brand-new
# ServiceIdentity service principal for that specific hired instance (the
# "AI"). The post-provision scripts only consent the *blueprint* SP — the AI
# SP has NO grants out of the box, so its very first message hits 400/403.
#
# This script grants the same Agent Tools (MCP) and Messaging Bot API
# (AgentData.ReadWrite) OAuth2 scopes used by the blueprint SP, plus the
# Cognitive Services / Foundry RBAC the AI needs to call the Foundry
# project's agent endpoint and Azure OpenAI Responses API.
#
# Usage:
#   ./grant-hired-instance-access.ps1 -AiClientId <appId of the hired instance>
#
# To find the AI client id, look in Entra → Enterprise Applications for a SP
# whose display name matches the hired digital worker (e.g. "Signal Digital
# Worker"). The AI is the one with servicePrincipalType=ServiceIdentity.
# -----------------------------------------------------------------------------

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "AppId / objectId of the hired-instance AI service principal")]
    [string] $AiClientId,

    [string] $SubscriptionId,
    [string] $ResourceGroup,
    [string] $AccountName,
    [string] $ProjectName
)

$ErrorActionPreference = "Stop"

# Force TLS 1.2 and enable connection reuse to avoid Windows TCP socket
# exhaustion (WinError 10048) when making several Graph calls back-to-back.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::DefaultConnectionLimit = 64

# Invoke-RestMethod with a small retry loop. Windows can run out of ephemeral
# TCP ports under bursty load and fail with WinError 10048. Retrying with a
# short backoff gives the OS a chance to recycle TIME_WAIT sockets.
function Invoke-Graph {
    param(
        [string]$Method = "GET",
        [string]$Uri,
        [hashtable]$Headers,
        $Body
    )
    $maxRetries = 5
    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            if ($Body) {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -Body $Body
            } else {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers
            }
        } catch {
            $msg = $_.Exception.Message
            $isSocket = $msg -like "*Only one usage of each socket address*" -or `
                        $msg -like "*Unable to read data from the transport*" -or `
                        $msg -like "*Max retries exceeded*"
            if ($isSocket -and $i -lt $maxRetries) {
                $delay = [Math]::Min(30, 2 * $i)
                Write-Host "  (Graph call hit socket exhaustion, retrying in ${delay}s — attempt $i/$maxRetries)"
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

# Pull defaults from azd env if not provided
$azd = @{}
azd env get-values 2>$null | ForEach-Object {
    if ($_ -match '^([A-Z_]+)="?(.*?)"?$') { $azd[$matches[1]] = $matches[2] }
}
if (-not $SubscriptionId)  { $SubscriptionId  = if ($azd["SUBSCRIPTION_ID"]) { $azd["SUBSCRIPTION_ID"] } else { $azd["AZURE_SUBSCRIPTION_ID"] } }
if (-not $ResourceGroup)   { $ResourceGroup   = if ($azd["RESOURCE_GROUP"])  { $azd["RESOURCE_GROUP"] }  else { $azd["AZURE_RESOURCE_GROUP"] } }
if (-not $AccountName)     { $AccountName     = $azd["ACCOUNT_NAME"] }
if (-not $ProjectName)     { $ProjectName     = $azd["PROJECT_NAME"] }

foreach ($n in "SubscriptionId","ResourceGroup","AccountName","ProjectName") {
    if (-not (Get-Variable -Name $n -ValueOnly)) {
        throw "$n could not be resolved. Pass -$n or run azd env get-values from a configured environment."
    }
}

Write-Host "Granting access to hired AI:"
Write-Host "  AI client id  : $AiClientId"
Write-Host "  Subscription  : $SubscriptionId"
Write-Host "  Account/RG    : $AccountName / $ResourceGroup"
Write-Host "  Project       : $ProjectName"
Write-Host ""

# Resolve AI SP object id (it might be passed as appId or objectId).
# Use Microsoft Graph directly — `az ad sp show` intermittently fails on
# Windows with TCP socket exhaustion (WinError 10048).
$graphTokenForLookup = az account get-access-token --resource https://graph.microsoft.com/ --query accessToken -o tsv
$aiObjectId = $null
# First try treating the input as an appId
try {
    $byAppId = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$AiClientId')?`$select=id" -Headers @{ Authorization = "Bearer $graphTokenForLookup" }
    if ($byAppId.id) { $aiObjectId = $byAppId.id }
} catch {
    # not found by appId — maybe it's already an objectId
    try {
        $byObjId = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$AiClientId`?`$select=id" -Headers @{ Authorization = "Bearer $graphTokenForLookup" }
        if ($byObjId.id) { $aiObjectId = $byObjId.id }
    } catch { }
}
if (-not $aiObjectId) {
    throw "Could not resolve service principal with id '$AiClientId'. Make sure the agent has been hired in Teams and the SP exists in your tenant."
}
Write-Host "Resolved AI service principal object id: $aiObjectId"
Write-Host ""

# -----------------------------------------------------------------------------
# 1) OAuth2 grants (delegated, AllPrincipals)
# -----------------------------------------------------------------------------
$graphToken = az account get-access-token --resource https://graph.microsoft.com/ --query accessToken -o tsv
$gh = @{ Authorization = "Bearer $graphToken"; "Content-Type" = "application/json" }

# Resource SPs we grant to
$mcpAppId = "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1" # Agent Tools (MCP)
$apxAppId = "5a807f24-c9de-44ee-a3a7-329e88a00ffc" # Messaging Bot API (APX)

$mcpSpId = (Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$mcpAppId')?`$select=id" -Headers @{Authorization="Bearer $graphToken"}).id
$apxSpId = (Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$apxAppId')?`$select=id" -Headers @{Authorization="Bearer $graphToken"}).id

$mcpScopes = "McpServers.M365Admin.All McpServers.DASearch.All McpServers.WebSearch.All McpServers.Files.All AgentTools.MOSEvents.All McpServers.Admin365Graph.All McpServers.ERPAnalytics.All McpServers.DataverseCustom.All McpServers.Dataverse.All McpServers.D365Service.All McpServers.D365Sales.All McpServers.Management.All McpServersMetadata.Read.All McpServers.Developer.All McpServers.CopilotMCP.All McpServers.OneDriveSharepoint.All McpServers.OneDrive.All McpServers.SharePoint.All McpServers.Mail.All McpServers.Teams.All McpServers.Me.All McpServers.Calendar.All McpServers.SharepointLists.All McpServers.Knowledge.All McpServers.Excel.All McpServers.Word.All McpServers.PowerPoint.All"

function Set-OAuth2Grant {
    param([string]$Client, [string]$Resource, [string]$Scope, [string]$Label)

    $existing = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$Client' and resourceId eq '$Resource'" -Headers $gh
    if ($existing.value -and $existing.value.Count -gt 0) {
        $g = $existing.value[0]
        $current = (($g.scope -split "\s+") | Where-Object { $_ }) | Sort-Object -Unique
        $needed  = (($Scope -split "\s+") | Where-Object { $_ }) | Sort-Object -Unique
        $missing = $needed | Where-Object { $_ -notin $current }
        if (-not $missing) {
            Write-Host "  [$Label] already has all required scopes — skipping."
        } else {
            $merged = ($current + $missing | Sort-Object -Unique) -join " "
            $body = @{ scope = $merged } | ConvertTo-Json
            Invoke-Graph -Method PATCH -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$($g.id)" -Headers $gh -Body $body | Out-Null
            Write-Host "  [$Label] patched: added $($missing -join ', ')"
        }
    } else {
        $body = @{ clientId = $Client; consentType = "AllPrincipals"; resourceId = $Resource; scope = $Scope } | ConvertTo-Json
        Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" -Headers $gh -Body $body | Out-Null
        Write-Host "  [$Label] created new grant."
    }
}

Write-Host "=== Step 1: OAuth2 grants (delegated, AllPrincipals) ==="
Set-OAuth2Grant -Client $aiObjectId -Resource $mcpSpId -Scope $mcpScopes -Label "Agent Tools (MCP)"
Set-OAuth2Grant -Client $aiObjectId -Resource $apxSpId -Scope "AgentData.ReadWrite" -Label "Messaging Bot API"
Write-Host ""

# -----------------------------------------------------------------------------
# 2) Azure RBAC on the Foundry account + project
# -----------------------------------------------------------------------------
$accountScope = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.CognitiveServices/accounts/$AccountName"
$projectScope = "$accountScope/projects/$ProjectName"

$roles = @(
    @{ name = "Foundry User";                   id = "53ca6127-db72-4b80-b1b0-d745d6d5456d"; scopes = @($accountScope, $projectScope) },
    @{ name = "Cognitive Services User";        id = "a97b65f3-24c7-4388-baec-2e87135dc908"; scopes = @($accountScope) },
    @{ name = "Cognitive Services OpenAI User"; id = "5e0bd9bd-7b93-4f28-af87-19fc36ad61bd"; scopes = @($accountScope) }
)

Write-Host "=== Step 2: Azure RBAC role assignments ==="
foreach ($role in $roles) {
    foreach ($scope in $role.scopes) {
        $out = az role assignment create --assignee-object-id $aiObjectId --assignee-principal-type ServicePrincipal --role $role.id --scope $scope 2>&1 | Out-String
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  [$($role.name)] assigned on $($scope.Substring([Math]::Max(0, $scope.Length - 80)))"
        } elseif ($out -match "RoleAssignmentExists") {
            Write-Host "  [$($role.name)] already exists on $($scope.Substring([Math]::Max(0, $scope.Length - 80)))"
        } else {
            Write-Host "  [$($role.name)] ERROR: $out"
        }
    }
}

Write-Host ""
Write-Host "Done. RBAC propagation can take 1-2 minutes — retry your chat in Teams after that."
