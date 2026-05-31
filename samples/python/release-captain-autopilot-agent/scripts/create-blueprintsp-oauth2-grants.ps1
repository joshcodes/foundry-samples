$ErrorActionPreference = "Stop"

# Force TLS 1.2 and enable connection reuse to avoid Windows TCP socket
# exhaustion (WinError 10048) when making several Graph calls back-to-back.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::DefaultConnectionLimit = 64

# Invoke-RestMethod with a small retry loop. Windows can run out of ephemeral
# TCP ports under bursty load and fail with "Only one usage of each socket
# address (protocol/network address/port) is normally permitted." Retrying
# with a short backoff gives the OS a chance to recycle TIME_WAIT sockets.
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

# Resolve a service principal's objectId from its appId by hitting Microsoft Graph
# directly. The `az ad sp show / list` commands intermittently fail on Windows
# with TCP socket exhaustion (WinError 10048), so we go through REST.
function Get-SpObjectId([string]$AppId, [string]$Token) {
    $uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$AppId'"
    $resp = Invoke-Graph -Method GET -Uri $uri -Headers @{ Authorization = "Bearer $Token" }
    if (-not $resp.value -or $resp.value.Count -eq 0) {
        return $null
    }
    return $resp.value[0].id
}

$graphToken = az account get-access-token --resource https://graph.microsoft.com/ --query accessToken -o tsv

$blueprintSP = Get-SpObjectId -AppId $env:AGENT_IDENTITY_BLUEPRINT_ID -Token $graphToken

if ([string]::IsNullOrEmpty($blueprintSP)) {
    throw "Failed to get service principal for blueprint ID $($env:AGENT_IDENTITY_BLUEPRINT_ID)"
}

Write-Host "Creating OAuth2 permission grants for blueprint service principal..."


$apxAppId = "5a807f24-c9de-44ee-a3a7-329e88a00ffc"

$apxSP = Get-SpObjectId -AppId $apxAppId -Token $graphToken
if ([string]::IsNullOrEmpty($apxSP)) {
    throw "Failed to get service principal for APEX app ID $apxAppId"
}

$prodMCPAppId = "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1"
$prodMCP_SP = Get-SpObjectId -AppId $prodMCPAppId -Token $graphToken

if ([string]::IsNullOrEmpty($prodMCP_SP)) {
    throw "Failed to get service principal for Prod MCP app ID $prodMCPAppId"
}

# 00000003-0000-0000-c000-000000000000 is graph appId
# (graphToken already acquired above for SP lookups)


$mcpOauthGrant = @"
{
  "clientId": "$blueprintSP",
  "consentType": "AllPrincipals",
  "principalId": null,
  "resourceId": "$prodMCP_SP",
  "scope": "McpServers.M365Admin.All McpServers.DASearch.All McpServers.WebSearch.All McpServers.Files.All AgentTools.MOSEvents.All McpServers.Admin365Graph.All McpServers.ERPAnalytics.All McpServers.DataverseCustom.All McpServers.Dataverse.All McpServers.D365Service.All McpServers.D365Sales.All McpServers.Management.All McpServersMetadata.Read.All McpServers.Developer.All McpServers.CopilotMCP.All McpServers.OneDriveSharepoint.All McpServers.OneDrive.All McpServers.SharePoint.All McpServers.Mail.All McpServers.Teams.All McpServers.Me.All McpServers.Calendar.All McpServers.SharepointLists.All McpServers.Knowledge.All McpServers.Excel.All McpServers.Word.All McpServers.PowerPoint.All"
}
"@
# Catch "Permission entry already exists" error and continue
try {
    $response = Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" `
        -Headers @{
            "Content-Type" = "application/json"
            "Accept"       = "application/json"
            "Authorization" = "Bearer $($graphToken)"
        } `
        -Body $mcpOauthGrant

    Write-Host ""
    Write-Host "MCP oauth grant response:"
    $response | ConvertTo-Json -Depth 5 | Write-Host

} catch {
    $err = $_.ErrorDetails.Message | ConvertFrom-Json
    if ($err.error.code -eq "Request_BadRequest" -and
        $err.error.message -like "*Permission entry already exists*") {

        Write-Host "Permission already exists  ignoring."
    }
    else {
        throw
    }
}


try {
    $apxOauthGrant = @"
    {
        "clientId": "$blueprintSP",
        "consentType": "AllPrincipals",
        "principalId": null,
        "resourceId": "$apxSP",
        "scope": "AgentData.ReadWrite"
    }
"@

    $response = Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" `
        -Headers @{
            "Content-Type" = "application/json"
            "Accept"       = "application/json"
            "Authorization" = "Bearer $($graphToken)"
        } `
        -Body $apxOauthGrant

    Write-Host ""
    Write-Host "APX oauth grant response:"
    $response | ConvertTo-Json -Depth 5 | Write-Host
}
catch {
    $err = $_.ErrorDetails.Message | ConvertFrom-Json
    if ($err.error.code -eq "Request_BadRequest" -and
        $err.error.message -like "*Permission entry already exists*") {

        Write-Host "Permission already exists  ignoring."
    }
    else {
        throw
    }
}
