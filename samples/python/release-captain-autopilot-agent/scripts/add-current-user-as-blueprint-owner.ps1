#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

# Force TLS 1.2 and enable connection reuse to avoid Windows TCP socket
# exhaustion (WinError 10048) when making several Graph calls back-to-back.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::DefaultConnectionLimit = 64

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

Write-Host "Adding current az login user as owner on the blueprint application..."

$blueprintAppId = $env:AGENT_IDENTITY_BLUEPRINT_ID
if ([string]::IsNullOrEmpty($blueprintAppId)) {
    throw "AGENT_IDENTITY_BLUEPRINT_ID environment variable is not set."
}

$graphToken = az account get-access-token --resource https://graph.microsoft.com/ --query accessToken -o tsv
if ([string]::IsNullOrEmpty($graphToken)) {
    throw "Failed to acquire a Microsoft Graph access token."
}

# Get the current signed-in user's object ID via Microsoft Graph (avoids
# `az ad signed-in-user show` which intermittently fails on Windows with TCP
# socket exhaustion).
$meResp = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/me" -Headers @{ Authorization = "Bearer $graphToken" }
$currentUserId = $meResp.id
if ([string]::IsNullOrEmpty($currentUserId)) {
    throw "Failed to get the current signed-in user's object ID. Make sure you are logged in via 'az login'."
}

Write-Host "Current user object ID: $currentUserId"

# Resolve the blueprint application's object ID from its App ID via Microsoft Graph.
# The `az ad app show / list` commands intermittently fail on Windows with TCP
# socket exhaustion (WinError 10048), so we go through REST.
$appLookupUri = "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$blueprintAppId'"
$appResp = Invoke-Graph -Method GET -Uri $appLookupUri -Headers @{ Authorization = "Bearer $graphToken" }
$blueprintAppObjectId = if ($appResp.value -and $appResp.value.Count -gt 0) { $appResp.value[0].id } else { $null }
if ([string]::IsNullOrEmpty($blueprintAppObjectId)) {
    throw "Failed to get application object ID for blueprint app ID $blueprintAppId"
}

$ownerBody = @{
    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$currentUserId"
} | ConvertTo-Json

try {
    $response = Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$blueprintAppObjectId/owners/`$ref" `
        -Headers @{
            "Content-Type"  = "application/json"
            "Accept"        = "application/json"
            "Authorization" = "Bearer $graphToken"
        } `
        -Body $ownerBody

    Write-Host "Current user added as owner of blueprint application $blueprintAppId."
    if ($response) {
        $response | ConvertTo-Json -Depth 5 | Write-Host
    }
}
catch {
    $err = $null
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
        try { $err = $_.ErrorDetails.Message | ConvertFrom-Json } catch { $err = $null }
    }

    if ($err -and $err.error -and $err.error.message -like "*One or more added object references already exist*") {
        Write-Host "Current user is already an owner of the blueprint application; ignoring."
    }
    else {
        throw
    }
}
