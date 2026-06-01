#!/usr/bin/env pwsh
# -----------------------------------------------------------------------------
# Reverses everything ./grant-hired-instance-access.ps1 did for a single hired
# digital worker instance (Agent Identity / "AI" SP):
#
#   1. Deletes the OAuth2 delegated grants (AllPrincipals) on the Agent Tools
#      (MCP) and Messaging Bot API (APX) resources.
#   2. Deletes the three Azure RBAC role assignments on the Foundry account /
#      project (Foundry User, Cognitive Services User, Cognitive Services
#      OpenAI User).
#
# It does NOT delete the hired-instance service principal itself, nor does it
# touch the agent blueprint, the digital worker publish record, or any Azure
# resources provisioned by `azd provision`. Run this once per hired instance
# you want to retire, then `azd down` to tear down the rest of the stack.
#
# Usage:
#   ./cleanup-hired-instance-access.ps1 -AiClientId <hired-instance-appId>
#   ./cleanup-hired-instance-access.ps1 -AiClientId <hired-instance-appId> -WhatIf
#   ./cleanup-hired-instance-access.ps1 -AiClientId <hired-instance-appId> -Force
# -----------------------------------------------------------------------------

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "AppId / objectId of the hired-instance AI service principal")]
    [string] $AiClientId,

    [string] $SubscriptionId,
    [string] $ResourceGroup,
    [string] $AccountName,
    [string] $ProjectName,

    [switch] $WhatIf,
    [switch] $Force
)

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
                Write-Host "  (Graph call hit socket exhaustion, retrying in ${delay}s - attempt $i/$maxRetries)"
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

Write-Host "Cleaning up access for hired AI:"
Write-Host "  AI client id  : $AiClientId"
Write-Host "  Subscription  : $SubscriptionId"
Write-Host "  Account/RG    : $AccountName / $ResourceGroup"
Write-Host "  Project       : $ProjectName"
Write-Host "  Mode          : $(if ($WhatIf) { 'WhatIf (dry run)' } else { 'Apply' })"
Write-Host ""

if (-not $WhatIf -and -not $Force) {
    $confirm = Read-Host "This will revoke OAuth2 grants and RBAC role assignments for the SP above. Continue? (y/N)"
    if ($confirm -ne 'y' -and $confirm -ne 'Y') {
        Write-Host "Aborted."
        exit 0
    }
}

# Resolve AI SP object id (accept either appId or objectId).
$graphTokenForLookup = az account get-access-token --resource https://graph.microsoft.com/ --query accessToken -o tsv
$aiObjectId = $null
try {
    $byAppId = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$AiClientId')?`$select=id" -Headers @{ Authorization = "Bearer $graphTokenForLookup" }
    if ($byAppId.id) { $aiObjectId = $byAppId.id }
} catch {
    try {
        $byObjId = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$AiClientId`?`$select=id" -Headers @{ Authorization = "Bearer $graphTokenForLookup" }
        if ($byObjId.id) { $aiObjectId = $byObjId.id }
    } catch { }
}
if (-not $aiObjectId) {
    throw "Could not resolve service principal with id '$AiClientId'."
}
Write-Host "Resolved AI service principal object id: $aiObjectId"
Write-Host ""

# -----------------------------------------------------------------------------
# 1) Delete OAuth2 grants
# -----------------------------------------------------------------------------
$graphToken = az account get-access-token --resource https://graph.microsoft.com/ --query accessToken -o tsv
$gh = @{ Authorization = "Bearer $graphToken"; "Content-Type" = "application/json" }

$mcpAppId = "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1" # Agent Tools (MCP)
$apxAppId = "5a807f24-c9de-44ee-a3a7-329e88a00ffc" # Messaging Bot API (APX)

$mcpSpId = (Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$mcpAppId')?`$select=id" -Headers @{Authorization="Bearer $graphToken"}).id
$apxSpId = (Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$apxAppId')?`$select=id" -Headers @{Authorization="Bearer $graphToken"}).id

function Remove-OAuth2Grant {
    param([string]$Client, [string]$Resource, [string]$Label)

    $existing = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$Client' and resourceId eq '$Resource'" -Headers $gh
    if (-not $existing.value -or $existing.value.Count -eq 0) {
        Write-Host "  [$Label] no grant found - skipping."
        return
    }
    foreach ($g in $existing.value) {
        if ($WhatIf) {
            Write-Host "  [$Label] would delete grant $($g.id) (scope: $($g.scope))"
        } else {
            Invoke-Graph -Method DELETE -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$($g.id)" -Headers $gh | Out-Null
            Write-Host "  [$Label] deleted grant $($g.id)"
        }
    }
}

Write-Host "=== Step 1: Delete OAuth2 grants ==="
Remove-OAuth2Grant -Client $aiObjectId -Resource $mcpSpId -Label "Agent Tools (MCP)"
Remove-OAuth2Grant -Client $aiObjectId -Resource $apxSpId -Label "Messaging Bot API"
Write-Host ""

# -----------------------------------------------------------------------------
# 2) Delete Azure RBAC role assignments
# -----------------------------------------------------------------------------
$accountScope = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.CognitiveServices/accounts/$AccountName"
$projectScope = "$accountScope/projects/$ProjectName"

$roles = @(
    @{ name = "Foundry User";                   id = "53ca6127-db72-4b80-b1b0-d745d6d5456d"; scopes = @($accountScope, $projectScope) },
    @{ name = "Cognitive Services User";        id = "a97b65f3-24c7-4388-baec-2e87135dc908"; scopes = @($accountScope) },
    @{ name = "Cognitive Services OpenAI User"; id = "5e0bd9bd-7b93-4f28-af87-19fc36ad61bd"; scopes = @($accountScope) }
)

Write-Host "=== Step 2: Delete Azure RBAC role assignments ==="
foreach ($role in $roles) {
    foreach ($scope in $role.scopes) {
        $scopeTail = $scope.Substring([Math]::Max(0, $scope.Length - 80))
        if ($WhatIf) {
            Write-Host "  [$($role.name)] would delete on $scopeTail"
            continue
        }
        $out = az role assignment delete --assignee-object-id $aiObjectId --assignee-principal-type ServicePrincipal --role $role.id --scope $scope 2>&1 | Out-String
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  [$($role.name)] deleted on $scopeTail"
        } else {
            Write-Host "  [$($role.name)] WARN: $out"
        }
    }
}

Write-Host ""
Write-Host "Done. The hired-instance SP is no longer authorized against Agent Tools, the Messaging Bot API, or the Foundry account/project."
Write-Host "To tear down the rest of the deployment, run: azd down"
