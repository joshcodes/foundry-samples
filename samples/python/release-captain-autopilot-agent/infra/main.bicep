targetScope = 'resourceGroup'

// =================================================================================================
// Main parameters
// =================================================================================================

@minLength(1)
@maxLength(64)
@description('Name of the application. Used to ensure resource names are unique.')
param environmentName string

@minLength(1)
@description('Primary location for all resources')
param location string

// =================================================================================================
// Existing Foundry resource parameters
//
// This sample reuses an existing Foundry account, project, and Container Registry
// (provisioned by the foundry-ai-teammate sample). The bicep here ONLY creates
// the resources needed to publish a new agent on top of that shared infrastructure:
//   1. (SKIPPED) project module — the account / project / ACR already exist
//   2. Deployment-script UMI (with role grants on the RG)
//   3. Managed Agent Identity Blueprint (MAIB) via dataplane PowerShell script
//   4. Bot Service + Teams channel for the new agent
// =================================================================================================

@description('Name of the EXISTING Cognitive Services (Foundry) account to reuse.')
param accountName string = 'signalteammateacct'

@description('Name of the EXISTING Cognitive Services (Foundry) project to reuse.')
param projectName string = 'signalteammateproj'

@description('Name of the EXISTING Container Registry to reuse.')
param containerRegistryName string = 'signalteammateacr'

param agentName string = 'release-captain'

param maibName string = '${agentName}-maib'

// =================================================================================================
// Bot Service module parameters
// =================================================================================================

@description('Name of the Bot Service')
param botName string = '${agentName}-bot'

@description('Display name of the bot')
param botDisplayName string = '${agentName} Bot'

@description('SKU of the Bot Service')
param botServiceSku string = 'F0'

@description('Model name (deployed on the existing Foundry account)')
param modelName string = 'gpt-5.5'

// =================================================================================================
// Common parameters
// =================================================================================================

@description('Tags to apply to all resources')
param tags object = {}

// =================================================================================================
// Existing resource references (NOT redeployed)
// =================================================================================================

resource existingAccount 'Microsoft.CognitiveServices/accounts@2025-09-01' existing = {
  name: accountName
}

resource existingProject 'Microsoft.CognitiveServices/accounts/projects@2026-03-01' existing = {
  parent: existingAccount
  name: projectName
}

resource existingAcr 'Microsoft.ContainerRegistry/registries@2023-07-01' existing = {
  name: containerRegistryName
}

// =================================================================================================
// Module deployments
// =================================================================================================

// 1. SKIPPED — existing project, account, ACR are reused.

// 2. Create deployment script UMI and grant roles on RG.
module deploymentScriptUmi 'modules/deployment-script-umi.bicep' = {
  name: 'deployment-script-umi'
}

// 3. Create managed agent identity blueprint using a deployment script as that is a dataplane operation.
module deploymentScriptAgent 'modules/maib-creation-script.bicep' = {
  name: 'maib-creation-script'
  params: {
    uamiResourceId: deploymentScriptUmi.outputs.uamiResourceId
    azureAIProjectEndpoint: existingProject.properties.endpoints['AI Foundry API']
    maibName: maibName
  }
  dependsOn: [
    deploymentScriptUmi
  ]
}


// 4. Deploy the bot service module
module botService 'modules/botservice.bicep' = {
  name: 'botservice-deployment'
  params: {
    botName: botName
    displayName: botDisplayName
    msaAppId: deploymentScriptAgent.outputs.blueprintClientId
    endpoint: 'https://${accountName}.services.ai.azure.com/api/projects/${projectName}/agents/${agentName}/endpoint/protocols/activityProtocol?api-version=2025-05-15-preview'
    botServiceSku: botServiceSku
  }
  dependsOn: [
    deploymentScriptAgent
  ]
}

// =================================================================================================
// Outputs - These become environment variables consumed by the post-provision scripts
// =================================================================================================

@description('ACR login server endpoint')
output AZURE_CONTAINER_REGISTRY_ENDPOINT string = existingAcr.properties.loginServer

output AZURE_AI_PROJECT_ENDPOINT string = existingProject.properties.endpoints['AI Foundry API']

@description('Agent identity blueprint ID')
output AGENT_IDENTITY_BLUEPRINT_ID string = deploymentScriptAgent.outputs.blueprintClientId

output SUBSCRIPTION_ID string = subscription().subscriptionId

output RESOURCE_GROUP string = resourceGroup().name

output LOCATION string = location

output ACCOUNT_NAME string = accountName

output PROJECT_NAME string = projectName

output AGENT_NAME string = agentName

output TENANT_ID string = tenant().tenantId

output PROJECT_PRINCIPAL_ID string = existingProject.identity.principalId

output MAIB_NAME string = maibName

output MODEL_NAME string = modelName
