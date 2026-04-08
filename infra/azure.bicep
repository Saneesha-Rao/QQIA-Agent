// QQIA Agent - Azure Infrastructure for Teams Toolkit
// Provisions: App Service + Bot Service (simplified for TTK)

@description('Base name for resources')
param appName string = 'qqia-agent'

@description('Azure region')
param location string = resourceGroup().location

@description('Bot AAD App ID (from Teams Toolkit)')
param botAadAppClientId string

@secure()
@description('Bot AAD App Secret')
param botAadAppClientSecret string

@description('Tenant ID')
param tenantId string = subscription().tenantId

// ---- Naming ----
var uniqueSuffix = uniqueString(resourceGroup().id)
var appServicePlanName = '${appName}-plan-${uniqueSuffix}'
var webAppName = '${appName}-app-${uniqueSuffix}'
var botServiceName = '${appName}-bot-${uniqueSuffix}'

// ---- App Service Plan ----
resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: appServicePlanName
  location: location
  sku: {
    name: 'B1'
  }
  kind: 'linux'
  properties: {
    reserved: true
  }
}

// ---- Web App ----
resource webApp 'Microsoft.Web/sites@2023-01-01' = {
  name: webAppName
  location: location
  kind: 'app,linux'
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    siteConfig: {
      linuxFxVersion: 'NODE|20-lts'
      alwaysOn: true
      webSocketsEnabled: true
      appSettings: [
        { name: 'MICROSOFT_APP_ID', value: botAadAppClientId }
        { name: 'MICROSOFT_APP_PASSWORD', value: botAadAppClientSecret }
        { name: 'MICROSOFT_APP_TENANT_ID', value: tenantId }
        { name: 'PORT', value: '8080' }
        { name: 'NODE_ENV', value: 'production' }
        { name: 'SCM_DO_BUILD_DURING_DEPLOYMENT', value: 'true' }
      ]
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
    }
  }
}

// ---- Bot Service ----
resource botService 'Microsoft.BotService/botServices@2022-09-15' = {
  name: botServiceName
  location: 'global'
  kind: 'azurebot'
  sku: { name: 'F0' }
  properties: {
    displayName: 'QQIA Agent'
    description: 'FY27 Mint Rollover Tracker'
    endpoint: 'https://${webApp.properties.defaultHostName}/api/messages'
    msaAppId: botAadAppClientId
    msaAppTenantId: tenantId
    msaAppType: 'SingleTenant'
  }
}

// ---- Teams Channel ----
resource teamsChannel 'Microsoft.BotService/botServices/channels@2022-09-15' = {
  parent: botService
  name: 'MsTeamsChannel'
  location: 'global'
  properties: {
    channelName: 'MsTeamsChannel'
    properties: { isEnabled: true }
  }
}

// ---- Outputs ----
output botWebAppName string = webApp.name
output botWebAppEndpoint string = webApp.properties.defaultHostName
output botWebAppResourceId string = webApp.id
