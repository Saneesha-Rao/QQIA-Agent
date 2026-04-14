// QQIA Agent - Azure App Service (Free Tier)
// Just a web app — no Cosmos DB, no Bot Service, no Key Vault needed

@description('Base name for resources')
param appName string = 'qqia-agent'

@description('Azure region')
param location string = resourceGroup().location

@description('App Service SKU (F1=free, B1=always-on)')
param sku string = 'F1'

@description('Team access code for the web UI')
@secure()
param accessCode string = ''

// ---- Naming ----
var uniqueSuffix = uniqueString(resourceGroup().id)
var appServicePlanName = '${appName}-plan-${uniqueSuffix}'
var webAppName = '${appName}-${uniqueSuffix}'

// ---- App Service Plan ----
resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: appServicePlanName
  location: location
  sku: {
    name: sku
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
      alwaysOn: sku != 'F1' // Free tier doesn't support always-on
      webSocketsEnabled: true
      appSettings: [
        { name: 'PORT', value: '8080' }
        { name: 'NODE_ENV', value: 'production' }
        { name: 'ACCESS_CODE', value: accessCode }
        { name: 'SCM_DO_BUILD_DURING_DEPLOYMENT', value: 'false' }
      ]
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
    }
  }
}

// ---- Outputs ----
output webAppName string = webApp.name
output webAppUrl string = 'https://${webApp.properties.defaultHostName}'
