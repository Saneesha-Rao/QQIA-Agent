// ============================================================
// QQIA Agent - Azure Infrastructure (Bicep)
// Provisions: App Service, Cosmos DB, Bot Service, Key Vault
// ============================================================

@description('Base name for all resources')
param appName string = 'qqia-agent'

@description('Azure region for resources')
param location string = resourceGroup().location

@description('App Service SKU')
param appServiceSku string = 'B1'

@description('Microsoft App ID for the bot (from Azure AD app registration)')
param microsoftAppId string

@secure()
@description('Microsoft App Password (client secret)')
param microsoftAppPassword string

@description('Azure AD Tenant ID')
param tenantId string = subscription().tenantId

// ---- Naming ----
var uniqueSuffix = uniqueString(resourceGroup().id)
var appServicePlanName = '${appName}-plan-${uniqueSuffix}'
var webAppName = '${appName}-app-${uniqueSuffix}'
var cosmosAccountName = '${appName}-db-${uniqueSuffix}'
var botServiceName = '${appName}-bot-${uniqueSuffix}'
var keyVaultName = '${appName}-kv-${uniqueSuffix}'

// ============================================================
// App Service Plan
// ============================================================
resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: appServicePlanName
  location: location
  sku: {
    name: appServiceSku
  }
  kind: 'linux'
  properties: {
    reserved: true
  }
}

// ============================================================
// Web App (Node.js runtime for the bot)
// ============================================================
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
        { name: 'MICROSOFT_APP_ID', value: microsoftAppId }
        { name: 'MICROSOFT_APP_PASSWORD', value: microsoftAppPassword }
        { name: 'MICROSOFT_APP_TENANT_ID', value: tenantId }
        { name: 'COSMOS_ENDPOINT', value: cosmosAccount.properties.documentEndpoint }
        { name: 'COSMOS_KEY', value: cosmosAccount.listKeys().primaryMasterKey }
        { name: 'COSMOS_DATABASE', value: 'qqia-agent' }
        { name: 'PORT', value: '8080' }
        { name: 'NODE_ENV', value: 'production' }
        { name: 'WEBSITE_NODE_DEFAULT_VERSION', value: '~20' }
        { name: 'SCM_DO_BUILD_DURING_DEPLOYMENT', value: 'true' }
      ]
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
    }
  }
}

// ============================================================
// Azure Cosmos DB (NoSQL API)
// ============================================================
resource cosmosAccount 'Microsoft.DocumentDB/databaseAccounts@2023-11-15' = {
  name: cosmosAccountName
  location: location
  kind: 'GlobalDocumentDB'
  properties: {
    databaseAccountOfferType: 'Standard'
    consistencyPolicy: {
      defaultConsistencyLevel: 'Session'
    }
    locations: [
      {
        locationName: location
        failoverPriority: 0
        isZoneRedundant: false
      }
    ]
    capabilities: [
      { name: 'EnableServerless' }
    ]
  }
}

resource cosmosDatabase 'Microsoft.DocumentDB/databaseAccounts/sqlDatabases@2023-11-15' = {
  parent: cosmosAccount
  name: 'qqia-agent'
  properties: {
    resource: {
      id: 'qqia-agent'
    }
  }
}

resource stepsContainer 'Microsoft.DocumentDB/databaseAccounts/sqlDatabases/containers@2023-11-15' = {
  parent: cosmosDatabase
  name: 'steps'
  properties: {
    resource: {
      id: 'steps'
      partitionKey: { paths: ['/workstream'], kind: 'Hash' }
      indexingPolicy: {
        automatic: true
        indexingMode: 'consistent'
      }
    }
  }
}

resource milestonesContainer 'Microsoft.DocumentDB/databaseAccounts/sqlDatabases/containers@2023-11-15' = {
  parent: cosmosDatabase
  name: 'milestones'
  properties: {
    resource: {
      id: 'milestones'
      partitionKey: { paths: ['/category'], kind: 'Hash' }
    }
  }
}

resource auditContainer 'Microsoft.DocumentDB/databaseAccounts/sqlDatabases/containers@2023-11-15' = {
  parent: cosmosDatabase
  name: 'audit'
  properties: {
    resource: {
      id: 'audit'
      partitionKey: { paths: ['/stepId'], kind: 'Hash' }
      defaultTtl: 7776000 // 90 days retention
    }
  }
}

resource usersContainer 'Microsoft.DocumentDB/databaseAccounts/sqlDatabases/containers@2023-11-15' = {
  parent: cosmosDatabase
  name: 'users'
  properties: {
    resource: {
      id: 'users'
      partitionKey: { paths: ['/role'], kind: 'Hash' }
    }
  }
}

// ============================================================
// Azure Bot Service
// ============================================================
resource botService 'Microsoft.BotService/botServices@2022-09-15' = {
  name: botServiceName
  location: 'global'
  kind: 'azurebot'
  sku: {
    name: 'S1'
  }
  properties: {
    displayName: 'QQIA Agent'
    description: 'FY27 Mint Rollover Status Tracker for Teams'
    endpoint: 'https://${webApp.properties.defaultHostName}/api/messages'
    msaAppId: microsoftAppId
    msaAppTenantId: tenantId
    msaAppType: 'SingleTenant'
  }
}

// Enable Teams channel
resource teamsChannel 'Microsoft.BotService/botServices/channels@2022-09-15' = {
  parent: botService
  name: 'MsTeamsChannel'
  location: 'global'
  properties: {
    channelName: 'MsTeamsChannel'
    properties: {
      isEnabled: true
    }
  }
}

// ============================================================
// Key Vault (for secure secret storage)
// ============================================================
resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: keyVaultName
  location: location
  properties: {
    sku: { family: 'A', name: 'standard' }
    tenantId: tenantId
    enableRbacAuthorization: true
    enableSoftDelete: true
    softDeleteRetentionInDays: 30
  }
}

// ============================================================
// Outputs
// ============================================================
output webAppName string = webApp.name
output webAppUrl string = 'https://${webApp.properties.defaultHostName}'
output botEndpoint string = 'https://${webApp.properties.defaultHostName}/api/messages'
output cosmosEndpoint string = cosmosAccount.properties.documentEndpoint
output botServiceName string = botService.name
output keyVaultName string = keyVault.name
