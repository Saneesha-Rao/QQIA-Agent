<#
.SYNOPSIS
  One-command deployment of QQIA Agent to Azure.

.DESCRIPTION
  This script:
  1. Creates an Azure AD App Registration (bot identity)
  2. Provisions Azure infrastructure via Bicep (App Service, Cosmos DB, Bot Service)
  3. Builds and deploys the bot code
  4. Generates the Teams app package for sideloading

.PARAMETER ResourceGroup
  Azure resource group name (will be created if it doesn't exist)

.PARAMETER Location
  Azure region (default: westus2)

.PARAMETER Subscription
  Azure subscription ID (optional, uses default)

.EXAMPLE
  .\deploy.ps1 -ResourceGroup "rg-qqia-agent" -Location "westus2"
#>

param(
  [Parameter(Mandatory=$true)]
  [string]$ResourceGroup,

  [string]$Location = "westus2",
  [string]$Subscription = "",
  [string]$AppName = "qqia-agent"
)

$ErrorActionPreference = "Stop"

Write-Host "`n🚀 QQIA Agent - Azure Deployment" -ForegroundColor Cyan
Write-Host "================================`n" -ForegroundColor Cyan

# ---- Pre-flight checks ----
Write-Host "🔍 Checking prerequisites..." -ForegroundColor Yellow

$azVersion = az version 2>&1 | ConvertFrom-Json
if (-not $azVersion) {
  Write-Host "❌ Azure CLI not found. Install from: https://aka.ms/installazurecli" -ForegroundColor Red
  exit 1
}
Write-Host "   ✅ Azure CLI: $($azVersion.'azure-cli')"

$account = az account show 2>&1 | ConvertFrom-Json
if (-not $account) {
  Write-Host "   ⚠️ Not logged in. Running 'az login'..." -ForegroundColor Yellow
  az login
  $account = az account show | ConvertFrom-Json
}
Write-Host "   ✅ Logged in as: $($account.user.name)"
Write-Host "   ✅ Subscription: $($account.name)"

if ($Subscription) {
  az account set --subscription $Subscription
  Write-Host "   ✅ Switched to subscription: $Subscription"
}

# ---- Step 1: Create Resource Group ----
Write-Host "`n📦 Step 1: Resource Group" -ForegroundColor Yellow
$rgExists = az group exists --name $ResourceGroup 2>&1
if ($rgExists -eq "true") {
  Write-Host "   ✅ Resource group '$ResourceGroup' already exists"
} else {
  Write-Host "   Creating resource group '$ResourceGroup' in '$Location'..."
  az group create --name $ResourceGroup --location $Location --output none
  Write-Host "   ✅ Resource group created"
}

# ---- Step 2: Create Azure AD App Registration ----
Write-Host "`n🔐 Step 2: Azure AD App Registration" -ForegroundColor Yellow

$existingApp = az ad app list --display-name "$AppName-bot" --query "[0]" 2>&1 | ConvertFrom-Json
if ($existingApp -and $existingApp.appId) {
  $appId = $existingApp.appId
  Write-Host "   ✅ App Registration already exists: $appId"
} else {
  Write-Host "   Creating app registration '$AppName-bot'..."
  $app = az ad app create `
    --display-name "$AppName-bot" `
    --sign-in-audience "AzureADMyOrg" `
    2>&1 | ConvertFrom-Json
  $appId = $app.appId
  Write-Host "   ✅ App Registration created: $appId"
}

# Create client secret
Write-Host "   Creating client secret..."
$secret = az ad app credential reset `
  --id $appId `
  --display-name "qqia-deploy-$(Get-Date -Format 'yyyyMMdd')" `
  --years 2 `
  2>&1 | ConvertFrom-Json
$appPassword = $secret.password
$tenantId = $secret.tenant
Write-Host "   ✅ Client secret created (save this — it won't be shown again)"

# ---- Step 3: Deploy Azure Infrastructure via Bicep ----
Write-Host "`n🏗️  Step 3: Deploying Azure Infrastructure" -ForegroundColor Yellow
Write-Host "   This provisions: App Service, Cosmos DB, Bot Service, Key Vault..."

$deployResult = az deployment group create `
  --resource-group $ResourceGroup `
  --template-file "infra/main.bicep" `
  --parameters appName=$AppName `
               microsoftAppId=$appId `
               microsoftAppPassword=$appPassword `
               location=$Location `
  --query "properties.outputs" `
  2>&1 | ConvertFrom-Json

$webAppName = $deployResult.webAppName.value
$webAppUrl = $deployResult.webAppUrl.value
$botEndpoint = $deployResult.botEndpoint.value
$cosmosEndpoint = $deployResult.cosmosEndpoint.value

Write-Host "   ✅ Infrastructure deployed:"
Write-Host "      Web App:  $webAppUrl"
Write-Host "      Bot:      $botEndpoint"
Write-Host "      Cosmos:   $cosmosEndpoint"

# ---- Step 4: Build and Deploy Code ----
Write-Host "`n📤 Step 4: Building and Deploying Code" -ForegroundColor Yellow

Write-Host "   Building TypeScript..."
npm run build 2>&1 | Out-Null
Write-Host "   ✅ Build complete"

# Create deployment package
Write-Host "   Creating deployment package..."
$zipPath = "$env:TEMP\qqia-agent-deploy.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath }

Compress-Archive -Path @(
  "dist",
  "node_modules",
  "package.json",
  "package-lock.json"
) -DestinationPath $zipPath -Force
Write-Host "   ✅ Package created: $zipPath"

# Deploy to App Service
Write-Host "   Deploying to Azure App Service..."
az webapp deploy `
  --resource-group $ResourceGroup `
  --name $webAppName `
  --src-path $zipPath `
  --type zip `
  --output none `
  2>&1
Write-Host "   ✅ Code deployed to $webAppUrl"

# ---- Step 5: Verify deployment ----
Write-Host "`n✅ Step 5: Verifying Deployment" -ForegroundColor Yellow
Start-Sleep -Seconds 10

try {
  $health = Invoke-RestMethod -Uri "$webAppUrl/api/health" -TimeoutSec 30
  Write-Host "   ✅ Health check passed: $($health.status)"
  Write-Host "      Tracked steps: $($health.trackedSteps)"
} catch {
  Write-Host "   ⚠️ Health check pending (app may still be starting). Check: $webAppUrl/api/health" -ForegroundColor Yellow
}

# ---- Step 6: Generate Teams App Package ----
Write-Host "`n📱 Step 6: Generating Teams App Package" -ForegroundColor Yellow

$manifestContent = Get-Content "appPackage/manifest.json" -Raw
$manifestContent = $manifestContent.Replace("{{MICROSOFT_APP_ID}}", $appId)
$teamsPackageDir = "$env:TEMP\qqia-teams-package"
if (Test-Path $teamsPackageDir) { Remove-Item $teamsPackageDir -Recurse }
New-Item -ItemType Directory -Path $teamsPackageDir | Out-Null

$manifestContent | Set-Content "$teamsPackageDir/manifest.json"

# Create placeholder icons
$iconScript = @"
Add-Type -AssemblyName System.Drawing
`$bmp = New-Object System.Drawing.Bitmap(192, 192)
`$g = [System.Drawing.Graphics]::FromImage(`$bmp)
`$g.FillRectangle([System.Drawing.Brushes]::DodgerBlue, 0, 0, 192, 192)
`$font = New-Object System.Drawing.Font('Arial', 48, [System.Drawing.FontStyle]::Bold)
`$g.DrawString('QQ', `$font, [System.Drawing.Brushes]::White, 40, 60)
`$bmp.Save('$teamsPackageDir/color.png', [System.Drawing.Imaging.ImageFormat]::Png)
`$g.Dispose(); `$bmp.Dispose()
`$bmp2 = New-Object System.Drawing.Bitmap(32, 32)
`$g2 = [System.Drawing.Graphics]::FromImage(`$bmp2)
`$g2.FillRectangle([System.Drawing.Brushes]::DodgerBlue, 0, 0, 32, 32)
`$bmp2.Save('$teamsPackageDir/outline.png', [System.Drawing.Imaging.ImageFormat]::Png)
`$g2.Dispose(); `$bmp2.Dispose()
"@
try {
  Invoke-Expression $iconScript
  Write-Host "   ✅ Icons generated"
} catch {
  # Fallback: create minimal valid PNGs
  Write-Host "   ⚠️ Icon generation skipped (add color.png + outline.png manually)" -ForegroundColor Yellow
}

$teamsZipPath = "qqia-agent-teams-app.zip"
Compress-Archive -Path "$teamsPackageDir/*" -DestinationPath $teamsZipPath -Force
Write-Host "   ✅ Teams app package: $teamsZipPath"

# ---- Summary ----
Write-Host "`n" -NoNewLine
Write-Host "============================================" -ForegroundColor Green
Write-Host " ✅ QQIA Agent Deployment Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Web App URL:    $webAppUrl" -ForegroundColor Cyan
Write-Host "  Bot Endpoint:   $botEndpoint" -ForegroundColor Cyan
Write-Host "  App ID:         $appId" -ForegroundColor Cyan
Write-Host "  Tenant ID:      $tenantId" -ForegroundColor Cyan
Write-Host "  Cosmos DB:      $cosmosEndpoint" -ForegroundColor Cyan
Write-Host ""
Write-Host "  📱 Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Upload '$teamsZipPath' to Teams Admin Center or sideload it"
Write-Host "     Teams → Apps → Manage your apps → Upload a custom app"
Write-Host "  2. Add the bot to your QQIA team/channel"
Write-Host "  3. Send 'help' to the bot to get started"
Write-Host "  4. Set up Power Automate flows per docs/power-automate-integration.md"
Write-Host ""
Write-Host "  🔑 Save these credentials securely:" -ForegroundColor Red
Write-Host "  App Password:   $($appPassword.Substring(0,4))****"
Write-Host ""

# Clean up
Remove-Item $zipPath -ErrorAction SilentlyContinue
Remove-Item $teamsPackageDir -Recurse -ErrorAction SilentlyContinue
