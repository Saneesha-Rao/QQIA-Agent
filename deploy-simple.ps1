<#
.SYNOPSIS
  Deploy QQIA Agent to Azure App Service (Free Tier).

.DESCRIPTION
  Deploys the web UI bot to a free Azure App Service.
  No Cosmos DB, Bot Service, or Key Vault needed.

.PARAMETER ResourceGroup
  Azure resource group name (created if it doesn't exist)

.PARAMETER Subscription
  Azure subscription ID

.PARAMETER AccessCode
  Team access code to restrict access (optional)

.PARAMETER Location
  Azure region (default: westus2)

.EXAMPLE
  .\deploy-simple.ps1 -ResourceGroup "rg-qqia" -Subscription "e0809ea2-..." -AccessCode "MintFY27"
#>

param(
  [Parameter(Mandatory=$true)]
  [string]$ResourceGroup,

  [Parameter(Mandatory=$true)]
  [string]$Subscription,

  [string]$AccessCode = "",
  [string]$Location = "westus2"
)

$ErrorActionPreference = "Stop"

Write-Host "`n=== QQIA Agent - Azure Deployment ===" -ForegroundColor Cyan

# ---- Pre-flight ----
$account = az account show 2>&1 | ConvertFrom-Json
if (-not $account) {
  Write-Host "Not logged in. Running 'az login'..." -ForegroundColor Yellow
  az login
}

az account set --subscription $Subscription
Write-Host "Subscription: $Subscription" -ForegroundColor Green

# ---- Step 1: Resource Group ----
Write-Host "`n[1/4] Resource Group..." -ForegroundColor Yellow
$rgExists = az group exists --name $ResourceGroup 2>&1
if ($rgExists -eq "true") {
  Write-Host "  Already exists: $ResourceGroup"
} else {
  az group create --name $ResourceGroup --location $Location --output none
  Write-Host "  Created: $ResourceGroup in $Location"
}

# ---- Step 2: Deploy Infrastructure (Bicep) ----
Write-Host "`n[2/4] Provisioning App Service..." -ForegroundColor Yellow

$deployResult = az deployment group create `
  --resource-group $ResourceGroup `
  --template-file "infra/azure.bicep" `
  --parameters accessCode="$AccessCode" `
  --query "properties.outputs" `
  -o json 2>&1 | ConvertFrom-Json

$webAppName = $deployResult.webAppName.value
$webAppUrl = $deployResult.webAppUrl.value

Write-Host "  Web App: $webAppName"
Write-Host "  URL:     $webAppUrl"

# ---- Step 3: Build ----
Write-Host "`n[3/4] Building..." -ForegroundColor Yellow
npm run build 2>&1 | Out-Null
Write-Host "  Build complete"

# ---- Step 4: Deploy Code ----
Write-Host "`n[4/4] Deploying code..." -ForegroundColor Yellow

# Create zip with built code + dependencies + static files + data
$zipPath = "$env:TEMP\qqia-deploy.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath }

# Create a staging folder with only what's needed
$stagingDir = "$env:TEMP\qqia-staging"
if (Test-Path $stagingDir) { Remove-Item $stagingDir -Recurse -Force }
New-Item -ItemType Directory -Path $stagingDir | Out-Null

# Copy required files
Copy-Item -Path "dist" -Destination "$stagingDir\dist" -Recurse
Copy-Item -Path "public" -Destination "$stagingDir\public" -Recurse
Copy-Item -Path "data" -Destination "$stagingDir\data" -Recurse
Copy-Item -Path "package.json" -Destination "$stagingDir\package.json"
Copy-Item -Path "package-lock.json" -Destination "$stagingDir\package-lock.json" -ErrorAction SilentlyContinue

# Install production dependencies in staging
Push-Location $stagingDir
npm install --omit=dev 2>&1 | Out-Null
Pop-Location

# Zip it
Compress-Archive -Path "$stagingDir\*" -DestinationPath $zipPath -Force
Write-Host "  Package created"

# Deploy
az webapp deploy `
  --resource-group $ResourceGroup `
  --name $webAppName `
  --src-path $zipPath `
  --type zip `
  --output none 2>&1
Write-Host "  Deployed!"

# ---- Verify ----
Write-Host "`nWaiting for startup..." -ForegroundColor Yellow
Start-Sleep -Seconds 15

try {
  $health = Invoke-RestMethod -Uri "$webAppUrl/api/health" -TimeoutSec 30
  Write-Host "  Health: OK - $($health.trackedSteps) steps loaded" -ForegroundColor Green
} catch {
  Write-Host "  App may still be starting. Check: $webAppUrl/api/health" -ForegroundColor Yellow
}

# ---- Summary ----
Write-Host "`n=== Deployment Complete ===" -ForegroundColor Green
Write-Host "  URL: $webAppUrl" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Open $webAppUrl in your browser"
Write-Host "  2. Pin it as a Teams tab in your channel"
Write-Host "  3. Update your Office Script URL to: $webAppUrl/api/steps/json"
if ($AccessCode) {
  Write-Host "  4. Share the access code with your team" -ForegroundColor Yellow
}
Write-Host ""

# Cleanup
Remove-Item $zipPath -ErrorAction SilentlyContinue
Remove-Item $stagingDir -Recurse -Force -ErrorAction SilentlyContinue
