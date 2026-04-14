# QQIA Agent — Azure App Service Deployment Guide

## Overview

Deploy the QQIA Agent (FY27 Mint Rollover Tracker) to Azure App Service. This runs the web UI bot on a free or low-cost App Service, with no external dependencies (no Cosmos DB, no Bot Service registration needed).

**What you need:**
- Azure CLI installed (`az --version`)
- Contributor access to an Azure resource group
- ~5 minutes

---

## Option A: One-Command Deploy (Recommended)

### Prerequisites
1. Azure CLI installed and logged in:
   ```powershell
   az login
   ```
2. Know your **subscription ID** and a **resource group name** you have access to.

### Deploy
```powershell
.\deploy-simple.ps1 `
  -ResourceGroup "rg-qqia-agent" `
  -Subscription "YOUR-SUBSCRIPTION-ID" `
  -AccessCode "YourTeamAccessCode" `
  -Location "westus2"
```

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-ResourceGroup` | Yes | Resource group name (created if it doesn't exist) |
| `-Subscription` | Yes | Azure subscription ID |
| `-AccessCode` | No | Restricts access to your team (recommended) |
| `-Location` | No | Azure region (default: `westus2`) |

The script will:
1. Create the resource group (if needed)
2. Provision a Free (F1) App Service via Bicep
3. Build the TypeScript project
4. Package and deploy the code
5. Verify the health endpoint

---

## Option B: Manual Step-by-Step

### Step 1: Log in and set subscription

```powershell
az login
az account set --subscription "YOUR-SUBSCRIPTION-ID"
```

### Step 2: Create a resource group

```powershell
az group create --name rg-qqia-agent --location westus2
```

### Step 3: Deploy infrastructure (Bicep)

```powershell
az deployment group create `
  --resource-group rg-qqia-agent `
  --template-file infra/azure.bicep `
  --parameters accessCode="YourTeamAccessCode"
```

Note the outputs:
- `webAppName` — the App Service name
- `webAppUrl` — your public URL (e.g., `https://qqia-agent-abc123.azurewebsites.net`)

### Step 4: Build the project

```powershell
npm install
npm run build
```

### Step 5: Create deployment package

```powershell
# Create staging folder with production files only
$staging = "$env:TEMP\qqia-staging"
if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
New-Item -ItemType Directory -Path $staging | Out-Null

Copy-Item -Path dist -Destination $staging\dist -Recurse
Copy-Item -Path public -Destination $staging\public -Recurse
Copy-Item -Path data -Destination $staging\data -Recurse
Copy-Item -Path package.json -Destination $staging\package.json
Copy-Item -Path package-lock.json -Destination $staging\package-lock.json -ErrorAction SilentlyContinue

Push-Location $staging
npm install --omit=dev
Pop-Location

Compress-Archive -Path "$staging\*" -DestinationPath "$env:TEMP\qqia-deploy.zip" -Force
```

### Step 6: Deploy to App Service

```powershell
az webapp deploy `
  --resource-group rg-qqia-agent `
  --name YOUR-WEBAPP-NAME `
  --src-path "$env:TEMP\qqia-deploy.zip" `
  --type zip
```

### Step 7: Verify

Open in browser: `https://YOUR-WEBAPP-NAME.azurewebsites.net`

Or check the health endpoint:
```powershell
Invoke-RestMethod -Uri "https://YOUR-WEBAPP-NAME.azurewebsites.net/api/health"
```

---

## After Deployment

### 1. Pin as a Teams Tab
1. Go to your Teams channel ("QQIA Agent" in "FY27 QQIA Rollover")
2. Click **+** → **Website** tab
3. Enter the Azure URL: `https://YOUR-WEBAPP-NAME.azurewebsites.net`
4. Name it "QQIA Agent"

### 2. Update Office Script URL
If you have the Office Script ("Sync from QQIA Agent") in your Excel:
1. Open the script in Excel Online → **Automate** tab → **Sync from QQIA Agent**
2. Update the `botUrl` variable to your new Azure URL:
   ```
   var botUrl = "https://YOUR-WEBAPP-NAME.azurewebsites.net/api/steps/json";
   ```
   Or re-copy the script from: `https://YOUR-WEBAPP-NAME.azurewebsites.net/api/sync-script`
   (The access code is auto-injected.)

### 3. Share with Your Team
- Share the access code via Teams chat
- Teammates open the Teams tab — enter the code once and they're in for 7 days

---

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | `3978` (local) / `8080` (Azure) | Server port |
| `NODE_ENV` | `development` | Set to `production` on Azure |
| `ACCESS_CODE` | (empty) | Team access code for the web UI |

To change the access code after deployment:
```powershell
az webapp config appsettings set `
  --resource-group rg-qqia-agent `
  --name YOUR-WEBAPP-NAME `
  --settings ACCESS_CODE="NewAccessCode"
```

---

## Updating the Bot

After making code changes:

```powershell
# Rebuild and redeploy
npm run build

$staging = "$env:TEMP\qqia-staging"
if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
New-Item -ItemType Directory -Path $staging | Out-Null
Copy-Item -Path dist -Destination $staging\dist -Recurse
Copy-Item -Path public -Destination $staging\public -Recurse
Copy-Item -Path data -Destination $staging\data -Recurse
Copy-Item -Path package.json -Destination $staging\package.json

Push-Location $staging; npm install --omit=dev; Pop-Location
Compress-Archive -Path "$staging\*" -DestinationPath "$env:TEMP\qqia-deploy.zip" -Force

az webapp deploy `
  --resource-group rg-qqia-agent `
  --name YOUR-WEBAPP-NAME `
  --src-path "$env:TEMP\qqia-deploy.zip" `
  --type zip
```

---

## Cost

| SKU | Monthly Cost | Always-On | Notes |
|-----|-------------|-----------|-------|
| **F1 (Free)** | $0 | No — sleeps after ~20 min idle, wakes on request (~5s) | Good enough for team use |
| **B1 (Basic)** | ~$13/month | Yes | No cold starts, best experience |

The Bicep template defaults to F1. To upgrade:
```powershell
az appservice plan update `
  --resource-group rg-qqia-agent `
  --name YOUR-PLAN-NAME `
  --sku B1
```

---

## Troubleshooting

**App returns 500 / blank page:**
```powershell
az webapp log tail --resource-group rg-qqia-agent --name YOUR-WEBAPP-NAME
```

**Check app settings:**
```powershell
az webapp config appsettings list --resource-group rg-qqia-agent --name YOUR-WEBAPP-NAME -o table
```

**Restart the app:**
```powershell
az webapp restart --resource-group rg-qqia-agent --name YOUR-WEBAPP-NAME
```

**No access to subscription:**
Ask your manager for Contributor access to a resource group:
> "I need Contributor access to a resource group in [subscription name] to deploy a free-tier App Service for the FY27 rollover tracker."
