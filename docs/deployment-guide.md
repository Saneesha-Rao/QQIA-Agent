# QQIA Agent - Azure Deployment Guide

## Prerequisites
- Azure CLI installed (`az --version`)
- Azure subscription with Contributor access
- Node.js 20+ installed
- PowerShell 7+

## Quick Deploy (One Command)

```powershell
cd C:\Users\salingal\qqia-agent
.\deploy.ps1 -ResourceGroup "rg-qqia-agent" -Location "westus2"
```

This single command will:
1. ✅ Create an Azure AD App Registration for the bot identity
2. ✅ Provision all Azure resources (App Service, Cosmos DB, Bot Service, Key Vault)
3. ✅ Build and deploy the bot code
4. ✅ Generate a Teams app package (.zip) for sideloading

## What Gets Created in Azure

| Resource | Type | Purpose |
|----------|------|---------|
| `qqia-agent-app-*` | App Service (B1) | Hosts the bot Node.js server |
| `qqia-agent-plan-*` | App Service Plan | Linux Node.js 20 hosting plan |
| `qqia-agent-db-*` | Cosmos DB (Serverless) | Steps, milestones, audit, users |
| `qqia-agent-bot-*` | Azure Bot Service | Teams channel integration |
| `qqia-agent-kv-*` | Key Vault | Secure secret storage |

**Estimated monthly cost**: ~$15–30/month (B1 App Service + Serverless Cosmos DB)

## After Deployment

### 1. Install the Teams App
- Open Microsoft Teams
- Go to **Apps** → **Manage your apps** → **Upload a custom app**
- Upload the generated `qqia-agent-teams-app.zip`
- Add the bot to your QQIA team/channel

### 2. Test the Bot
Send these messages to the bot in Teams:
- `help` — see all commands
- `dashboard` — overall rollover progress
- `my tasks` — your assigned steps
- `status 1.A` — check a specific step
- `summary` — leadership executive summary

### 3. Configure Graph API for Excel Sync (Optional)
To enable automatic Excel sync via SharePoint:

1. In Azure Portal → Azure AD → App Registrations → your bot app
2. Add API permissions: `Files.ReadWrite.All`, `Sites.ReadWrite.All`
3. Grant admin consent
4. In App Service → Configuration, add:
   - `GRAPH_CLIENT_ID` = your app ID
   - `GRAPH_CLIENT_SECRET` = your client secret
   - `GRAPH_TENANT_ID` = your tenant ID
   - `EXCEL_DRIVE_ID` = OneDrive drive ID
   - `EXCEL_ITEM_ID` = Excel file item ID

### 4. Set Up Power Automate (Optional)
See `docs/power-automate-integration.md` for webhook flows that auto-update
step statuses from ADO, pipelines, and environment checks.

## CI/CD (GitHub Actions)
Push to `main` branch auto-deploys. Set these GitHub repository secrets:

| Secret | Value |
|--------|-------|
| `AZURE_CLIENT_ID` | Bot App Registration client ID |
| `AZURE_TENANT_ID` | Azure AD tenant ID |
| `AZURE_SUBSCRIPTION_ID` | Azure subscription ID |
| `AZURE_WEBAPP_NAME` | App Service name (from deploy output) |

## Troubleshooting

| Issue | Fix |
|-------|-----|
| Bot doesn't respond | Check App Service logs: `az webapp log tail --name <app> --resource-group <rg>` |
| Excel sync fails | Verify Graph API permissions and file path in app settings |
| Teams app upload fails | Ensure manifest.json has correct App ID, try Teams Admin Center |
| Cosmos DB errors | Check connection string in App Service Configuration |
