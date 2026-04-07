# QQIA Agent — Azure Portal Manual Provisioning Guide

Follow these steps in order. The entire process takes ~20 minutes.

---

## Step 1: Create Azure AD App Registration (Bot Identity)

1. Go to **[Azure Portal → Azure AD → App Registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)**
2. Click **+ New registration**
   - **Name**: `qqia-agent-bot`
   - **Supported account types**: Accounts in this organizational directory only (Single tenant)
   - **Redirect URI**: Leave blank
3. Click **Register**
4. **Copy these values** (you'll need them later):
   - **Application (client) ID**: `________________________`
   - **Directory (tenant) ID**: `________________________`
5. Go to **Certificates & secrets** → **+ New client secret**
   - **Description**: `qqia-deploy`
   - **Expires**: 24 months
   - Click **Add**
   - **⚠️ Copy the secret Value now** (it won't be shown again): `________________________`

---

## Step 2: Create Resource Group

1. Go to **[Azure Portal → Resource Groups](https://portal.azure.com/#view/HubsExtension/BrowseResourceGroups)**
2. Click **+ Create**
   - **Subscription**: Your subscription
   - **Resource group**: `rg-qqia-agent`
   - **Region**: West US 2 (or your preferred region)
3. Click **Review + create** → **Create**

---

## Step 3: Create Azure Cosmos DB (Serverless)

1. Go to **[Azure Portal → Create Cosmos DB](https://portal.azure.com/#create/Microsoft.DocumentDB)**
2. Select **Azure Cosmos DB for NoSQL** → **Create**
   - **Subscription**: Your subscription
   - **Resource group**: `rg-qqia-agent`
   - **Account name**: `qqia-agent-db` (must be globally unique, add a suffix if needed)
   - **Location**: West US 2
   - **Capacity mode**: **Serverless** ← Important! (cheapest option, ~$0-1/month)
3. Click **Review + create** → **Create**
4. Once created, go to the resource → **Keys**
   - **Copy URI**: `________________________`
   - **Copy PRIMARY KEY**: `________________________`

### Create the database and containers:
5. In Cosmos DB → **Data Explorer** → **New Database**
   - **Database id**: `qqia-agent`
6. Click the database → **New Container** (repeat 4 times):

   | Container ID | Partition Key |
   |-------------|---------------|
   | `steps` | `/workstream` |
   | `milestones` | `/category` |
   | `audit` | `/stepId` |
   | `users` | `/role` |

---

## Step 4: Create App Service (Bot Hosting)

1. Go to **[Azure Portal → Create Web App](https://portal.azure.com/#create/Microsoft.WebSite)**
   - **Subscription**: Your subscription
   - **Resource group**: `rg-qqia-agent`
   - **Name**: `qqia-agent-app` (must be globally unique, add suffix if needed)
   - **Publish**: Code
   - **Runtime stack**: **Node 20 LTS**
   - **Operating System**: **Linux**
   - **Region**: West US 2
   - **Pricing plan**: **Basic B1** (~$13/month)
2. Click **Review + create** → **Create**
3. Once created, go to the resource → **Configuration** → **Application settings**
4. Add these settings (click **+ New application setting** for each):

   | Name | Value |
   |------|-------|
   | `MICROSOFT_APP_ID` | *(App Registration Client ID from Step 1)* |
   | `MICROSOFT_APP_PASSWORD` | *(Client Secret from Step 1)* |
   | `MICROSOFT_APP_TENANT_ID` | *(Tenant ID from Step 1)* |
   | `COSMOS_ENDPOINT` | *(Cosmos DB URI from Step 3)* |
   | `COSMOS_KEY` | *(Cosmos DB Primary Key from Step 3)* |
   | `COSMOS_DATABASE` | `qqia-agent` |
   | `PORT` | `8080` |
   | `NODE_ENV` | `production` |
   | `SCM_DO_BUILD_DURING_DEPLOYMENT` | `true` |

5. Click **Save** at the top

### Deploy the code:
6. Go to **Deployment Center** → **Settings**
   - **Source**: **Local Git** (or GitHub if you push to a repo)
   - Click **Save**
7. Copy the **Git Clone Uri** shown
8. From your local machine terminal:
   ```powershell
   cd C:\Users\salingal\qqia-agent
   npm run build
   git init
   git add .
   git commit -m "Initial QQIA Agent deployment"
   git remote add azure <GIT_CLONE_URI_FROM_PORTAL>
   git push azure main
   ```

---

## Step 5: Create Azure Bot Service

1. Go to **[Azure Portal → Create Azure Bot](https://portal.azure.com/#create/Microsoft.AzureBot)**
   - **Bot handle**: `qqia-agent-bot`
   - **Subscription**: Your subscription
   - **Resource group**: `rg-qqia-agent`
   - **Pricing tier**: **Standard**
   - **Type of App**: **Single Tenant**
   - **Creation type**: **Use existing app registration**
   - **App ID**: *(Client ID from Step 1)*
   - **App tenant ID**: *(Tenant ID from Step 1)*
2. Click **Review + create** → **Create**
3. Once created, go to the resource → **Configuration**
   - **Messaging endpoint**: `https://<YOUR_APP_SERVICE_NAME>.azurewebsites.net/api/messages`
   - Click **Apply**
4. Go to **Channels** → Click **Microsoft Teams** → **Apply**
   - This enables the bot to work in Teams

---

## Step 6: Create & Upload Teams App Package

1. Open `C:\Users\salingal\qqia-agent\appPackage\manifest.json`
2. Replace `{{MICROSOFT_APP_ID}}` with your **App Registration Client ID** from Step 1
3. Add two icon files to the `appPackage` folder:
   - `color.png` (192×192 px, any icon)
   - `outline.png` (32×32 px, transparent outline)
4. Zip all 3 files together:
   ```powershell
   Compress-Archive -Path appPackage\* -DestinationPath qqia-agent-teams.zip
   ```
5. In **Microsoft Teams**:
   - Go to **Apps** → **Manage your apps** → **Upload an app**
   - Select **Upload a custom app**
   - Choose `qqia-agent-teams.zip`
6. Click **Add** to install it for yourself, or **Add to a team** to install in your QQIA channel

---

## Step 7: Verify Everything Works

1. In Teams, message the **QQIA Agent** bot:
   - Send: `help` → Should show command list
   - Send: `dashboard` → Should show rollover progress card
   - Send: `status 1.A` → Should show step details
2. Check the health endpoint in a browser:
   - `https://<YOUR_APP_NAME>.azurewebsites.net/api/health`
   - Should return `{"status":"healthy","trackedSteps":182,...}`

---

## Step 8: Enable Excel Sync via Graph API (Optional)

1. In **Azure AD → App Registrations → qqia-agent-bot**
2. Go to **API permissions** → **+ Add a permission**
   - Microsoft Graph → Application permissions
   - Add: `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `User.Read.All`
3. Click **Grant admin consent**
4. Find your Excel file IDs:
   - Go to SharePoint/OneDrive → Right-click the Excel file → **Details**
   - Or use Graph Explorer: `https://developer.microsoft.com/graph/graph-explorer`
   - Query: `GET /me/drive/root:/Seller Incentives/QQIA/FY27_Mint_RolloverTimeline.xlsx`
   - Copy the `id` (EXCEL_ITEM_ID) and `parentReference.driveId` (EXCEL_DRIVE_ID)
5. Add these App Service settings:
   | Name | Value |
   |------|-------|
   | `GRAPH_CLIENT_ID` | *(same as MICROSOFT_APP_ID)* |
   | `GRAPH_CLIENT_SECRET` | *(same as MICROSOFT_APP_PASSWORD)* |
   | `GRAPH_TENANT_ID` | *(same as MICROSOFT_APP_TENANT_ID)* |
   | `EXCEL_DRIVE_ID` | *(from step 4)* |
   | `EXCEL_ITEM_ID` | *(from step 4)* |

Once configured, the bot will auto-sync with Excel every 15 minutes.

---

## Cost Summary

| Resource | Monthly Cost |
|----------|-------------|
| App Service (B1) | ~$13 |
| Cosmos DB (Serverless) | ~$0-2 |
| Bot Service (Standard) | Free |
| Key Vault | ~$0.03 |
| **Total** | **~$13-15/month** |

---

## Sharing with POCs

Once deployed, any team member can:
1. Find **QQIA Agent** in Teams Apps
2. Start a 1:1 chat or add to a channel
3. Type `help` to see all commands
4. Type `my tasks` to see their assigned steps

No additional setup needed per user — the bot uses Azure AD for authentication automatically.
