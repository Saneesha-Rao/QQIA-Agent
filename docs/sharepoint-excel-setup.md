# SharePoint Excel Sync Setup Guide

This guide walks you through connecting the QQIA Agent to your Excel file on SharePoint/OneDrive for Business via Microsoft Graph API. Once configured, the bot reads/writes Excel data **directly through SharePoint** — no file locks, no OneDrive sync issues.

## Prerequisites

- Access to [Azure Portal](https://portal.azure.com) (Entra ID / Azure AD)
- Admin consent (or ask your IT admin) for the Graph API permissions
- Your Excel file hosted on OneDrive for Business or a SharePoint site

---

## Step 1: Create an App Registration

1. Go to **Azure Portal → Microsoft Entra ID → App registrations → New registration**
2. Name: `QQIA-Agent-Graph`
3. Supported account types: **Single tenant** (this org only)
4. Redirect URI: leave blank (not needed for app-only auth)
5. Click **Register**

**Save these values:**
- **Application (client) ID** → `GRAPH_CLIENT_ID`
- **Directory (tenant) ID** → `GRAPH_TENANT_ID`

## Step 2: Create a Client Secret

1. In the app registration, go to **Certificates & secrets → Client secrets → New client secret**
2. Description: `QQIA Agent Excel Access`
3. Expiry: 12 months (or your org policy)
4. Click **Add**

**Save the secret value immediately** → `GRAPH_CLIENT_SECRET`
(You won't be able to see it again!)

## Step 3: Add API Permissions

1. Go to **API permissions → Add a permission → Microsoft Graph → Application permissions**
2. Add these permissions:
   - `Files.ReadWrite.All` — Read/write all files (needed to access the Excel file)
   - `Sites.ReadWrite.All` — Read/write SharePoint sites (alternative, broader scope)
3. Click **Grant admin consent for [your org]** (requires admin role)

> **Note:** If you can't grant admin consent yourself, send your IT admin the App Registration ID and ask them to grant consent for `Files.ReadWrite.All`.

## Step 4: Find Your Excel File IDs

You need the **Drive ID** and **Item ID** of your Excel file. Here's how to find them:

### Option A: Using Graph Explorer (easiest)

1. Go to [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in with your Microsoft account
3. Run this query to find your OneDrive files:

```
GET https://graph.microsoft.com/v1.0/me/drive/root:/Seller Incentives/QQIA/FY27_Mint_RolloverTimeline.xlsx
```

4. From the response, copy:
   - `parentReference.driveId` → `EXCEL_DRIVE_ID`
   - `id` → `EXCEL_ITEM_ID`

### Option B: Using the helper script

Run the included PowerShell script:

```powershell
cd qqia-agent
node scripts/find-excel-ids.js
```

This will prompt for your credentials and output the Drive ID and Item ID.

### Option C: SharePoint site URL

If the file is on a SharePoint site (not personal OneDrive), you can find it via:

```
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:/{path-to-file}
```

To find your site ID:
```
GET https://graph.microsoft.com/v1.0/sites?search=Seller Incentives
```

## Step 5: Configure the .env File

Edit the `.env` file in the project root:

```env
# Microsoft Graph API
GRAPH_CLIENT_ID=<your-app-client-id>
GRAPH_CLIENT_SECRET=<your-client-secret>
GRAPH_TENANT_ID=<your-tenant-id>

# Excel file location on OneDrive/SharePoint
EXCEL_DRIVE_ID=<drive-id-from-step-4>
EXCEL_ITEM_ID=<item-id-from-step-4>
EXCEL_FILE_PATH=/Seller Incentives/QQIA/FY27_Mint_RolloverTimeline.xlsx
```

## Step 6: Test the Connection

Restart the bot:

```bash
npm run build && npm start
```

You should see in the console:
```
✅ Graph API initialized
📡 Excel sync: Graph API mode enabled (SharePoint direct read/write)
📡 Graph import: 181 steps, 30 milestones
```

Then try updating a step via chat:
```
update 1.C completed
```

The Excel file on SharePoint should update **immediately** — no file locks!

---

## How It Works

```
┌──────────────┐     Graph API      ┌──────────────────┐
│  QQIA Bot     │ ◄──────────────► │  SharePoint Excel  │
│  (Teams)      │  read/write cells │  FY27_Rollover     │
└──────────────┘                    └──────────────────┘
       │                                    ▲
       │ in-memory cache                    │
       ▼                                    │
┌──────────────┐                    Users edit Excel
│  Data Store   │                    directly in browser
│  (In-Memory   │                    or desktop app
│   or Cosmos)  │
└──────────────┘
```

- **Bot updates a status** → Graph API writes to the specific cell in SharePoint Excel
- **Someone edits Excel manually** → Next sync cycle (every 15 min) reads changes via Graph API
- **No file locks** — Graph API uses the Excel REST API which works concurrently
- **Local fallback** — If Graph API is down, bot uses a local file copy

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `Graph client not initialized` | Check GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_TENANT_ID in .env |
| `403 Forbidden` | Admin consent not granted. Ask IT admin to approve `Files.ReadWrite.All` |
| `404 Not Found` | EXCEL_DRIVE_ID or EXCEL_ITEM_ID is wrong. Re-run Step 4 |
| `The workbook is currently locked` | Another Graph session has the workbook. Wait 5 min and retry |
| Falls back to local file | Graph API failed — check network/credentials |
