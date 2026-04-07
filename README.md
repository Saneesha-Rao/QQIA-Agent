# QQIA Agent 🤖

**FY27 Mint Rollover Status Tracker for Microsoft Teams**

An intelligent Teams bot that manages the FY27 Mint Rollover process — tracking 182 steps across 15 workstreams, 29 DRIs, with automated notifications and bi-directional Excel sync.

## Features

| Feature | Description |
|---------|-------------|
| 📊 **Dashboard** | Real-time rollover progress with Adaptive Cards |
| 💬 **Natural Language** | Update and query steps via Teams chat |
| 🔔 **Proactive Alerts** | Deadline, overdue, blocker, and weekly digest notifications |
| 🔗 **Dependency Engine** | DAG-based critical path and blocker visualization |
| 📋 **Excel Sync** | Bi-directional sync with SharePoint Excel every 15 min |
| ⚡ **Automation** | Power Automate webhooks for ADO, pipeline, and environment monitoring |
| 👥 **RBAC** | DRI, PM, Leadership, and Admin roles via Azure AD |

## Bot Commands

```
dashboard          → Overall rollover progress card
my tasks           → Your assigned steps
status 1.A         → View step 1.A details
update 1.A done    → Mark step 1.A complete
blockers           → View all blocked steps
overdue            → View all overdue steps
workstream System Rollover → View workstream
tasks for Jim R    → View a person's tasks
critical path      → Current critical path
summary            → Executive summary for leadership
sync               → Trigger Excel sync
help               → Full command list
```

## Architecture

```
Teams Bot (Bot Framework SDK v4)
  ├── Intent Handler (NLP commands)
  ├── Azure Cosmos DB (steps, milestones, audit, users)
  ├── SharePoint Excel Sync (Graph API, 15-min interval)
  ├── Dependency Engine (DAG critical path)
  ├── Notification Engine (deadlines, blockers, digests)
  └── Automation Webhooks (Power Automate, ADO)
```

## Quick Start

```bash
npm install
npm run build
npm run import-excel   # Validate Excel parsing
npm start              # Start the bot server
```

## Deployment

See [`docs/azure-portal-setup.md`](docs/azure-portal-setup.md) for full Azure Portal provisioning guide, or run:

```powershell
.\deploy.ps1 -ResourceGroup "rg-qqia-agent" -Location "westus2"
```

## Data Source

Reads from `FY27_Mint_RolloverTimeline.xlsx` with 14 sheets:
- **FY27_Rollover** (182 steps) — primary tracker
- **HighLevelMilestones** (30 milestones)
- **KeyBusinessDates**, **RAID Log**, **Decision Log**, and more

## Cost

| Resource | Monthly |
|----------|---------|
| Azure App Service (B1) | ~$13 (free via VS Enterprise) |
| Cosmos DB (Serverless) | ~$0-2 |
| Bot Service | Free |
| **Total** | **~$13-15** |
