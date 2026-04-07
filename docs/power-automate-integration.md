# Power Automate Integration Guide for QQIA Agent

## Overview
The QQIA Agent exposes REST API webhooks that Power Automate flows can call
to automatically update step statuses based on external system events.

## Webhook Endpoints

### 1. Automation Status Update
**POST** `/api/automation/status`

Automatically updates a rollover step status when triggered by external events
(ADO work item state changes, pipeline completions, environment checks).

**Request body:**
```json
{
  "stepId": "1.A",
  "status": "Completed",
  "track": "Corp",
  "source": "power-automate-ado",
  "notes": "ADO work item #57003488 moved to Done"
}
```

**Status values:** `Not Started`, `In Progress`, `Completed`, `Blocked`

### 2. Manual Sync Trigger
**POST** `/api/sync`

Forces an immediate bi-directional Excel sync.

---

## Sample Power Automate Flows

### Flow 1: ADO Work Item → Auto-Update Step Status (P0)
**Trigger:** Azure DevOps — When a work item is updated
**Condition:** Work item state changed to "Done" or "Closed"
**Action:** HTTP POST to `/api/automation/status`

**Mapped Steps:**
| ADO Work Item | QQIA Step | Description |
|---------------|-----------|-------------|
| #57003488     | 1.F       | Rollover perspectives/entities |
| (environment) | 1.A       | Create FY27 bucket |
| (deployment)  | 1.D       | Deploy MintOrch dropdown |
| (container)   | 1.E       | Create datalake container |

### Flow 2: Scheduled Environment Check (P1)
**Trigger:** Recurrence — Every 4 hours
**Action:** HTTP request to Mint backend API to check FY27 bucket exists
**Condition:** If bucket exists → POST to `/api/automation/status`
```json
{ "stepId": "1.A", "status": "Completed", "source": "env-check" }
```

### Flow 3: Pipeline Completion → Orchestration Status (P1)
**Trigger:** Azure DevOps — When a pipeline run completes
**Condition:** Pipeline = "FY27_Rollover_Orchestration" AND result = "Succeeded"
**Action:** POST to `/api/automation/status`
```json
{ "stepId": "5.A", "status": "Completed", "source": "pipeline-monitor" }
```

### Flow 4: Daily Excel Sync Verification (P2)
**Trigger:** Recurrence — Daily at 6 AM
**Action:** POST to `/api/sync`
**Follow-up:** Email notification if sync reports conflicts

### Flow 5: Train Run Success Monitor (P1)
**Trigger:** When an ADF pipeline run succeeds
**Condition:** Pipeline name contains "FY27" AND "MintStudio"
**Action:** POST to `/api/automation/status`
```json
{ "stepId": "5.C", "status": "In Progress", "source": "adf-monitor", "notes": "Train run succeeded" }
```

---

## Automation Priority Catalog

### P0 — Immediate Value
| Step | Automation | Trigger |
|------|-----------|---------|
| 1.A  | Detect Mint bucket creation | Environment API check |
| 1.D  | Detect MintOrch dropdown deployment | Deployment pipeline |
| 1.E  | Detect datalake container creation | Storage API check |
| ADO-linked steps | Track ADO work item state changes | ADO webhook |

### P1 — High Value
| Step | Automation | Trigger |
|------|-----------|---------|
| 5.A-5.D | Orchestration train run completion | ADF pipeline events |
| 7.A-7.D | MintX rollover steps | MintX API checks |
| 6.A-6.D | Quota ML baseline completion | Pipeline completion |
| 8.A-8.F | Blueprint/EDM data checks | Scheduled API polls |

### P2 — Medium Value
| Step | Automation | Trigger |
|------|-----------|---------|
| 9.A-9.F | ICBM rollover steps | ICBM API monitoring |
| Daily processing | Auto-advance on schedule | Recurrence triggers |

### P3 — Nice to Have
| Step | Automation | Trigger |
|------|-----------|---------|
| Dependency auto-advance | When all predecessors complete, notify next owners | Sync engine (built-in) |
| Weekly digest | Auto-generate Monday summary | Built into bot (automatic) |

---

## Setup Instructions

1. Create a new Power Automate flow in your Microsoft 365 environment
2. Add the appropriate trigger (ADO, Recurrence, etc.)
3. Add an HTTP action pointing to `https://<your-bot-url>/api/automation/status`
4. Set the request body with the stepId, status, and source
5. Test the flow and verify the step updates in Teams and Excel

## Estimated Impact
- **12+ hours/week saved** during peak rollover
- **14+ steps** can be fully automated
- **Zero manual status tracking** for foundational infrastructure steps
