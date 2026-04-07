# 📘 QQIA Agent — User Guide

**FY27 Mint Rollover Tracking Bot for Microsoft Teams**

---

## What Is the QQIA Agent?

The QQIA Agent is a Microsoft Teams bot that helps the FY27 Mint Rollover team track, update, and coordinate over **182 rollover steps** across **15 workstreams**. It works alongside the existing Excel tracker — any updates you make in Teams automatically sync back to the shared Excel file, and vice versa.

**You don't need to install anything.** Just message the bot in Teams like you would a colleague.

---

## 🚀 Getting Started

### Finding the Bot
1. Open **Microsoft Teams**
2. In the chat pane, search for **"QQIA Agent"**
3. Start a 1:1 chat with the bot
4. The bot will greet you and show available commands

> **Tip:** You can also @mention the bot in a team channel to use it publicly.

### First Message
Type **help** to see all available commands:

```
help
```

---

## 📊 Viewing Status & Dashboards

### Overall Dashboard
See a visual overview of the entire rollover — progress bars, workstream breakdown, alerts:

```
dashboard
```

The dashboard card shows:
- ✅ Completed / 🔄 In Progress / 🚫 Blocked / ⏳ Not Started counts
- Per-workstream progress
- Overdue and blocked step alerts
- Quick-action buttons to drill into details

### My Tasks
See all steps assigned to **you** (matched by your Teams display name):

```
my tasks
```

### Check a Specific Step
Look up any step by its ID (e.g., 1.A, 3.B.1, 7.D):

```
status 1.A
```

This shows the step's description, dates, Corp/Fed status, DRI, dependencies, and notes.

### View a Workstream
See all steps within a workstream:

```
workstream System Rollover
workstream Orchestration
ws Quota Issuance
```

> **Shortcut:** You can use `ws` instead of `workstream`.

### See Someone Else's Tasks
Check what steps are assigned to a specific person:

```
tasks for Jim R
tasks for Saneesha
owner Saneesha
```

> Use the name as it appears in the Excel tracker (WWIC POC, Fed POC, or Engineering DRI columns).

---

## ⚠️ Blockers, Overdue & Critical Path

### View All Blocked Steps
```
blockers
```

### View Overdue Steps
Steps past their end date that aren't yet completed:
```
overdue
```

### Critical Path
See the longest chain of dependent steps — the sequence that determines the earliest the rollover can complete:
```
critical path
```

---

## ✏️ Updating Step Status

### Mark a Step Complete (Corp)
```
update 1.A completed
mark 1.A done
complete 1.A
```

### Mark In Progress
```
update 1.A in progress
mark 1.A started
```

### Mark Blocked
```
update 1.A blocked
```

### Reset to Not Started
```
update 1.A not started
```

### Update Fed Track Status
Prefix any update command with **fed** to update the Fed status instead of Corp:
```
fed update 1.A completed
fed mark 3.B in progress
```

### Add a Note to a Step
Attach a comment or context to any step:
```
note 1.A Waiting on ADO work item approval from Finance team
add note 3.B.1 Discussed with Jim - will complete by Thursday
```

---

## 📋 Leadership & Summary Views

### Executive Summary
Get a high-level overview suitable for leadership updates:
```
summary
exec summary
leadership update
```

This provides:
- Overall completion percentage (Corp & Fed)
- Count of blocked/overdue items
- Workstream-by-workstream progress
- Key risks and upcoming deadlines

---

## 🏢 Corp vs Fed Tracking

Every step has **independent Corp and Fed statuses**. By default, commands show Corp track data.

| To see... | Command |
|-----------|---------|
| Corp dashboard | `dashboard` |
| Fed dashboard | `fed` or `fed status` |
| Update Corp status | `update 1.A completed` |
| Update Fed status | `fed update 1.A completed` |

---

## 🔔 Automatic Notifications

The bot proactively sends you messages — **you don't need to check manually**:

| Notification | When | Who Gets It |
|-------------|------|-------------|
| ⏰ **Deadline approaching** | 3 days and 1 day before a step is due | Step DRI/POC |
| 🚨 **Overdue alert** | When a step passes its due date | Step DRI/POC + PM |
| ✅ **Predecessor completed** | When a blocking step finishes | DRI of the now-unblocked step |
| 🔓 **Step unblocked** | All predecessors are done — you can start | Step DRI/POC |
| 📊 **Weekly digest** | Every Monday at 8:00 AM | All active DRIs/POCs |
| 🚩 **Escalation** | Step overdue by 3+ days | PM + Leadership |

> Notifications appear as 1:1 messages from the bot. No action needed to opt in — if you're a DRI or POC on any step, you'll get relevant alerts.

---

## 🔄 Excel Sync

The QQIA Agent stays synchronized with the shared **FY27_Mint_RolloverTimeline.xlsx** file on SharePoint:

- **Auto-sync every 15 minutes** — changes in Teams appear in Excel and vice versa
- **You can still update Excel directly** — the bot will pick up your changes
- **No data loss** — if both are edited, the most recent change wins and both versions are logged

### Trigger Manual Sync
If you need the latest data right now:
```
sync
```

---

## 💡 Natural Language Queries

Don't remember the exact command? The bot understands natural language too:

| What you type | What happens |
|--------------|-------------|
| *"How many steps are done?"* | Shows summary counts |
| *"What's due this week?"* | Lists steps with deadlines in the next 7 days |
| *"Who owns step 3.B?"* | Shows step details including DRI |

---

## 📌 Quick Reference Card

| Command | Description |
|---------|-------------|
| `help` | Show all commands |
| `dashboard` | Overall rollover progress |
| `my tasks` | Steps assigned to you |
| `status 1.A` | View step 1.A details |
| `update 1.A completed` | Mark 1.A as completed (Corp) |
| `update 1.A in progress` | Mark 1.A as in progress |
| `update 1.A blocked` | Mark 1.A as blocked |
| `fed update 1.A completed` | Mark 1.A as completed (Fed) |
| `note 1.A <text>` | Add a note to step 1.A |
| `workstream Orchestration` | View Orchestration workstream |
| `tasks for Jim R` | View Jim R's steps |
| `blockers` | All blocked steps |
| `overdue` | All overdue steps |
| `critical path` | Critical dependency chain |
| `summary` | Executive/leadership summary |
| `fed` | Fed track dashboard |
| `sync` | Trigger Excel sync now |

---

## ❓ FAQ

### Q: Do I need to install anything?
**No.** The bot runs in Microsoft Teams. Just search for "QQIA Agent" and start chatting.

### Q: Will my updates show up in the Excel tracker?
**Yes.** The bot syncs with the Excel file every 15 minutes. Your updates will appear in the shared spreadsheet automatically.

### Q: Can I still update the Excel file directly?
**Yes.** The sync is bi-directional. Update either Teams or Excel — both stay current.

### Q: What if the bot doesn't recognize my name for "my tasks"?
Your Teams display name must match the name in the Excel tracker columns (WWIC POC, Fed POC, or Engineering DRI). Try using `tasks for <name>` with the exact name from the tracker.

### Q: Can I use the bot in a team channel?
**Yes.** @mention the bot in any channel: `@QQIA Agent dashboard`. The response will be visible to everyone in the channel.

### Q: What step ID format should I use?
Step IDs follow the format from the tracker: `1.A`, `3.B`, `7.D.1`, etc. You can find them in the Excel file or by browsing a workstream.

### Q: Who can update step statuses?
DRIs and POCs can update their own steps. PMs can update any step. Leadership has view-only access.

### Q: What if I make a mistake in an update?
All changes are audit-logged. Simply update the step again to the correct status, or ask a PM to correct it.

---

## 🛠 For PMs & Admins

### View Any Person's Tasks
```
tasks for <name>
```

### Override Dependencies
PMs can update steps even if predecessors aren't complete.

### Automation Webhooks
External systems (ADO, pipelines) can auto-update step statuses via webhooks. See the [Power Automate Integration Guide](./power-automate-integration.md) for setup.

---

## 📞 Support

Having trouble? Reach out to:
- **Bot Admin**: Saneesha (salingal)
- **Teams Channel**: Post in the QQIA Rollover channel
- **GitHub Issues**: [Saneesha-Rao/QQIA-Agent](https://github.com/Saneesha-Rao/QQIA-Agent/issues)

---

*Last updated: April 2026 | QQIA Agent v1.0*
