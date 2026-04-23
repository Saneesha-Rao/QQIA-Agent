import {
  TeamsActivityHandler,
  TurnContext,
  CardFactory,
  MessageFactory,
  TeamsInfo,
} from 'botbuilder';
import { DataService } from '../services/dataService';
import { DependencyEngine } from '../services/dependencyEngine';
import { ExcelSyncService } from '../services/excelSyncService';
import { PowerAutomateSyncService } from '../services/powerAutomateSyncService';
import { ExcelComSyncService } from '../services/excelComSyncService';
import { NotificationService } from '../services/notificationService';
import { RolloverStep } from '../models/types';
import {
  buildOverallDashboardCard,
  buildStepDetailCard,
  buildStepListCard,
  buildMyTasksCard,
} from '../cards/adaptiveCards';

/**
 * QQIA Bot - Teams bot handler for FY27 Mint Rollover management.
 * Handles natural language intents for status updates, queries, and dashboards.
 */
export class QQIABot extends TeamsActivityHandler {
  private dataService: DataService;
  private dependencyEngine: DependencyEngine;
  private excelSync: ExcelSyncService;
  private notificationService: NotificationService;
  private paSyncService?: PowerAutomateSyncService;
  private comSync?: ExcelComSyncService;
  /** Track last viewed step per conversation for follow-up questions */
  private lastViewedStep: Map<string, string> = new Map();

  constructor(
    dataService: DataService,
    dependencyEngine: DependencyEngine,
    excelSync: ExcelSyncService,
    notificationService: NotificationService,
    paSyncService?: PowerAutomateSyncService,
    comSync?: ExcelComSyncService
  ) {
    super();
    this.dataService = dataService;
    this.dependencyEngine = dependencyEngine;
    this.excelSync = excelSync;
    this.notificationService = notificationService;
    this.paSyncService = paSyncService;
    this.comSync = comSync;

    // Handle incoming messages
    this.onMessage(async (context, next) => {
      await this.handleMessage(context);
      await next();
    });

    // Handle card action submissions
    this.onAdaptiveCardInvoke = async (context) => {
      return await this.handleCardAction(context);
    };

    // Welcome new members
    this.onMembersAdded(async (context, next) => {
      for (const member of context.activity.membersAdded || []) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(
            `👋 Welcome to the **QQIA Agent**!\n\n` +
            `I help manage the FY27 Mint Rollover process. Here's what I can do:\n\n` +
            `- **"dashboard"** → Overall rollover progress\n` +
            `- **"my tasks"** → Your assigned steps\n` +
            `- **"status [step ID]"** → Check a specific step\n` +
            `- **"update [step ID] completed"** → Update step status\n` +
            `- **"blockers"** → View all blocked steps\n` +
            `- **"overdue"** → View overdue steps\n` +
            `- **"workstream [name]"** → View workstream status\n` +
            `- **"critical path"** → View the critical path\n` +
            `- **"summary"** → Leadership summary\n` +
            `- **"help"** → Show all commands`
          );
        }
      }
      await next();
    });
  }

  /** Main message handler - routes intents to appropriate handlers */
  private async handleMessage(context: TurnContext): Promise<void> {
    // Handle Adaptive Card Action.Submit button clicks (value is set, text is empty)
    const actionData = context.activity.value;
    if (actionData?.action) {
      await this.handleCardSubmitAction(context, actionData);
      return;
    }

    // Normalize: lowercase, collapse extra spaces, fix common typos
    let text = (context.activity.text || '').trim().toLowerCase().replace(/\s+/g, ' ');
    if (!text) {
      await context.sendActivity('Please type a command or **help** to see available options.');
      return;
    }
    // Fix common typos for key commands
    text = text.replace(/\bstatis\b/g, 'status').replace(/\bstaus\b/g, 'status')
               .replace(/\budpate\b/g, 'update').replace(/\bupdat\b/g, 'update')
               .replace(/\bcomlete\b/g, 'complete').replace(/\bcomplte\b/g, 'complete')
               .replace(/\btaks\b/g, 'tasks').replace(/\btask\b/g, 'tasks')
               .replace(/\bactivites\b/g, 'activities').replace(/\bactivitys\b/g, 'activities')
               .replace(/\bworksteam\b/g, 'workstream').replace(/\bworkstrem\b/g, 'workstream');

    // Synonym expansion — normalize common synonyms before intent routing
    text = text.replace(/\bfinished\b/g, 'completed').replace(/\bfinish\b/g, 'complete')
               .replace(/\bstuck\b/g, 'blocked').replace(/\bhold\b/g, 'blocked').replace(/\bon hold\b/g, 'blocked')
               .replace(/\bpending\b/g, 'not started').replace(/\bnot begun\b/g, 'not started').replace(/\bwaiting\b/g, 'not started')
               .replace(/\bin-progress\b/g, 'in progress').replace(/\bstarted\b/g, 'in progress').replace(/\bunderway\b/g, 'in progress').replace(/\bongoing\b/g, 'in progress')
               .replace(/\bassigned to\b/g, 'owner').replace(/\bowned by\b/g, 'owner')
               .replace(/\bteam\b/g, 'workstream').replace(/\bgroup\b/g, 'workstream')
               .replace(/\bdelayed\b/g, 'overdue').replace(/\bpast due\b/g, 'overdue').replace(/\bslipped\b/g, 'overdue')
               .replace(/\bdeps\b/g, 'dependencies').replace(/\bprereqs?\b/g, 'dependencies').replace(/\bprerequisites?\b/g, 'dependencies')
               .replace(/\btimeline\b/g, 'upcoming').replace(/\bschedule\b/g, 'upcoming').replace(/\bcalendar\b/g, 'upcoming')
               .replace(/\bprogress report\b/g, 'summary').replace(/\bstatus report\b/g, 'summary');
    const userName = context.activity.from.name || 'Unknown';

    try {
      // Greeting & chitchat — handle before intent routing
      if (/^(hi|hello|hey|howdy|good morning|good afternoon|good evening|yo)\b/.test(text)) {
        await context.sendActivity(`👋 Hi ${userName}! Type **help** to see what I can do, or just ask me about FY27 rollover.`);
        return;
      }
      if (/^(thanks|thank you|thx|ty|cheers|appreciate it)\b/.test(text)) {
        await context.sendActivity(`You're welcome! Let me know if you need anything else. 😊`);
        return;
      }
      if (/^(bye|goodbye|see you|later|cya|good night)\b/.test(text)) {
        await context.sendActivity(`See you later, ${userName}! 👋`);
        return;
      }

      // Fuzzy command matching — catch typos like "dashbord", "sumary", "blokced"
      const firstWord = text.split(' ')[0];
      const fuzzyTarget = this.fuzzyMatchCommand(firstWord);
      if (fuzzyTarget && fuzzyTarget !== firstWord) {
        text = fuzzyTarget + text.slice(firstWord.length);
      }

      // Intent routing
      if (text === 'help' || text === '?') {
        await this.handleHelp(context);
      } else if (text === 'dashboard' || text === 'dash' || text.startsWith('show dashboard')) {
        await this.handleDashboard(context);
      } else if (text === 'my tasks' || text === 'my steps' || text === 'my items') {
        await this.handleMyTasks(context, userName);
      } else if (text.startsWith('status ') || text.startsWith('step ') || text.startsWith('show ')) {
        // Check if it's a negative query before treating as step lookup
        if (this.isNegativeQuery(text)) {
          await this.handleNegativeQuery(context, text);
        } else if (this.isDateRangeQuery(text)) {
          await this.handleDateRangeQuery(context, text);
        } else {
          await this.handleStepQuery(context, text);
        }
      } else if (text.startsWith('update ') || text.startsWith('mark ') || text.startsWith('complete ')) {
        await this.handleStatusUpdate(context, text, userName);
      } else if (text === 'blockers' || text === 'blocked' || text === 'show blockers') {
        await this.handleBlockers(context);
      } else if (text === 'overdue' || text === 'show overdue') {
        await this.handleOverdue(context);
      } else if (text.startsWith('workstream ') || text.startsWith('ws ')) {
        await this.handleWorkstream(context, text);
      } else if (text === 'critical path' || text === 'cp') {
        await this.handleCriticalPath(context);
      } else if (text === 'summary' || text === 'exec summary' || text.startsWith('leadership')) {
        await this.handleSummary(context);
      } else if (text === 'upcoming' || text === 'coming up' || text === 'next steps' || text.startsWith('upcoming ')) {
        const daysMatch = text.match(/(\d+)\s*days?/);
        await this.handleUpcoming(context, daysMatch ? parseInt(daysMatch[1]) : 7);
      } else if (text.startsWith('tasks for ') || text.startsWith('owner ')) {
        await this.handleOwnerTasks(context, text);
      } else if (text === 'sync' || text === 'refresh') {
        await this.handleSync(context);
      } else if (text === 'changes' || text === 'changelog' || text.startsWith('what changed') || text.startsWith('show changes')) {
        await this.handleWhatChanged(context, text);
      } else if (text.startsWith('note ') || text.startsWith('add note ')) {
        await this.handleAddNote(context, text, userName);
      } else if (text.startsWith('fed ') || text.startsWith('fed status ')) {
        await this.handleFedQuery(context, text);
      } else if (text.startsWith('dependencies on ') || text.startsWith('deps on ') || text.startsWith('dependency on ')) {
        await this.handleTeamDependencies(context, text);
      } else {
        await this.handleNaturalLanguage(context, text, userName);
      }
    } catch (error: any) {
      console.error('Error handling message:', error);
      await context.sendActivity(`❌ Error: ${error.message}. Please try again or type **help**.`);
    }
  }

  // ---- Intent Handlers ----

  private async handleHelp(context: TurnContext): Promise<void> {
    await context.sendActivity(
      `## 📚 QQIA Agent Commands\n\n` +
      `### Status & Queries\n` +
      `- **dashboard** → Overall rollover progress card\n` +
      `- **my tasks** → Your assigned steps\n` +
      `- **status 1.A** → View step 1.A details\n` +
      `- **workstream System Rollover** → View workstream\n` +
      `- **tasks for Jim R** → View a person's tasks\n` +
      `- **dependencies on SPM** → Steps involving a team\n` +
      `- **blockers** → All blocked steps\n` +
      `- **overdue** → All overdue steps\n` +
      `- **critical path** → Current critical path\n` +
      `- **summary** → Leadership executive summary\n` +
      `- **changes** → What changed in the last 24 hours\n` +
      `- **unassigned** → Steps with no owner\n\n` +
      `### Updates\n` +
      `- **update 1.A completed** → Mark step complete (Corp)\n` +
      `- **mark 1.A, 1.B, 1.C completed** → Batch update\n` +
      `- **update 1.A in progress** → Mark in progress\n` +
      `- **update 1.A blocked** → Mark blocked\n` +
      `- **fed update 1.A completed** → Update Fed status\n` +
      `- **note 1.A <your note>** → Add a note to a step\n\n` +
      `### Date Filters\n` +
      `- **due before May 1** → Steps due by a date\n` +
      `- **what changed today** → Recent modifications\n` +
      `- **due this month** → Current month deadlines\n\n` +
      `### Admin\n` +
      `- **sync** → Trigger Excel sync now\n` +
      `- **help** → Show this menu`
    );
  }

  private async handleDashboard(context: TurnContext, track: 'Corp' | 'Fed' = 'Corp'): Promise<void> {
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);

    const stats = this.dependencyEngine.getWorkstreamStats(allSteps, track);
    const overdue = this.dependencyEngine.getOverdueSteps(allSteps, track);
    const blocked = allSteps.filter(s =>
      (track === 'Corp' ? s.corpStatus : s.fedStatus) === 'Blocked'
    );

    const card = buildOverallDashboardCard(allSteps, stats, overdue, blocked, track);
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    await this.sendSuggestedActions(context, 'dashboard');
  }

  private async handleMyTasks(context: TurnContext, userName: string): Promise<void> {
    const steps = await this.dataService.getStepsByOwner(userName);
    if (steps.length === 0) {
      await context.sendActivity(`No steps found for **${userName}**. Try "tasks for [name]" with the exact name from the tracker.`);
      return;
    }
    const card = buildMyTasksCard(userName, steps);
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    await this.sendSuggestedActions(context, 'my_tasks');
  }

  private async handleStepQuery(context: TurnContext, text: string): Promise<void> {
    const stepId = this.extractStepId(text);
    if (!stepId) {
      await context.sendActivity('Please provide a step ID, e.g., **status 1.A**');
      return;
    }

    const step = await this.dataService.getStep(stepId);
    if (!step) {
      await context.sendActivity(`Step **${stepId}** not found. Check the step ID and try again.`);
      return;
    }

    // Track this step as the last viewed for follow-up questions
    const convId = context.activity.conversation.id;
    this.lastViewedStep.set(convId, stepId);

    this.dependencyEngine.buildGraph(await this.dataService.getAllSteps());
    const blockers = this.dependencyEngine.getBlockers(stepId);
    const blockedBy = this.dependencyEngine.getBlockedBy(stepId);

    const card = buildStepDetailCard(step, blockers, blockedBy);
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    await this.sendSuggestedActions(context, 'step_detail');
  }

  private async handleStatusUpdate(context: TurnContext, text: string, userName: string): Promise<void> {
    // Extract ALL step IDs for batch support (e.g. "mark 1.A, 1.B, and 1.C as completed")
    const stepIds = this.extractAllStepIds(text);
    if (stepIds.length === 0) {
      await context.sendActivity('Please provide a step ID, e.g., **update 1.A completed**');
      return;
    }

    let newStatus = 'In Progress';
    const lowerText = text.toLowerCase();
    if (lowerText.includes('completed') || lowerText.includes('complete') || lowerText.includes('done')) {
      newStatus = 'Completed';
    } else if (lowerText.includes('blocked') || lowerText.includes('block')) {
      newStatus = 'Blocked';
    } else if (lowerText.includes('in progress') || lowerText.includes('started') || lowerText.includes('start')) {
      newStatus = 'In Progress';
    } else if (lowerText.includes('not started') || lowerText.includes('reset')) {
      newStatus = 'Not Started';
    }

    const isFed = lowerText.startsWith('fed ');
    const field = isFed ? 'fedStatus' : 'corpStatus';
    const track = isFed ? 'Fed' : 'Corp';

    // Single step — original path
    if (stepIds.length === 1) {
      const stepId = stepIds[0];
      const updated = await this.dataService.updateStepStatus(stepId, field as any, newStatus, userName, 'bot');
      if (!updated) {
        await context.sendActivity(`Step **${stepId}** not found.`);
        return;
      }

      // Immediately sync to Excel file (COM → Power Automate → local fallback)
      let excelSynced = false;
      try {
        if (this.comSync?.isAvailable) {
          await this.comSync.syncToExcel();
          excelSynced = true;
        } else if (this.paSyncService?.isConfigured) {
          await this.paSyncService.syncToExcel();
          excelSynced = true;
        } else {
          await this.excelSync.syncToExcel();
        }
      } catch {
        // Sync failed — will show reminder below
      }

      let msg = `✅ Step **${stepId}** ${track} status updated to **${newStatus}** by ${userName}.`;
      if (newStatus === 'Completed') {
        msg += `\nCompleted date set to ${new Date().toISOString().split('T')[0]}.`;
      }
      if (excelSynced) {
        msg += `\n📊 Excel file updated.`;
        await context.sendActivity(msg);
      } else {
        const card = {
          type: 'AdaptiveCard',
          $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
          version: '1.5',
          body: [
            { type: 'TextBlock', text: msg, wrap: true },
            { type: 'TextBlock', text: '📊 Excel not yet synced — click below to sync now.', wrap: true, isSubtle: true, spacing: 'Small' },
          ],
          actions: [
            { type: 'Action.Submit', title: '🔄 Sync to Excel', data: { action: 'sync_excel' } },
          ],
        };
        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
      }

      // Trigger dependency notifications if completed
      if (newStatus === 'Completed') {
        const notifications = await this.notificationService.notifyPredecessorComplete(stepId, track as any);
        if (notifications.length > 0) {
          await context.sendActivity(
            `🔓 ${notifications.length} step(s) are now unblocked: ${notifications.map(n => n.steps[0]?.id).join(', ')}`
          );
        }
      }
      return;
    }

    // Batch update — multiple step IDs
    const succeeded: string[] = [];
    const failed: string[] = [];
    for (const sid of stepIds) {
      const ok = await this.dataService.updateStepStatus(sid, field as any, newStatus, userName, 'bot');
      if (ok) { succeeded.push(sid); } else { failed.push(sid); }
    }

    // Try Excel sync once after all updates
    let excelSynced = false;
    try {
      if (this.comSync?.isAvailable) {
        await this.comSync.syncToExcel();
        excelSynced = true;
      } else if (this.paSyncService?.isConfigured) {
        await this.paSyncService.syncToExcel();
        excelSynced = true;
      } else {
        await this.excelSync.syncToExcel();
      }
    } catch { /* sync failed */ }

    let msg = `✅ **Batch update**: ${succeeded.length} step(s) set to **${newStatus}** (${track}) by ${userName}.`;
    if (succeeded.length > 0) msg += `\n Updated: ${succeeded.join(', ')}`;
    if (failed.length > 0) msg += `\n⚠️ Not found: ${failed.join(', ')}`;
    if (newStatus === 'Completed') msg += `\nCompleted date set to ${new Date().toISOString().split('T')[0]}.`;

    if (excelSynced) {
      msg += `\n📊 Excel file updated.`;
      await context.sendActivity(msg);
    } else {
      const card = {
        type: 'AdaptiveCard',
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.5',
        body: [
          { type: 'TextBlock', text: msg, wrap: true },
          { type: 'TextBlock', text: '📊 Excel not yet synced — click below to sync now.', wrap: true, isSubtle: true, spacing: 'Small' },
        ],
        actions: [
          { type: 'Action.Submit', title: '🔄 Sync to Excel', data: { action: 'sync_excel' } },
        ],
      };
      await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    }

    // Notify dependencies for each completed step
    if (newStatus === 'Completed') {
      for (const sid of succeeded) {
        const notifications = await this.notificationService.notifyPredecessorComplete(sid, track as any);
        if (notifications.length > 0) {
          await context.sendActivity(
            `🔓 ${notifications.length} step(s) unblocked by ${sid}: ${notifications.map(n => n.steps[0]?.id).join(', ')}`
          );
        }
      }
    }
  }

  private async handleBlockers(context: TurnContext): Promise<void> {
    const blocked = await this.dataService.getStepsByStatus('Blocked', 'Corp');
    if (blocked.length === 0) {
      await context.sendActivity('✅ No blocked steps! Everything is moving forward.');
      return;
    }
    const card = buildStepListCard('🛑 Blocked Steps', blocked, 'Corp');
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    await this.sendSuggestedActions(context, 'blockers');
  }

  private async handleOverdue(context: TurnContext): Promise<void> {
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);
    const overdue = this.dependencyEngine.getOverdueSteps(allSteps, 'Corp');

    if (overdue.length === 0) {
      await context.sendActivity('✅ No overdue steps! All on track.');
      return;
    }
    const card = buildStepListCard('⚠️ Overdue Steps', overdue, 'Corp');
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    await this.sendSuggestedActions(context, 'overdue');
  }

  private async handleWorkstream(context: TurnContext, text: string): Promise<void> {
    const wsName = text.replace(/^(workstream|ws)\s+/i, '').trim();
    const steps = await this.dataService.getStepsByWorkstream(wsName);
    if (steps.length === 0) {
      // Try fuzzy match
      const allSteps = await this.dataService.getAllSteps();
      const workstreams = [...new Set(allSteps.map(s => s.workstream))];
      const match = workstreams.find(ws => ws.toLowerCase().includes(wsName.toLowerCase()));
      if (match) {
        const matchedSteps = allSteps.filter(s => s.workstream === match);
        const card = buildStepListCard(`📦 ${match}`, matchedSteps, 'Corp');
        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
        return;
      }
      await context.sendActivity(
        `Workstream "${wsName}" not found. Available: ${workstreams.join(', ')}`
      );
      return;
    }
    const card = buildStepListCard(`📦 ${wsName}`, steps, 'Corp');
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
  }

  private async handleCriticalPath(context: TurnContext): Promise<void> {
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);
    const path = this.dependencyEngine.getCriticalPath('Corp');

    if (path.length === 0) {
      await context.sendActivity('No critical path found (all steps may be complete or independent).');
      return;
    }

    const pathSteps = [];
    for (const id of path) {
      const step = await this.dataService.getStep(id);
      if (step) pathSteps.push(step);
    }

    let msg = `## 🔗 Critical Path (${path.length} steps)\n\n`;
    for (let i = 0; i < pathSteps.length; i++) {
      const s = pathSteps[i];
      const arrow = i < pathSteps.length - 1 ? ' →' : '';
      msg += `**${s.id}** ${s.description} [${s.corpStatus}]${arrow}\n`;
    }
    await context.sendActivity(msg);
  }

  private async handleSummary(context: TurnContext): Promise<void> {
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);

    const total = allSteps.length;
    const completed = allSteps.filter(s => s.corpStatus === 'Completed').length;
    const inProgress = allSteps.filter(s => s.corpStatus === 'In Progress').length;
    const blocked = allSteps.filter(s => s.corpStatus === 'Blocked').length;
    const overdue = this.dependencyEngine.getOverdueSteps(allSteps, 'Corp');
    const upcoming = this.dependencyEngine.getUpcomingDeadlines(allSteps, 7, 'Corp');
    const pct = Math.round((completed / total) * 100);

    const fedCompleted = allSteps.filter(s => s.fedStatus === 'Completed').length;
    const fedPct = Math.round((fedCompleted / total) * 100);

    let summary = `## 📊 FY27 Rollover Executive Summary\n\n`;
    summary += `### Overall Progress\n`;
    summary += `- **Corp**: ${completed}/${total} steps complete (${pct}%)\n`;
    summary += `- **Fed**: ${fedCompleted}/${total} steps complete (${fedPct}%)\n`;
    summary += `- **In Progress**: ${inProgress} | **Blocked**: ${blocked}\n`;
    summary += `- **Overdue**: ${overdue.length} | **Due this week**: ${upcoming.length}\n\n`;

    if (overdue.length > 0) {
      summary += `### ⚠️ Top Overdue Items\n`;
      for (const s of overdue.slice(0, 5)) {
        summary += `- **${s.id}** ${s.description} (Due: ${s.corpEndDate}, Owner: ${s.wwicPoc || s.engineeringDri})\n`;
      }
      summary += '\n';
    }

    if (blocked > 0) {
      const blockedSteps = allSteps.filter(s => s.corpStatus === 'Blocked');
      summary += `### 🛑 Blocked Items\n`;
      for (const s of blockedSteps.slice(0, 5)) {
        summary += `- **${s.id}** ${s.description} (Owner: ${s.wwicPoc || s.engineeringDri})\n`;
      }
    }

    await context.sendActivity(summary);
    await this.sendSuggestedActions(context, 'summary');
  }

  private async handleOwnerTasks(context: TurnContext, text: string): Promise<void> {
    const name = text.replace(/^(tasks for|owner)\s+/i, '').trim();
    const steps = await this.dataService.getStepsByOwner(name);
    if (steps.length === 0) {
      await context.sendActivity(`No steps found for **${name}**.`);
      return;
    }
    const card = buildMyTasksCard(name, steps);
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
  }

  private async handleSync(context: TurnContext): Promise<void> {
    await context.sendActivity('🔄 Syncing with Excel...');
    const result = await this.excelSync.fullSync();
    await context.sendActivity(
      `✅ Sync complete!\n- **${result.fromExcel}** updates from Excel\n- **${result.toExcel}** updates pushed to Excel`
    );
  }

  private async handleAddNote(context: TurnContext, text: string, userName: string): Promise<void> {
    const match = text.match(/^(?:add\s+)?note\s+(\d+\.\w+(?:\.\d+)?)\s+(.+)/i);
    if (!match) {
      await context.sendActivity('Usage: **note 1.A your note text here**');
      return;
    }
    const [, stepId, note] = match;
    const step = await this.dataService.getStep(stepId);
    if (!step) {
      await context.sendActivity(`Step **${stepId}** not found.`);
      return;
    }
    const oldNotes = step.referenceNotes;
    step.referenceNotes = step.referenceNotes
      ? `${step.referenceNotes}\n[${new Date().toISOString().split('T')[0]} ${userName}]: ${note}`
      : `[${new Date().toISOString().split('T')[0]} ${userName}]: ${note}`;
    step.lastModified = new Date().toISOString();
    step.lastModifiedBy = userName;
    step.lastModifiedSource = 'bot';
    await this.dataService.upsertStep(step);
    await this.dataService.logAudit(stepId, 'referenceNotes', oldNotes, step.referenceNotes, userName, 'bot');
    await context.sendActivity(`📝 Note added to step **${stepId}**.`);
  }

  private async handleFedQuery(context: TurnContext, text: string): Promise<void> {
    const cleaned = text.replace(/^fed\s+(status\s+)?/i, '').trim();
    if (cleaned.startsWith('update') || cleaned.startsWith('mark') || cleaned.startsWith('complete')) {
      await this.handleStatusUpdate(context, `fed ${cleaned}`, context.activity.from.name || 'Unknown');
    } else {
      // Show fed dashboard
      await this.handleDashboard(context, 'Fed');
    }
  }

  /** Handle queries about dependencies on a team or keyword (e.g. "dependencies on SPM team") */
  private async handleTeamDependencies(context: TurnContext, text: string): Promise<void> {
    // Extract the team/keyword from the query
    let keyword = text
      .replace(/^(what are the |what's the |show |list |get )?/i, '')
      .replace(/^(dependencies|dependency|deps)\s+(on|for|of)\s+/i, '')
      .replace(/\s*(team|group|org)\s*$/i, '')
      .replace(/[?]/g, '')
      .trim();

    if (!keyword) {
      await context.sendActivity('Please specify a team or keyword, e.g., **dependencies on SPM**');
      return;
    }

    const allSteps = await this.dataService.getAllSteps();
    const lower = keyword.toLowerCase();

    // Search across description, engineeringDependent, owner fields, and workstream
    const matchingSteps = allSteps.filter(s =>
      s.description.toLowerCase().includes(lower) ||
      s.engineeringDependent.toLowerCase().includes(lower) ||
      s.wwicPoc.toLowerCase().includes(lower) ||
      s.fedPoc.toLowerCase().includes(lower) ||
      s.engineeringDri.toLowerCase().includes(lower) ||
      s.engineeringLead.toLowerCase().includes(lower) ||
      s.workstream.toLowerCase().includes(lower) ||
      s.referenceNotes.toLowerCase().includes(lower)
    );

    if (matchingSteps.length === 0) {
      await context.sendActivity(`No steps found related to **${keyword}**. Try a different team or keyword.`);
      return;
    }

    // Build dependency graph for blocker info
    this.dependencyEngine.buildGraph(allSteps);

    // Group by status for a useful summary
    const completed = matchingSteps.filter(s => s.corpStatus === 'Completed');
    const inProgress = matchingSteps.filter(s => s.corpStatus === 'In Progress');
    const notStarted = matchingSteps.filter(s => s.corpStatus === 'Not Started');
    const blocked = matchingSteps.filter(s => s.corpStatus === 'Blocked');

    let msg = `## 🔗 Dependencies on "${keyword}" (${matchingSteps.length} steps)\n\n`;
    msg += `**Completed**: ${completed.length} | **In Progress**: ${inProgress.length} | **Not Started**: ${notStarted.length} | **Blocked**: ${blocked.length}\n\n`;

    const formatSteps = (steps: RolloverStep[], limit: number = 15) => {
      let result = '';
      for (const s of steps.slice(0, limit)) {
        const owner = s.wwicPoc || s.engineeringDri || s.engineeringLead || 'Unassigned';
        const dates = s.corpEndDate ? ` (Due: ${s.corpEndDate})` : '';
        const blockers = this.dependencyEngine.getBlockers(s.id);
        const blockerInfo = blockers.length > 0 ? ` ⛔ Blocked by: ${blockers.join(', ')}` : '';
        result += `- **${s.id}** ${s.description} [${s.corpStatus}]${dates} — ${owner}${blockerInfo}\n`;
      }
      if (steps.length > limit) result += `  _...and ${steps.length - limit} more_\n`;
      return result;
    };

    if (blocked.length > 0) {
      msg += `### 🛑 Blocked\n${formatSteps(blocked)}\n`;
    }
    if (inProgress.length > 0) {
      msg += `### 🔄 In Progress\n${formatSteps(inProgress)}\n`;
    }
    if (notStarted.length > 0) {
      msg += `### ⏳ Not Started\n${formatSteps(notStarted)}\n`;
    }
    if (completed.length > 0) {
      msg += `### ✅ Completed\n${formatSteps(completed)}\n`;
    }

    await context.sendActivity(msg);
  }

  /** Handle follow-up questions referencing the last viewed step */
  private async handleFollowUp(context: TurnContext, text: string, stepId: string): Promise<boolean> {
    const step = await this.dataService.getStep(stepId);
    if (!step) return false;

    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);

    // "what are dependencies for this?" / "what does this depend on?" / "what blocks this?"
    if (text.includes('dependenc') || text.includes('depend') || text.includes('block') || text.includes('prerequisite') || text.includes('predecessor')) {
      const blockers = this.dependencyEngine.getBlockers(stepId);
      const blockedBy = this.dependencyEngine.getBlockedBy(stepId);

      let msg = `## 🔗 Dependencies for Step ${stepId}\n**${step.description}**\n\n`;

      if (step.dependencies.length > 0) {
        msg += `### Prerequisites (${step.dependencies.length})\n`;
        for (const depId of step.dependencies) {
          const dep = await this.dataService.getStep(depId);
          if (dep) {
            const status = dep.corpStatus;
            const icon = status === 'Completed' ? '✅' : status === 'In Progress' ? '🔄' : status === 'Blocked' ? '🛑' : '⏳';
            msg += `- ${icon} **${dep.id}** ${dep.description} [${status}]\n`;
          } else {
            msg += `- **${depId}** (not found)\n`;
          }
        }
      } else {
        msg += `### Prerequisites\nNone — this step has no prerequisites.\n`;
      }

      msg += '\n';

      if (blockedBy.length > 0) {
        msg += `### Blocks These Steps (${blockedBy.length})\n`;
        for (const bid of blockedBy) {
          const bStep = await this.dataService.getStep(bid);
          if (bStep) {
            msg += `- **${bStep.id}** ${bStep.description} [${bStep.corpStatus}]\n`;
          }
        }
      } else {
        msg += `### Blocks These Steps\nNo downstream steps are waiting on this.\n`;
      }

      if (blockers.length > 0) {
        msg += `\n⛔ **Currently blocked by**: ${blockers.join(', ')}\n`;
      }

      await context.sendActivity(msg);
      return true;
    }

    // "who owns this?" / "who is responsible?"
    if (text.includes('who') || text.includes('owner') || text.includes('responsible') || text.includes('assigned') || text.includes('poc') || text.includes('dri')) {
      let msg = `## 👤 Owners for Step ${stepId}\n**${step.description}**\n\n`;
      if (step.wwicPoc) msg += `- **WWIC POC**: ${step.wwicPoc}\n`;
      if (step.fedPoc) msg += `- **Fed POC**: ${step.fedPoc}\n`;
      if (step.engineeringDri) msg += `- **Engineering DRI**: ${step.engineeringDri}\n`;
      if (step.engineeringLead) msg += `- **Engineering Lead**: ${step.engineeringLead}\n`;
      if (!step.wwicPoc && !step.fedPoc && !step.engineeringDri && !step.engineeringLead) {
        msg += `No owners assigned.\n`;
      }
      await context.sendActivity(msg);
      return true;
    }

    // "when is this due?" / "what's the timeline?"
    if (text.includes('when') || text.includes('due') || text.includes('date') || text.includes('timeline') || text.includes('schedule')) {
      let msg = `## 📅 Timeline for Step ${stepId}\n**${step.description}**\n\n`;
      msg += `- **Corp**: ${step.corpStartDate || 'TBD'} → ${step.corpEndDate || 'TBD'} [${step.corpStatus}]\n`;
      msg += `- **Fed**: ${step.fedStartDate || 'TBD'} → ${step.fedEndDate || 'TBD'} [${step.fedStatus}]\n`;
      if (step.corpCompletedDate) msg += `- **Completed**: ${step.corpCompletedDate}\n`;
      await context.sendActivity(msg);
      return true;
    }

    // "what's the status?" / "is it done?"
    if (text.includes('status') || text.includes('done') || text.includes('complete') || text.includes('progress')) {
      await this.handleStepQuery(context, `status ${stepId}`);
      return true;
    }

    // Generic "tell me about this" / "more info" / "details"
    if (text.includes('about') || text.includes('info') || text.includes('detail') || text.includes('more') || text.includes('explain') || text.includes('describe')) {
      await this.handleStepQuery(context, `status ${stepId}`);
      return true;
    }

    return false;
  }

  /** Fallback handler for natural language queries */
  private async handleNaturalLanguage(context: TurnContext, text: string, userName: string): Promise<void> {
    // Check if user typed a bare step ID (e.g., "1.C" or "2.D")
    const stepId = this.extractStepId(text);
    if (stepId && text.replace(/\s/g, '').length <= stepId.length + 2) {
      // Bare step ID — treat as status query
      await this.handleStepQuery(context, `status ${stepId}`);
      return;
    }

    // Check if message contains "to" pattern like "update 1.C to completed"
    // Also handle batch: "mark 1.A, 1.B, 1.C as completed"
    const allIds = this.extractAllStepIds(text);
    const hasStatusWord = /\b(completed|complete|done|in progress|blocked|not started|reset)\b/.test(text);
    if (allIds.length > 0 && hasStatusWord) {
      await this.handleStatusUpdate(context, text, userName);
      return;
    }
    if (stepId && (text.includes(' to ') || hasStatusWord)) {
      await this.handleStatusUpdate(context, text, userName);
      return;
    }

    // Resolve "this"/"it"/"that" to the last viewed step for follow-up questions
    const convId = context.activity.conversation.id;
    const lastStep = this.lastViewedStep.get(convId);
    const refersToContext = /\b(this|it|that|the step|this step|that step|this one)\b/.test(text);

    if (refersToContext && lastStep && !stepId) {
      if (await this.handleFollowUp(context, text, lastStep)) {
        return;
      }
    }

    // Detect personal pronouns — "I", "my", "me" → filter to user's tasks
    const isPersonal = /\b(i |i'[a-z]|my |me |mine)\b/.test(text) || text.startsWith('do i ') || text.startsWith('what do i');

    // Detect third-person name references like "Amit Tiwari's tasks", "what does John need to"
    let mentionedPerson: string | null = null;
    const possessiveMatch = text.match(/(?:what(?:'s| is| are)\s+)?(\w[\w\s]*?)'s\s+(?:tasks?|activities|steps|items|work|upcoming)/i);
    const forMatch = text.match(/(?:tasks?|activities|steps|work|upcoming)\s+(?:for|of|assigned to)\s+(\w[\w\s]*?)(?:\s+(?:this|next|that|due|need)|[?.]|$)/i);
    const doesMatch = text.match(/what\s+(?:does|should|will|can)\s+(\w[\w\s]*?)\s+(?:need|have|do|start|complete|own)/i);
    if (possessiveMatch) mentionedPerson = possessiveMatch[1].trim();
    else if (forMatch) mentionedPerson = forMatch[1].trim();
    else if (doesMatch) mentionedPerson = doesMatch[1].trim();

    // Detect time-window keywords and extract days
    const timeMatch = text.match(/(?:this|next)\s+(\d+\s+)?(week|month|sprint)|(?:next|coming|upcoming)\s+(\d+)\s+days?|(?:in the next|within)\s+(\d+)\s+days?/i);
    let lookAheadDays = 7;
    if (timeMatch) {
      if (timeMatch[2] === 'month') lookAheadDays = 30;
      else if (timeMatch[3]) lookAheadDays = parseInt(timeMatch[3]);
      else if (timeMatch[4]) lookAheadDays = parseInt(timeMatch[4]);
    }

    // Third-person name + activity/task query → show THEIR tasks
    if (mentionedPerson && (text.includes('task') || text.includes('activities') || text.includes('step') ||
        text.includes('need') || text.includes('start') || text.includes('do') ||
        text.includes('due') || text.includes('upcoming') || text.includes('work') ||
        text.includes('this week') || text.includes('this month'))) {
      await this.handleMyUpcoming(context, mentionedPerson, lookAheadDays);
    // Personal + activity/task/upcoming query → show MY upcoming tasks
    } else if (isPersonal && (text.includes('activities') || text.includes('task') || text.includes('upcoming') ||
        text.includes('need to') || text.includes('start') || text.includes('do') ||
        text.includes('due') || text.includes('pending') || text.includes('this week') ||
        text.includes('next week') || text.includes('this month') || text.includes('assigned') ||
        text.includes('own') || text.includes('working') || text.includes('responsible'))) {
      await this.handleMyUpcoming(context, userName, lookAheadDays);
    // Upcoming / due / activities (non-personal) → show all
    } else if (text.includes('upcoming') || text.includes('coming up') || text.includes('what\'s next') ||
        text.includes('whats next') || text.includes('need to be completed') || text.includes('needs to be') ||
        text.includes('activities') || text.includes('pending') ||
        text.includes('scheduled') || text.includes('planned') || text.includes('remaining') ||
        (text.includes('what') && (text.includes('due') || text.includes('deadline') || text.includes('next') || text.includes('left') || text.includes('remain'))) ||
        text.includes('due soon') || text.includes('due this') || text.includes('due next') ||
        text.includes('this week') || text.includes('next week') || text.includes('this month')) {
      await this.handleUpcoming(context, lookAheadDays);
    } else if (text.includes('how many') || text.includes('count') || text.includes('total') ||
               text.includes('progress') || text.includes('overview') || text.includes('where are we') ||
               text.includes('how are we') || text.includes('status update') || text.includes('update me') ||
               text.includes('how much') || text.includes('how far') || text.includes('readiness') ||
               (text.includes('ready') && (text.includes('how') || text.includes('what')))) {
      // Check if there's a keyword filter (e.g., "how much of FY27 train is ready")
      const progressKeywords = this.extractProgressKeywords(text);
      if (progressKeywords) {
        await this.handleFilteredProgress(context, progressKeywords);
      } else {
        await this.handleSummary(context);
      }
    } else if (text.includes('who') && text.includes('own')) {
      await context.sendActivity('Try: **tasks for [name]** to see tasks for a specific person.');
    } else if (text.includes('blocker') || text.includes('blocking') || text.includes('stuck') || text.includes('at risk')) {
      await this.handleBlockers(context);
    } else if (text.includes('overdue') || text.includes('late') || text.includes('behind') || text.includes('missed')) {
      await this.handleOverdue(context);
    } else if (text.includes('summary') || text.includes('leadership') || text.includes('exec') || text.includes('report')) {
      await this.handleSummary(context);
    } else if (text.includes('dependenc') || text.includes('depends on') || text.includes('dependent on') || text.includes('deps on')) {
      await this.handleTeamDependencies(context, text);
    // "What changed" / audit queries — "what changed today", "who updated steps", "recent changes"
    } else if (/\b(changed|changelog|change log|changes|modified|updated|edited|who updated|who changed|recent changes)\b/.test(text)) {
      await this.handleWhatChanged(context, text);
    // Negative / inverse queries — "what hasn't started", "unassigned steps", "no owner"
    } else if (this.isNegativeQuery(text)) {
      await this.handleNegativeQuery(context, text);
    // Date-range queries — "due before May 1", "changed since yesterday", "steps due this month"
    } else if (this.isDateRangeQuery(text)) {
      await this.handleDateRangeQuery(context, text);
    } else {
      // Last resort: keyword search across all step descriptions
      const found = await this.handleKeywordSearch(context, text);
      if (!found) {
        await context.sendActivity(
          `I didn't understand that. Here are some things you can try:\n\n` +
          `- **dashboard** → Overall progress\n` +
          `- **my tasks** → Your assigned steps\n` +
          `- **upcoming** → Activities due this week\n` +
          `- **status 1.A** → Check a step\n` +
          `- **update 1.A completed** → Update a step\n` +
          `- **overdue** → Past-due steps\n` +
          `- **blockers** → Blocked steps\n` +
          `- **summary** → Leadership summary\n` +
          `- **help** → Full command list`
        );
      }
    }
  }

  /** Handle "what changed" / audit queries */
  private async handleWhatChanged(context: TurnContext, text: string): Promise<void> {
    // Determine time range from text
    const dateRange = this.parseNaturalDate(text);
    let sinceDate: Date;
    let label: string;

    if (dateRange?.after) {
      sinceDate = dateRange.after;
      label = dateRange.before
        ? `${sinceDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} – ${dateRange.before.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}`
        : `since ${sinceDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}`;
    } else {
      // Default: last 24 hours
      sinceDate = new Date(Date.now() - 24 * 60 * 60 * 1000);
      label = 'last 24 hours';
    }

    // Try audit log first
    const auditEntries = await this.dataService.getRecentChanges(sinceDate.toISOString(), 50);

    if (auditEntries.length > 0) {
      // Group by step
      const byStep = new Map<string, typeof auditEntries>();
      for (const entry of auditEntries) {
        if (dateRange?.before && entry.changedAt >= dateRange.before.toISOString()) continue;
        const list = byStep.get(entry.stepId) || [];
        list.push(entry);
        byStep.set(entry.stepId, list);
      }

      let msg = `## 📝 Changes — ${label}\n\n`;
      msg += `**${auditEntries.length} change(s)** across **${byStep.size} step(s)**\n\n`;

      for (const [sid, entries] of byStep) {
        msg += `### Step ${sid}\n`;
        for (const e of entries.slice(0, 5)) {
          const time = new Date(e.changedAt).toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
          msg += `- ${time}: **${e.field}** ${e.previousValue} → **${e.newValue}** _(by ${e.changedBy} via ${e.source})_\n`;
        }
        if (entries.length > 5) msg += `  _...and ${entries.length - 5} more_\n`;
        msg += '\n';
      }

      await context.sendActivity(msg);
    } else {
      // Fall back to lastModified field on steps
      const allSteps = await this.dataService.getAllSteps();
      const changed = allSteps.filter(s => {
        if (!s.lastModified) return false;
        const d = new Date(s.lastModified);
        if (d < sinceDate) return false;
        if (dateRange?.before && d >= dateRange.before) return false;
        return true;
      }).sort((a, b) => new Date(b.lastModified).getTime() - new Date(a.lastModified).getTime());

      if (changed.length === 0) {
        await context.sendActivity(`No changes found in the ${label}. Try **changes this week** for a wider window.`);
        await this.sendSuggestedActions(context, 'changes');
        return;
      }

      let msg = `## 📝 Steps Modified — ${label}\n\n`;
      msg += `**${changed.length} step(s) changed**\n\n`;
      for (const s of changed.slice(0, 20)) {
        const modDate = new Date(s.lastModified).toLocaleString('en-US', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
        msg += `- **${s.id}** ${s.description} — [${s.corpStatus}] modified by ${s.lastModifiedBy || 'system'} at ${modDate}\n`;
      }
      if (changed.length > 20) msg += `\n_...and ${changed.length - 20} more_\n`;

      await context.sendActivity(msg);
    }

    await this.sendSuggestedActions(context, 'changes');
  }

  /** Send context-aware suggested action buttons after a response */
  private async sendSuggestedActions(context: TurnContext, afterIntent: string): Promise<void> {
    const suggestions: { title: string; action: string }[] = [];

    switch (afterIntent) {
      case 'dashboard':
        suggestions.push(
          { title: '🛑 Blockers', action: 'blockers' },
          { title: '⏰ Overdue', action: 'overdue' },
          { title: '📋 My Tasks', action: 'my tasks' },
          { title: '📝 Changes', action: 'changes' },
        );
        break;
      case 'step_detail':
        suggestions.push(
          { title: '📊 Dashboard', action: 'dashboard' },
          { title: '🛑 Blockers', action: 'blockers' },
          { title: '📅 Upcoming', action: 'upcoming' },
        );
        break;
      case 'blockers':
        suggestions.push(
          { title: '📊 Dashboard', action: 'dashboard' },
          { title: '⏰ Overdue', action: 'overdue' },
          { title: '📅 Upcoming', action: 'upcoming' },
        );
        break;
      case 'overdue':
        suggestions.push(
          { title: '📊 Dashboard', action: 'dashboard' },
          { title: '🛑 Blockers', action: 'blockers' },
          { title: '📋 Summary', action: 'summary' },
        );
        break;
      case 'my_tasks':
        suggestions.push(
          { title: '📊 Dashboard', action: 'dashboard' },
          { title: '📅 Upcoming', action: 'upcoming' },
          { title: '🛑 Blockers', action: 'blockers' },
        );
        break;
      case 'summary':
        suggestions.push(
          { title: '📊 Dashboard', action: 'dashboard' },
          { title: '⏰ Overdue', action: 'overdue' },
          { title: '📝 Changes', action: 'changes' },
        );
        break;
      case 'changes':
        suggestions.push(
          { title: '📊 Dashboard', action: 'dashboard' },
          { title: '📋 Summary', action: 'summary' },
          { title: '📅 Upcoming', action: 'upcoming' },
        );
        break;
      case 'upcoming':
        suggestions.push(
          { title: '📊 Dashboard', action: 'dashboard' },
          { title: '🛑 Blockers', action: 'blockers' },
          { title: '📋 My Tasks', action: 'my tasks' },
        );
        break;
      default:
        suggestions.push(
          { title: '📊 Dashboard', action: 'dashboard' },
          { title: '📋 My Tasks', action: 'my tasks' },
          { title: '❓ Help', action: 'help' },
        );
    }

    const card = {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.5',
      body: [
        { type: 'TextBlock', text: '💡 Quick actions:', size: 'Small', isSubtle: true },
      ],
      actions: suggestions.map(s => ({
        type: 'Action.Submit',
        title: s.title,
        data: { action: 'quick_action', text: s.action },
      })),
    };
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
  }

  /** Detect negative/inverse query patterns */
  private isNegativeQuery(text: string): boolean {
    return /\b(hasn't|haven't|hasn't|not started|not begun|no owner|unassigned|without owner|not completed|incomplete|not done|not finished|missing|empty)\b/.test(text) ||
           /\b(which|what|show|list|find)\b.*\b(no |not |un|without )\b/.test(text);
  }

  /** Handle negative/inverse queries like "what hasn't started", "unassigned steps" */
  private async handleNegativeQuery(context: TurnContext, text: string): Promise<void> {
    const allSteps = await this.dataService.getAllSteps();
    let filtered: typeof allSteps = [];
    let title = '';

    if (/\b(hasn't started|haven't started|not started|not begun)\b/.test(text)) {
      filtered = allSteps.filter(s => s.corpStatus === 'Not Started');
      title = '⏳ Steps Not Started';
    } else if (/\b(not completed|not done|not finished|incomplete|hasn't been completed|haven't completed)\b/.test(text)) {
      filtered = allSteps.filter(s => s.corpStatus !== 'Completed' && s.corpStatus !== 'N/A');
      title = '📋 Incomplete Steps';
    } else if (/\b(no owner|unassigned|without owner|no poc|not assigned|nobody)\b/.test(text)) {
      filtered = allSteps.filter(s => !s.wwicPoc && !s.engineeringDri && !s.fedPoc);
      title = '❓ Unassigned Steps (No Owner)';
    } else if (/\b(no date|no deadline|without date|missing date|no end date)\b/.test(text)) {
      filtered = allSteps.filter(s => !s.corpEndDate);
      title = '📅 Steps Without Due Dates';
    } else if (/\b(no dependencies|no deps|independent|standalone)\b/.test(text)) {
      filtered = allSteps.filter(s => !s.dependencies || s.dependencies.length === 0);
      title = '🔗 Steps With No Dependencies';
    } else {
      // Generic negative — try "not <status>"
      filtered = allSteps.filter(s => s.corpStatus === 'Not Started');
      title = '⏳ Steps Not Started';
    }

    if (filtered.length === 0) {
      await context.sendActivity(`✅ No matching steps found for that filter.`);
      return;
    }

    // Show max 25 via step list card, with count
    const display = filtered.slice(0, 25);
    const card = buildStepListCard(`${title} (${filtered.length} total)`, display, 'Corp');
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    if (filtered.length > 25) {
      await context.sendActivity(`_Showing first 25 of ${filtered.length} steps._`);
    }
  }

  /** Detect date-range query patterns */
  private isDateRangeQuery(text: string): boolean {
    return /\b(due|changed|modified|updated|created)\b.*\b(before|after|since|by|until|between|today|yesterday|this month|next month)\b/.test(text) ||
           /\b(before|after|since|by)\s+\w+\s+\d{1,2}/.test(text) ||
           /\bwhat('s| is| was)?\s+(due|changed|modified)/.test(text);
  }

  /** Handle date-range queries like "steps due before May 1", "what changed today" */
  private async handleDateRangeQuery(context: TurnContext, text: string): Promise<void> {
    const allSteps = await this.dataService.getAllSteps();
    const dateRange = this.parseNaturalDate(text);

    if (!dateRange) {
      await context.sendActivity('I couldn\'t parse that date. Try: **due before May 1**, **changed since yesterday**, or **due this month**.');
      return;
    }

    const isChangedQuery = /\b(changed|modified|updated)\b/.test(text);

    let filtered: typeof allSteps;
    let title: string;

    if (isChangedQuery) {
      // Filter by lastModified date
      filtered = allSteps.filter(s => {
        if (!s.lastModified) return false;
        const d = new Date(s.lastModified);
        if (isNaN(d.getTime())) return false;
        if (dateRange.after && d < dateRange.after) return false;
        if (dateRange.before && d >= dateRange.before) return false;
        return true;
      });
      const rangeLabel = dateRange.after && dateRange.before
        ? `${dateRange.after.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} – ${dateRange.before.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}`
        : dateRange.after ? `since ${dateRange.after.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}`
        : `before ${dateRange.before!.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}`;
      title = `📝 Steps Changed ${rangeLabel}`;
    } else {
      // Filter by corpEndDate (due date)
      filtered = allSteps.filter(s => {
        if (!s.corpEndDate) return false;
        if (s.corpStatus === 'Completed' || s.corpStatus === 'N/A') return false;
        const d = new Date(s.corpEndDate);
        if (isNaN(d.getTime())) return false;
        if (dateRange.after && d < dateRange.after) return false;
        if (dateRange.before && d >= dateRange.before) return false;
        return true;
      }).sort((a, b) => new Date(a.corpEndDate!).getTime() - new Date(b.corpEndDate!).getTime());
      const rangeLabel = dateRange.after && dateRange.before
        ? `${dateRange.after.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} – ${dateRange.before.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}`
        : dateRange.after ? `after ${dateRange.after.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}`
        : `by ${dateRange.before!.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}`;
      title = `📅 Steps Due ${rangeLabel}`;
    }

    if (filtered.length === 0) {
      await context.sendActivity(`No steps found matching that date range. Try **upcoming** to see what's coming up.`);
      return;
    }

    const display = filtered.slice(0, 25);
    const card = buildStepListCard(`${title} (${filtered.length} total)`, display, 'Corp');
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    if (filtered.length > 25) {
      await context.sendActivity(`_Showing first 25 of ${filtered.length} steps._`);
    }
  }

  /** Extract meaningful keywords from a progress query like "how much of FY27 train is ready" → ["fy27", "train"] */
  private extractProgressKeywords(text: string): string[] | null {
    const progressStops = new Set([
      'how', 'much', 'many', 'far', 'is', 'are', 'the', 'a', 'an', 'of', 'for',
      'ready', 'done', 'complete', 'completed', 'progress', 'status', 'along',
      'what', 'where', 'we', 'on', 'with', 'has', 'have', 'been', 'get', 'got',
      'readiness', 'overview', 'update', 'total', 'count',
      'fy27', 'fy26', 'fy25', 'rollover', // these are too generic — almost every step matches
    ]);
    const words = text.replace(/[?!.,;:'"]/g, '').split(/\s+/)
      .filter(w => w.length >= 2 && !progressStops.has(w));
    return words.length > 0 ? words : null;
  }

  /** Show a progress summary filtered to steps matching keywords */
  private async handleFilteredProgress(context: TurnContext, keywords: string[]): Promise<void> {
    const allSteps = await this.dataService.getAllSteps();

    // Find steps matching keywords — use AND logic when multiple keywords, falling back to OR
    let matching = allSteps.filter(step => {
      const searchText = [
        step.description, step.workstream, step.engineeringDependent,
        step.wwicPoc, step.fedPoc, step.engineeringDri, step.engineeringLead,
        step.referenceNotes,
      ].join(' ').toLowerCase();
      return keywords.every(kw => searchText.includes(kw));
    });

    // If AND yields no results, fall back to scoring by keyword match count
    if (matching.length === 0) {
      const scored = allSteps.map(step => {
        const searchText = [
          step.description, step.workstream, step.engineeringDependent,
          step.wwicPoc, step.fedPoc, step.engineeringDri, step.engineeringLead,
          step.referenceNotes,
        ].join(' ').toLowerCase();
        const score = keywords.filter(kw => searchText.includes(kw)).length;
        return { step, score };
      }).filter(s => s.score > 0);

      if (scored.length > 0) {
        const maxScore = Math.max(...scored.map(s => s.score));
        matching = scored.filter(s => s.score === maxScore).map(s => s.step);
      }
    }

    if (matching.length === 0) {
      await context.sendActivity(`No steps found matching **${keywords.join(' ')}**. Try **dashboard** for overall progress.`);
      return;
    }

    const completed = matching.filter(s => s.corpStatus === 'Completed');
    const inProgress = matching.filter(s => s.corpStatus === 'In Progress');
    const notStarted = matching.filter(s => s.corpStatus === 'Not Started');
    const blocked = matching.filter(s => s.corpStatus === 'Blocked');
    const na = matching.filter(s => s.corpStatus === 'N/A');
    const pct = Math.round(((completed.length + na.length) / matching.length) * 100);

    let msg = `## 📊 Progress: "${keywords.join(' ')}" (${matching.length} steps)\n\n`;
    msg += `### ${pct}% Complete\n`;
    msg += `✅ Completed: **${completed.length}**  |  🔄 In Progress: **${inProgress.length}**  |  ⏳ Not Started: **${notStarted.length}**  |  🛑 Blocked: **${blocked.length}**`;
    if (na.length > 0) msg += `  |  N/A: **${na.length}**`;
    msg += '\n\n';

    // Show incomplete steps with details
    const incomplete = [...blocked, ...inProgress, ...notStarted];
    if (incomplete.length > 0 && incomplete.length <= 20) {
      msg += `### Remaining Steps\n`;
      for (const s of incomplete) {
        const icon = s.corpStatus === 'Blocked' ? '🛑' : s.corpStatus === 'In Progress' ? '🔄' : '⏳';
        const owner = s.wwicPoc || s.engineeringDri || 'Unassigned';
        const dates = s.corpEndDate ? ` (Due: ${s.corpEndDate})` : '';
        msg += `- ${icon} **${s.id}** ${s.description} [${s.corpStatus}]${dates} — ${owner}\n`;
      }
    } else if (incomplete.length > 20) {
      msg += `### Next Up\n`;
      for (const s of incomplete.slice(0, 10)) {
        const icon = s.corpStatus === 'Blocked' ? '🛑' : s.corpStatus === 'In Progress' ? '🔄' : '⏳';
        const owner = s.wwicPoc || s.engineeringDri || 'Unassigned';
        msg += `- ${icon} **${s.id}** ${s.description} [${s.corpStatus}] — ${owner}\n`;
      }
      msg += `_...and ${incomplete.length - 10} more_\n`;
    }

    if (completed.length > 0 && completed.length <= 10) {
      msg += `\n### ✅ Completed\n`;
      for (const s of completed) {
        msg += `- **${s.id}** ${s.description}${s.corpCompletedDate ? ` (${s.corpCompletedDate})` : ''}\n`;
      }
    } else if (completed.length > 10) {
      msg += `\n_${completed.length} steps already completed._\n`;
    }

    await context.sendActivity(msg);
  }

  /** Keyword search fallback — extract meaningful words from the query and search step descriptions.
   *  Returns true if matching steps were found and displayed. */
  private async handleKeywordSearch(context: TurnContext, text: string): Promise<boolean> {
    // Strip common stop words to get meaningful keywords
    const stopWords = new Set([
      'what', 'when', 'where', 'who', 'how', 'is', 'are', 'the', 'a', 'an',
      'for', 'of', 'in', 'on', 'to', 'and', 'or', 'will', 'can', 'do', 'does',
      'did', 'be', 'been', 'has', 'have', 'had', 'was', 'were', 'it', 'its',
      'this', 'that', 'with', 'from', 'by', 'about', 'happen', 'happening',
      'date', 'dates', 'tell', 'me', 'show', 'give', 'get', 'find', 'look',
      'please', 'need', 'want', 'like', 'know', 'would', 'could', 'should',
      'there', 'here', 'any', 'all', 'some', 'which', 'also', 'just',
    ]);

    const words = text.replace(/[?!.,;:'"]/g, '').split(/\s+/)
      .filter(w => w.length >= 2 && !stopWords.has(w));

    if (words.length === 0) return false;

    const allSteps = await this.dataService.getAllSteps();

    // Score each step by how many keywords match its description and fields
    const scored: { step: RolloverStep; score: number }[] = [];
    for (const step of allSteps) {
      const searchText = [
        step.description,
        step.workstream,
        step.engineeringDependent,
        step.wwicPoc,
        step.fedPoc,
        step.engineeringDri,
        step.engineeringLead,
        step.referenceNotes,
      ].join(' ').toLowerCase();

      let score = 0;
      for (const word of words) {
        if (searchText.includes(word)) score++;
      }
      // Bonus: check if the full multi-word query (minus stop words) matches as a phrase
      const phrase = words.join(' ');
      if (phrase.length >= 3 && searchText.includes(phrase)) score += 2;

      if (score > 0) scored.push({ step, score });
    }

    if (scored.length === 0) return false;

    // Sort by score descending, then by step ID
    scored.sort((a, b) => b.score - a.score || a.step.id.localeCompare(b.step.id, undefined, { numeric: true }));

    const matches = scored.map(s => s.step);
    const queryHasWhen = text.includes('when') || text.includes('date') || text.includes('timeline') || text.includes('schedule');
    const queryHasStatus = text.includes('status') || text.includes('progress') || text.includes('done') || text.includes('ready');

    // If only 1 match, show the detail card directly
    if (matches.length === 1) {
      const s = matches[0];
      const convId = context.activity.conversation.id;
      this.lastViewedStep.set(convId, s.id);

      if (queryHasWhen) {
        let msg = `## 📅 Step ${s.id}: ${s.description}\n\n`;
        msg += `- **Corp**: ${s.corpStartDate || 'TBD'} → ${s.corpEndDate || 'TBD'} [${s.corpStatus}]\n`;
        msg += `- **Fed**: ${s.fedStartDate || 'TBD'} → ${s.fedEndDate || 'TBD'} [${s.fedStatus}]\n`;
        if (s.corpCompletedDate) msg += `- **Completed**: ${s.corpCompletedDate}\n`;
        msg += `\n**Owner**: ${s.wwicPoc || s.engineeringDri || 'Unassigned'}`;
        await context.sendActivity(msg);
      } else {
        this.dependencyEngine.buildGraph(allSteps);
        const blockers = this.dependencyEngine.getBlockers(s.id);
        const blockedBy = this.dependencyEngine.getBlockedBy(s.id);
        const card = buildStepDetailCard(s, blockers, blockedBy);
        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
      }
      return true;
    }

    // Multiple matches — if asking "when", show dates in a list
    if (queryHasWhen && matches.length <= 10) {
      let msg = `## 📅 Found ${matches.length} steps matching "${words.join(' ')}"\n\n`;
      for (const s of matches) {
        const owner = s.wwicPoc || s.engineeringDri || 'Unassigned';
        msg += `- **${s.id}** ${s.description}\n  Corp: ${s.corpStartDate || 'TBD'} → ${s.corpEndDate || 'TBD'} [${s.corpStatus}] — ${owner}\n`;
      }
      await context.sendActivity(msg);
      // Track first match as context for follow-ups
      this.lastViewedStep.set(context.activity.conversation.id, matches[0].id);
      return true;
    }

    // Multiple matches — show as a step list card
    const topMatches = matches.slice(0, 20);
    const card = buildStepListCard(`🔍 Found ${matches.length} steps matching "${words.join(' ')}"`, topMatches, 'Corp');
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    this.lastViewedStep.set(context.activity.conversation.id, matches[0].id);
    return true;
  }

  /** Show upcoming activities due within N days */
  private async handleUpcoming(context: TurnContext, days: number = 7): Promise<void> {
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);

    const now = new Date();
    const cutoff = new Date(now.getTime() + days * 86400000);
    const active = allSteps.filter(s => {
      if (s.corpStatus === 'Completed' || s.corpStatus === 'N/A') return false;
      const endDate = s.corpEndDate ? new Date(s.corpEndDate) : null;
      return endDate && endDate >= now && endDate <= cutoff;
    }).sort((a, b) => {
      const da = a.corpEndDate ? new Date(a.corpEndDate).getTime() : Infinity;
      const db = b.corpEndDate ? new Date(b.corpEndDate).getTime() : Infinity;
      return da - db;
    });

    if (active.length > 0) {
      const label = days <= 7 ? 'This Week' : `Next ${days} Days`;
      const card = buildStepListCard(`📅 Upcoming Activities — ${label} (${active.length} steps)`, active, 'Corp');
      await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    } else {
      // Fall back to showing all not-started / in-progress items
      const pending = allSteps.filter(s => s.corpStatus === 'Not Started' || s.corpStatus === 'In Progress')
        .sort((a, b) => {
          const da = a.corpEndDate ? new Date(a.corpEndDate).getTime() : Infinity;
          const db = b.corpEndDate ? new Date(b.corpEndDate).getTime() : Infinity;
          return da - db;
        }).slice(0, 20);
      if (pending.length > 0) {
        const card = buildStepListCard(`📅 Next Pending Activities (${pending.length} shown)`, pending, 'Corp');
        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
      } else {
        await context.sendActivity(`✅ All activities are completed!`);
      }
    }
    await this.sendSuggestedActions(context, 'upcoming');
  }

  /** Show upcoming activities for a specific user within N days */
  private async handleMyUpcoming(context: TurnContext, userName: string, days: number = 7): Promise<void> {
    const mySteps = await this.dataService.getStepsByOwner(userName);

    if (mySteps.length === 0) {
      await context.sendActivity(
        `No steps found assigned to **${userName}**.\n\n` +
        `Your Teams name must match the tracker. Try **tasks for [exact name]** or **my tasks**.`
      );
      return;
    }

    const now = new Date();
    const cutoff = new Date(now.getTime() + days * 86400000);

    // Get all incomplete steps sorted by start/end date
    const allPending = mySteps.filter(s => s.corpStatus !== 'Completed' && s.corpStatus !== 'N/A')
      .sort((a, b) => {
        const da = a.corpStartDate ? new Date(a.corpStartDate).getTime() : (a.corpEndDate ? new Date(a.corpEndDate).getTime() : Infinity);
        const db = b.corpStartDate ? new Date(b.corpStartDate).getTime() : (b.corpEndDate ? new Date(b.corpEndDate).getTime() : Infinity);
        return da - db;
      });

    // Filter to steps that START or END within the time window
    const inWindow = allPending.filter(s => {
      const startDate = s.corpStartDate ? new Date(s.corpStartDate) : null;
      const endDate = s.corpEndDate ? new Date(s.corpEndDate) : null;
      return (startDate && startDate >= now && startDate <= cutoff) ||
             (endDate && endDate >= now && endDate <= cutoff) ||
             (startDate && endDate && startDate <= now && endDate >= now); // currently active
    });

    const label = days <= 7 ? 'This Week' : `Next ${days} Days`;

    if (inWindow.length > 0) {
      const card = buildMyTasksCard(`${userName} — ${label}`, inWindow);
      await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
    } else {
      // Nothing in window — show next 5 upcoming with a helpful message
      const nextUp = allPending.filter(s => {
        const startDate = s.corpStartDate ? new Date(s.corpStartDate) : null;
        return startDate && startDate > now;
      }).slice(0, 5);

      if (nextUp.length > 0) {
        const firstDate = nextUp[0].corpStartDate || 'TBD';
        await context.sendActivity(
          `📅 **${userName}** has no tasks starting or due ${label.toLowerCase()}.\n\n` +
          `Their next tasks begin **${firstDate}**. Here are the next ${nextUp.length} upcoming:`
        );
        const card = buildMyTasksCard(`${userName} — Next Upcoming`, nextUp);
        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
      } else if (allPending.length > 0) {
        await context.sendActivity(`**${userName}** has ${allPending.length} pending tasks but none with scheduled dates yet.`);
      } else {
        await context.sendActivity(`✅ All activities are completed for **${userName}**! 🎉`);
      }
    }
  }

  /** Handle Action.Submit button clicks from Adaptive Cards (sent as message with value) */
  private async handleCardSubmitAction(context: TurnContext, data: any): Promise<void> {
    const userName = context.activity.from.name || 'Unknown';

    switch (data.action) {
      case 'update_status':
        await this.dataService.updateStepStatus(data.stepId, data.field, data.newStatus, userName, 'bot');
        await context.sendActivity(`✅ Step **${data.stepId}** updated to **${data.newStatus}**`);
        break;

      case 'view_step':
        const step = await this.dataService.getStep(data.stepId);
        if (step) {
          this.lastViewedStep.set(context.activity.conversation.id, data.stepId);
          this.dependencyEngine.buildGraph(await this.dataService.getAllSteps());
          const card = buildStepDetailCard(
            step,
            this.dependencyEngine.getBlockers(data.stepId),
            this.dependencyEngine.getBlockedBy(data.stepId)
          );
          await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
        } else {
          await context.sendActivity(`Step **${data.stepId}** not found.`);
        }
        break;

      case 'view_overdue':
        await this.handleOverdue(context);
        break;

      case 'view_blocked':
        await this.handleBlockers(context);
        break;

      case 'critical_path':
        await this.handleCriticalPath(context);
        break;

      case 'dashboard':
        await this.handleDashboard(context, data.track || 'Corp');
        break;

      case 'sync_excel':
        await this.handleSyncExcelAction(context);
        break;

      case 'quick_action':
        // Re-route suggested action button click as if user typed it
        if (data.text) {
          context.activity.text = data.text;
          context.activity.value = undefined; // Clear to prevent re-entering action handler
          await this.handleMessage(context);
        }
        break;

      default:
        await context.sendActivity(`Unknown action: ${data.action}. Type **help** for commands.`);
    }
  }

  /** Handle Adaptive Card action submissions */
  private async handleCardAction(context: TurnContext): Promise<any> {
    const data = context.activity.value?.action ? context.activity.value : context.activity.value?.data;
    if (!data?.action) return { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'No action found' };

    const userName = context.activity.from.name || 'Unknown';

    switch (data.action) {
      case 'update_status':
        await this.dataService.updateStepStatus(data.stepId, data.field, data.newStatus, userName, 'bot');
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.activity.message',
          value: `✅ Step ${data.stepId} updated to ${data.newStatus}`,
        };

      case 'view_step':
        const step = await this.dataService.getStep(data.stepId);
        if (step) {
          this.dependencyEngine.buildGraph(await this.dataService.getAllSteps());
          const card = buildStepDetailCard(
            step,
            this.dependencyEngine.getBlockers(data.stepId),
            this.dependencyEngine.getBlockedBy(data.stepId)
          );
          await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
        }
        return { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: '' };

      case 'view_overdue':
      case 'view_blocked':
      case 'critical_path':
      case 'dashboard':
        // Re-route to existing handlers
        const fakeText = data.action.replace('view_', '').replace('_', ' ');
        await this.handleMessage({ ...context, activity: { ...context.activity, text: fakeText } } as any);
        return { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: '' };

      case 'sync_excel':
        await this.handleSyncExcelAction(context);
        return { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: '' };

      case 'quick_action':
        if (data.text) {
          context.activity.text = data.text;
          context.activity.value = undefined; // Clear to prevent re-entering action handler
          await this.handleMessage(context);
        }
        return { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: '' };

      default:
        return { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'Unknown action' };
    }
  }

  /** Handle the Sync to Excel button click */
  private async handleSyncExcelAction(context: TurnContext): Promise<void> {
    try {
      if (this.comSync?.isAvailable) {
        await this.comSync.syncToExcel();
        await context.sendActivity('✅ Excel file synced successfully via COM.');
      } else if (this.paSyncService?.isConfigured) {
        await this.paSyncService.syncToExcel();
        await context.sendActivity('✅ Excel file synced successfully via Power Automate.');
      } else {
        await this.excelSync.syncToExcel();
        await context.sendActivity('✅ Excel file synced (local copy updated). Download from the web dashboard to get the latest file.');
      }
    } catch {
      await context.sendActivity(
        '⚠️ Automatic sync is unavailable. To sync manually:\n\n' +
        '1. Open your Excel file in the browser\n' +
        '2. Go to **Automate** tab → run the **QQIA Sync** script\n\n' +
        'Or download the latest file from the [web dashboard](../)'
      );
    }
  }

  // ---- Utilities ----

  /** Known command words for fuzzy matching */
  private static readonly KNOWN_COMMANDS = [
    'help', 'dashboard', 'status', 'update', 'mark', 'complete', 'blockers', 'blocked',
    'overdue', 'workstream', 'summary', 'upcoming', 'sync', 'refresh', 'note', 'tasks',
    'dependencies', 'critical', 'show', 'step', 'owner', 'fed',
  ];

  /** Levenshtein distance between two strings */
  private static levenshtein(a: string, b: string): number {
    const m = a.length, n = b.length;
    const dp: number[][] = Array.from({ length: m + 1 }, (_, i) =>
      Array.from({ length: n + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0))
    );
    for (let i = 1; i <= m; i++) {
      for (let j = 1; j <= n; j++) {
        dp[i][j] = a[i - 1] === b[j - 1]
          ? dp[i - 1][j - 1]
          : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
      }
    }
    return dp[m][n];
  }

  /** Find the closest known command word if within edit distance 2 */
  private fuzzyMatchCommand(word: string): string | null {
    if (word.length < 3) return null; // too short to fuzzy-match reliably
    let best: string | null = null;
    let bestDist = 3; // max distance threshold
    for (const cmd of QQIABot.KNOWN_COMMANDS) {
      const dist = QQIABot.levenshtein(word, cmd);
      if (dist === 0) return word; // exact match, no correction needed
      if (dist < bestDist) {
        bestDist = dist;
        best = cmd;
      }
    }
    return best;
  }

  /** Extract step ID from user message (e.g., "1.A", "2.B.1") — handles spaces around dots and normalizes to uppercase */
  private extractStepId(text: string): string | null {
    // First try exact match like "1.A" or "2.B.1"
    let match = text.match(/\b(\d+)\s*\.\s*(\w+(?:\s*\.\s*\d+)?)\b/);
    if (match) {
      // Remove all spaces and uppercase the letter part: "1 . c" → "1.C"
      return (match[1] + '.' + match[2]).replace(/\s/g, '').toUpperCase();
    }
    // Also try just a number+letter combo like "1A" → "1.A"
    match = text.match(/\b(\d+)([A-Za-z])\b/);
    if (match) {
      return match[1] + '.' + match[2].toUpperCase();
    }
    return null;
  }

  /** Extract ALL step IDs from text for batch operations (e.g., "1.A, 1.B, and 1.C") */
  private extractAllStepIds(text: string): string[] {
    const ids = new Set<string>();
    // Match "1.A" or "2.B.1" patterns globally
    const dotPattern = /\b(\d+)\s*\.\s*(\w+(?:\s*\.\s*\d+)?)\b/g;
    let m: RegExpExecArray | null;
    while ((m = dotPattern.exec(text)) !== null) {
      ids.add((m[1] + '.' + m[2]).replace(/\s/g, '').toUpperCase());
    }
    // Also match "1A" shorthand
    const shortPattern = /\b(\d+)([A-Za-z])\b/g;
    while ((m = shortPattern.exec(text)) !== null) {
      const id = m[1] + '.' + m[2].toUpperCase();
      if (!ids.has(id)) ids.add(id);
    }
    return Array.from(ids);
  }

  /** Parse a natural language date reference into a Date object */
  private parseNaturalDate(text: string): { before?: Date; after?: Date } | null {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const result: { before?: Date; after?: Date } = {};

    // "today"
    if (/\btoday\b/.test(text)) {
      result.after = today;
      result.before = new Date(today.getTime() + 86400000);
      return result;
    }
    // "yesterday"
    if (/\byesterday\b/.test(text)) {
      result.after = new Date(today.getTime() - 86400000);
      result.before = today;
      return result;
    }
    // "this week"
    if (/\bthis week\b/.test(text)) {
      const dayOfWeek = now.getDay();
      result.after = new Date(today.getTime() - dayOfWeek * 86400000);
      result.before = new Date(result.after.getTime() + 7 * 86400000);
      return result;
    }
    // "next week"
    if (/\bnext week\b/.test(text)) {
      const dayOfWeek = now.getDay();
      result.after = new Date(today.getTime() + (7 - dayOfWeek) * 86400000);
      result.before = new Date(result.after.getTime() + 7 * 86400000);
      return result;
    }
    // "this month"
    if (/\bthis month\b/.test(text)) {
      result.after = new Date(now.getFullYear(), now.getMonth(), 1);
      result.before = new Date(now.getFullYear(), now.getMonth() + 1, 1);
      return result;
    }
    // "next month"
    if (/\bnext month\b/.test(text)) {
      result.after = new Date(now.getFullYear(), now.getMonth() + 1, 1);
      result.before = new Date(now.getFullYear(), now.getMonth() + 2, 1);
      return result;
    }

    // "before <date>" / "by <date>"
    const beforeMatch = text.match(/\b(?:before|by|until)\s+(\w+\s+\d{1,2}(?:,?\s*\d{4})?)/i);
    if (beforeMatch) {
      const d = new Date(beforeMatch[1]);
      if (!isNaN(d.getTime())) { result.before = d; return result; }
    }
    // "after <date>" / "since <date>"
    const afterMatch = text.match(/\b(?:after|since|from)\s+(\w+\s+\d{1,2}(?:,?\s*\d{4})?)/i);
    if (afterMatch) {
      const d = new Date(afterMatch[1]);
      if (!isNaN(d.getTime())) { result.after = d; return result; }
    }

    return null;
  }
}
