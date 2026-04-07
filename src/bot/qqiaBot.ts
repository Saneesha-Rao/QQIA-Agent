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

  constructor(
    dataService: DataService,
    dependencyEngine: DependencyEngine,
    excelSync: ExcelSyncService,
    notificationService: NotificationService
  ) {
    super();
    this.dataService = dataService;
    this.dependencyEngine = dependencyEngine;
    this.excelSync = excelSync;
    this.notificationService = notificationService;

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
    const userName = context.activity.from.name || 'Unknown';

    try {
      // Intent routing
      if (text === 'help' || text === '?') {
        await this.handleHelp(context);
      } else if (text === 'dashboard' || text === 'dash' || text.startsWith('show dashboard')) {
        await this.handleDashboard(context);
      } else if (text === 'my tasks' || text === 'my steps' || text === 'my items') {
        await this.handleMyTasks(context, userName);
      } else if (text.startsWith('status ') || text.startsWith('step ') || text.startsWith('show ')) {
        await this.handleStepQuery(context, text);
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
      } else if (text.startsWith('note ') || text.startsWith('add note ')) {
        await this.handleAddNote(context, text, userName);
      } else if (text.startsWith('fed ') || text.startsWith('fed status ')) {
        await this.handleFedQuery(context, text);
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
      `- **blockers** → All blocked steps\n` +
      `- **overdue** → All overdue steps\n` +
      `- **critical path** → Current critical path\n` +
      `- **summary** → Leadership executive summary\n\n` +
      `### Updates\n` +
      `- **update 1.A completed** → Mark step complete (Corp)\n` +
      `- **update 1.A in progress** → Mark in progress\n` +
      `- **update 1.A blocked** → Mark blocked\n` +
      `- **fed update 1.A completed** → Update Fed status\n` +
      `- **note 1.A <your note>** → Add a note to a step\n\n` +
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
  }

  private async handleMyTasks(context: TurnContext, userName: string): Promise<void> {
    const steps = await this.dataService.getStepsByOwner(userName);
    if (steps.length === 0) {
      await context.sendActivity(`No steps found for **${userName}**. Try "tasks for [name]" with the exact name from the tracker.`);
      return;
    }
    const card = buildMyTasksCard(userName, steps);
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
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

    this.dependencyEngine.buildGraph(await this.dataService.getAllSteps());
    const blockers = this.dependencyEngine.getBlockers(stepId);
    const blockedBy = this.dependencyEngine.getBlockedBy(stepId);

    const card = buildStepDetailCard(step, blockers, blockedBy);
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
  }

  private async handleStatusUpdate(context: TurnContext, text: string, userName: string): Promise<void> {
    const stepId = this.extractStepId(text);
    if (!stepId) {
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

    const updated = await this.dataService.updateStepStatus(stepId, field as any, newStatus, userName, 'bot');
    if (!updated) {
      await context.sendActivity(`Step **${stepId}** not found.`);
      return;
    }

    const track = isFed ? 'Fed' : 'Corp';
    await context.sendActivity(
      `✅ Step **${stepId}** ${track} status updated to **${newStatus}** by ${userName}.` +
      (newStatus === 'Completed' ? `\nCompleted date set to ${new Date().toISOString().split('T')[0]}.` : '') +
      `\n📊 Excel will be synced within 15 minutes.`
    );

    // Trigger dependency notifications if completed
    if (newStatus === 'Completed') {
      const notifications = await this.notificationService.notifyPredecessorComplete(stepId, track as any);
      if (notifications.length > 0) {
        await context.sendActivity(
          `🔓 ${notifications.length} step(s) are now unblocked: ${notifications.map(n => n.steps[0]?.id).join(', ')}`
        );
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
    if (stepId && (text.includes(' to ') || text.includes('completed') || text.includes('done') || text.includes('in progress') || text.includes('blocked'))) {
      await this.handleStatusUpdate(context, text, userName);
      return;
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
               text.includes('how are we') || text.includes('status update') || text.includes('update me')) {
      await this.handleSummary(context);
    } else if (text.includes('who') && text.includes('own')) {
      await context.sendActivity('Try: **tasks for [name]** to see tasks for a specific person.');
    } else if (text.includes('blocker') || text.includes('blocking') || text.includes('stuck') || text.includes('at risk')) {
      await this.handleBlockers(context);
    } else if (text.includes('overdue') || text.includes('late') || text.includes('behind') || text.includes('missed')) {
      await this.handleOverdue(context);
    } else if (text.includes('summary') || text.includes('leadership') || text.includes('exec') || text.includes('report')) {
      await this.handleSummary(context);
    } else {
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

      default:
        return { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'Unknown action' };
    }
  }

  // ---- Utilities ----

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
}
