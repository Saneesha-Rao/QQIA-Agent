import * as crypto from 'crypto';
import { DataService } from './services/dataService';
import { InMemoryDataService } from './services/inMemoryDataService';
import { DependencyEngine } from './services/dependencyEngine';
import { ExcelSyncService } from './services/excelSyncService';
import { PowerAutomateSyncService } from './services/powerAutomateSyncService';
import { ExcelComSyncService } from './services/excelComSyncService';
import { NotificationService } from './services/notificationService';
import { RolloverStep } from './models/types';
import {
  buildOverallDashboardCard,
  buildStepDetailCard,
  buildStepListCard,
  buildMyTasksCard,
} from './cards/adaptiveCards';

/** Response format for Teams outgoing webhook */
interface WebhookResponse {
  type: 'message';
  text?: string;
  attachments?: Array<{
    contentType: string;
    content: any;
  }>;
}

/**
 * Handles Teams Outgoing Webhook requests.
 * Bypasses Bot Framework — no AAD app registration needed.
 * Routes commands identically to QQIABot but returns webhook-format responses.
 */
export class WebhookHandler {
  private dataService: DataService | InMemoryDataService;
  private dependencyEngine: DependencyEngine;
  private excelSync: ExcelSyncService;
  private notificationService: NotificationService;
  private paSyncService?: PowerAutomateSyncService;
  private comSync?: ExcelComSyncService;
  private hmacSecret?: string;

  constructor(
    dataService: DataService | InMemoryDataService,
    dependencyEngine: DependencyEngine,
    excelSync: ExcelSyncService,
    notificationService: NotificationService,
    paSyncService?: PowerAutomateSyncService,
    comSync?: ExcelComSyncService,
    hmacSecret?: string
  ) {
    this.dataService = dataService;
    this.dependencyEngine = dependencyEngine;
    this.excelSync = excelSync;
    this.notificationService = notificationService;
    this.paSyncService = paSyncService;
    this.comSync = comSync;
    this.hmacSecret = hmacSecret;
  }

  /** Validate HMAC-SHA256 signature from Teams */
  validateSignature(body: string, authHeader: string): boolean {
    if (!this.hmacSecret) return true; // Skip validation if no secret configured
    const msgBuf = Buffer.from(body, 'utf8');
    const msgHash = 'HMAC ' + crypto
      .createHmac('sha256', Buffer.from(this.hmacSecret, 'base64'))
      .update(msgBuf)
      .digest('base64');
    return authHeader === msgHash;
  }

  /** Process an incoming outgoing-webhook request and return a response */
  async processRequest(body: any, userName?: string): Promise<WebhookResponse> {
    // Teams outgoing webhook sends: { type, text, from: { name }, ... }
    let text = (body.text || '').trim();
    const fromName = userName || body.from?.name || 'Unknown';

    // Handle card action button clicks (Action.Submit sends data in body.value)
    const actionData = body.value;
    if (actionData && actionData.action) {
      return this.handleActionData(actionData, fromName);
    }

    // Strip @mention (Teams prepends "<at>BotName</at> " to the message)
    text = text.replace(/<at>.*?<\/at>\s*/gi, '').trim();

    if (!text) {
      return this.textResponse('Please type a command or **help** to see available options.');
    }

    // Normalize: lowercase, collapse extra spaces, fix common typos
    text = text.toLowerCase().replace(/\s+/g, ' ');
    text = text.replace(/\bstatis\b/g, 'status').replace(/\bstaus\b/g, 'status')
               .replace(/\budpate\b/g, 'update').replace(/\bupdat\b/g, 'update')
               .replace(/\bcomlete\b/g, 'complete').replace(/\bcomplte\b/g, 'complete')
               .replace(/\btaks\b/g, 'tasks').replace(/\btask\b/g, 'tasks')
               .replace(/\bactivites\b/g, 'activities').replace(/\bactivitys\b/g, 'activities')
               .replace(/\bworksteam\b/g, 'workstream').replace(/\bworkstrem\b/g, 'workstream');

    try {
      // Intent routing (mirrors QQIABot.handleMessage)
      if (text === 'help' || text === '?') {
        return this.handleHelp();
      } else if (text === 'dashboard' || text === 'dash' || text.startsWith('show dashboard')) {
        return this.handleDashboard();
      } else if (text === 'my tasks' || text === 'my steps' || text === 'my items') {
        return this.handleMyTasks(fromName);
      } else if (text.startsWith('status ') || text.startsWith('step ') || text.startsWith('show ')) {
        return this.handleStepQuery(text);
      } else if (text.startsWith('update ') || text.startsWith('mark ') || text.startsWith('complete ')) {
        return this.handleStatusUpdate(text, fromName);
      } else if (text === 'blockers' || text === 'blocked' || text === 'show blockers') {
        return this.handleBlockers();
      } else if (text === 'overdue' || text === 'show overdue') {
        return this.handleOverdue();
      } else if (text.startsWith('workstream ') || text.startsWith('ws ')) {
        return this.handleWorkstream(text);
      } else if (text === 'critical path' || text === 'cp') {
        return this.handleCriticalPath();
      } else if (text === 'summary' || text === 'exec summary' || text.startsWith('leadership')) {
        return this.handleSummary();
      } else if (text === 'upcoming' || text === 'coming up' || text === 'next steps' || text.startsWith('upcoming ')) {
        const daysMatch = text.match(/(\d+)\s*days?/);
        return this.handleUpcoming(daysMatch ? parseInt(daysMatch[1]) : 7);
      } else if (text.startsWith('tasks for ') || text.startsWith('owner ')) {
        return this.handleOwnerTasks(text);
      } else if (text === 'sync' || text === 'refresh') {
        return this.handleSync();
      } else if (text.startsWith('note ') || text.startsWith('add note ')) {
        return this.handleAddNote(text, fromName);
      } else if (text.startsWith('fed ') || text.startsWith('fed status ')) {
        return this.handleFedQuery(text, fromName);
      } else {
        return this.handleNaturalLanguage(text, fromName);
      }
    } catch (error: any) {
      console.error('Webhook error:', error);
      return this.textResponse(`❌ Error: ${error.message}. Please try again or type **help**.`);
    }
  }

  // ---- Response Helpers ----

  private textResponse(text: string): WebhookResponse {
    return { type: 'message', text };
  }

  private cardResponse(card: any): WebhookResponse {
    return {
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: card,
      }],
    };
  }

  /** Handle button clicks from Adaptive Card actions */
  private async handleActionData(data: any, userName: string): Promise<WebhookResponse> {
    if (data.action === 'update_status') {
      const { stepId, field, newStatus } = data;
      const step = await this.dataService.updateStepStatus(
        stepId, field, newStatus, userName, 'webhook'
      );
      if (!step) {
        return this.textResponse(`Step **${stepId}** not found.`);
      }
      return this.textResponse(
        `✅ Step **${stepId}** ${field === 'fedStatus' ? '(Fed)' : '(Corp)'} updated to **${newStatus}** by ${userName}.`
      );
    } else if (data.action === 'view_step') {
      return this.handleStepQuery(`status ${data.stepId}`);
    }
    return this.textResponse('Unknown action.');
  }

  private extractStepId(text: string): string | null {
    const match = text.match(/\b(\d+\.\w+(?:\.\d+)?)\b/);
    return match ? match[1].toUpperCase() : null;
  }

  // ---- Intent Handlers ----

  private handleHelp(): WebhookResponse {
    return this.textResponse(
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

  private async handleDashboard(track: 'Corp' | 'Fed' = 'Corp'): Promise<WebhookResponse> {
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);
    const stats = this.dependencyEngine.getWorkstreamStats(allSteps, track);
    const overdue = this.dependencyEngine.getOverdueSteps(allSteps, track);
    const blocked = allSteps.filter(s =>
      (track === 'Corp' ? s.corpStatus : s.fedStatus) === 'Blocked'
    );
    const card = buildOverallDashboardCard(allSteps, stats, overdue, blocked, track);
    return this.cardResponse(card);
  }

  private async handleMyTasks(userName: string): Promise<WebhookResponse> {
    const steps = await this.dataService.getStepsByOwner(userName);
    if (steps.length === 0) {
      return this.textResponse(`No steps found for **${userName}**. Try "tasks for [name]" with the exact name from the tracker.`);
    }
    const card = buildMyTasksCard(userName, steps);
    return this.cardResponse(card);
  }

  private async handleStepQuery(text: string): Promise<WebhookResponse> {
    const stepId = this.extractStepId(text);
    if (!stepId) {
      return this.textResponse('Please provide a step ID, e.g., **status 1.A**');
    }
    const step = await this.dataService.getStep(stepId);
    if (!step) {
      return this.textResponse(`Step **${stepId}** not found. Check the step ID and try again.`);
    }
    this.dependencyEngine.buildGraph(await this.dataService.getAllSteps());
    const blockers = this.dependencyEngine.getBlockers(stepId);
    const blockedBy = this.dependencyEngine.getBlockedBy(stepId);
    const card = buildStepDetailCard(step, blockers, blockedBy);
    return this.cardResponse(card);
  }

  private async handleStatusUpdate(text: string, userName: string): Promise<WebhookResponse> {
    const stepId = this.extractStepId(text);
    if (!stepId) {
      return this.textResponse('Please provide a step ID, e.g., **update 1.A completed**');
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

    const updated = await this.dataService.updateStepStatus(stepId, field as any, newStatus, userName, 'webhook');
    if (!updated) {
      return this.textResponse(`Step **${stepId}** not found.`);
    }

    // Sync to Excel
    let syncMsg = '';
    try {
      if (this.comSync?.isAvailable) {
        await this.comSync.syncToExcel();
      } else if (this.paSyncService?.isConfigured) {
        await this.paSyncService.syncToExcel();
      } else {
        await this.excelSync.syncToExcel();
      }
      syncMsg = `\n📊 Excel file updated.`;
    } catch {
      syncMsg = `\n⚠️ Excel sync pending — will retry on next scheduled sync.`;
    }

    // Trigger dependency notifications if completed
    let unblockMsg = '';
    if (newStatus === 'Completed') {
      const notifications = await this.notificationService.notifyPredecessorComplete(stepId, track as any);
      if (notifications.length > 0) {
        unblockMsg = `\n🔓 ${notifications.length} step(s) are now unblocked: ${notifications.map(n => n.steps[0]?.id).join(', ')}`;
      }
    }

    return this.textResponse(
      `✅ Step **${stepId}** ${track} status updated to **${newStatus}** by ${userName}.` +
      (newStatus === 'Completed' ? `\nCompleted date set to ${new Date().toISOString().split('T')[0]}.` : '') +
      syncMsg + unblockMsg
    );
  }

  private async handleBlockers(): Promise<WebhookResponse> {
    const blocked = await this.dataService.getStepsByStatus('Blocked', 'Corp');
    if (blocked.length === 0) {
      return this.textResponse('✅ No blocked steps! Everything is moving forward.');
    }
    const card = buildStepListCard('🛑 Blocked Steps', blocked, 'Corp');
    return this.cardResponse(card);
  }

  private async handleOverdue(): Promise<WebhookResponse> {
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);
    const overdue = this.dependencyEngine.getOverdueSteps(allSteps, 'Corp');
    if (overdue.length === 0) {
      return this.textResponse('✅ No overdue steps! All on track.');
    }
    const card = buildStepListCard('⚠️ Overdue Steps', overdue, 'Corp');
    return this.cardResponse(card);
  }

  private async handleWorkstream(text: string): Promise<WebhookResponse> {
    const wsName = text.replace(/^(workstream|ws)\s+/i, '').trim();
    const steps = await this.dataService.getStepsByWorkstream(wsName);
    if (steps.length === 0) {
      const allSteps = await this.dataService.getAllSteps();
      const workstreams = [...new Set(allSteps.map(s => s.workstream))];
      const match = workstreams.find(ws => ws.toLowerCase().includes(wsName.toLowerCase()));
      if (match) {
        const matchedSteps = allSteps.filter(s => s.workstream === match);
        const card = buildStepListCard(`📦 ${match}`, matchedSteps, 'Corp');
        return this.cardResponse(card);
      }
      return this.textResponse(`Workstream "${wsName}" not found. Available: ${workstreams.join(', ')}`);
    }
    const card = buildStepListCard(`📦 ${wsName}`, steps, 'Corp');
    return this.cardResponse(card);
  }

  private async handleCriticalPath(): Promise<WebhookResponse> {
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);
    const path = this.dependencyEngine.getCriticalPath('Corp');
    if (path.length === 0) {
      return this.textResponse('No critical path found (all steps may be complete or independent).');
    }
    const pathSteps: RolloverStep[] = [];
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
    return this.textResponse(msg);
  }

  private async handleSummary(): Promise<WebhookResponse> {
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
    return this.textResponse(summary);
  }

  private async handleUpcoming(days: number = 7): Promise<WebhookResponse> {
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
      return this.cardResponse(card);
    }
    // Fallback: show all pending
    const pending = allSteps.filter(s => s.corpStatus === 'Not Started' || s.corpStatus === 'In Progress')
      .sort((a, b) => {
        const da = a.corpEndDate ? new Date(a.corpEndDate).getTime() : Infinity;
        const db = b.corpEndDate ? new Date(b.corpEndDate).getTime() : Infinity;
        return da - db;
      }).slice(0, 20);
    if (pending.length > 0) {
      const card = buildStepListCard(`📅 Next Pending Activities (${pending.length} shown)`, pending, 'Corp');
      return this.cardResponse(card);
    }
    return this.textResponse('✅ All activities are completed!');
  }

  private async handleOwnerTasks(text: string): Promise<WebhookResponse> {
    const name = text.replace(/^(tasks for|owner)\s+/i, '').trim();
    const steps = await this.dataService.getStepsByOwner(name);
    if (steps.length === 0) {
      return this.textResponse(`No steps found for **${name}**.`);
    }
    const card = buildMyTasksCard(name, steps);
    return this.cardResponse(card);
  }

  private async handleSync(): Promise<WebhookResponse> {
    const result = await this.excelSync.fullSync();
    return this.textResponse(
      `✅ Sync complete!\n- **${result.fromExcel}** updates from Excel\n- **${result.toExcel}** updates pushed to Excel`
    );
  }

  private async handleAddNote(text: string, userName: string): Promise<WebhookResponse> {
    const match = text.match(/^(?:add\s+)?note\s+(\d+\.\w+(?:\.\d+)?)\s+(.+)/i);
    if (!match) {
      return this.textResponse('Usage: **note 1.A your note text here**');
    }
    const [, stepId, note] = match;
    const step = await this.dataService.getStep(stepId);
    if (!step) {
      return this.textResponse(`Step **${stepId}** not found.`);
    }
    const oldNotes = step.referenceNotes;
    step.referenceNotes = step.referenceNotes
      ? `${step.referenceNotes}\n[${new Date().toISOString().split('T')[0]} ${userName}]: ${note}`
      : `[${new Date().toISOString().split('T')[0]} ${userName}]: ${note}`;
    step.lastModified = new Date().toISOString();
    step.lastModifiedBy = userName;
    step.lastModifiedSource = 'webhook';
    await this.dataService.upsertStep(step);
    await this.dataService.logAudit(stepId, 'referenceNotes', oldNotes, step.referenceNotes, userName, 'webhook');
    return this.textResponse(`📝 Note added to step **${stepId}**.`);
  }

  private async handleFedQuery(text: string, userName: string): Promise<WebhookResponse> {
    const cleaned = text.replace(/^fed\s+(status\s+)?/i, '').trim();
    if (cleaned.startsWith('update') || cleaned.startsWith('mark') || cleaned.startsWith('complete')) {
      return this.handleStatusUpdate(`fed ${cleaned}`, userName);
    }
    return this.handleDashboard('Fed');
  }

  /** Fallback handler for natural language queries */
  private async handleNaturalLanguage(text: string, userName: string): Promise<WebhookResponse> {
    const stepId = this.extractStepId(text);
    if (stepId && text.replace(/\s/g, '').length <= stepId.length + 2) {
      return this.handleStepQuery(`status ${stepId}`);
    }
    // Handle updates BEFORE queries — "change status of 15.B to completed" has both "status of" and "completed"
    if (stepId && (text.includes('change') || text.includes(' to ') || text.includes('completed') || text.includes('done') ||
        text.includes('in progress') || text.includes('blocked') || text.includes('complete') ||
        text.includes('mark') || text.includes('set'))) {
      // Extract the target status
      let newStatus = 'Completed';
      if (text.includes('in progress') || text.includes('in-progress') || text.includes('started')) newStatus = 'In Progress';
      else if (text.includes('blocked') || text.includes('block')) newStatus = 'Blocked';
      else if (text.includes('not started') || text.includes('reset')) newStatus = 'Not Started';
      const field = text.includes('fed') ? 'fedStatus' : 'corpStatus';
      const step = await this.dataService.updateStepStatus(stepId, field as any, newStatus, userName, 'webhook');
      if (!step) return this.textResponse(`Step **${stepId}** not found.`);
      return this.textResponse(`✅ Step **${stepId}** ${field === 'fedStatus' ? '(Fed)' : '(Corp)'} updated to **${newStatus}** by ${userName}.`);
    }
    // Handle natural language status queries like "what is the status of step 1.A"
    if (stepId && (text.includes('status of') || text.includes('what is') || text.includes('what\'s') ||
        text.includes('whats') || text.includes('check') || text.includes('show me') ||
        text.includes('tell me') || text.includes('details') || text.includes('info'))) {
      return this.handleStepQuery(`status ${stepId}`);
    }

    const isPersonal = /\b(i |i'[a-z]|my |me |mine)\b/.test(text) || text.startsWith('do i ') || text.startsWith('what do i');

    // Detect time-window keywords
    const timeMatch = text.match(/(?:this|next)\s+(\d+\s+)?(week|month|sprint)|(?:next|coming|upcoming)\s+(\d+)\s+days?|(?:in the next|within)\s+(\d+)\s+days?/i);
    let lookAheadDays = 7;
    if (timeMatch) {
      if (timeMatch[2] === 'month') lookAheadDays = 30;
      else if (timeMatch[3]) lookAheadDays = parseInt(timeMatch[3]);
      else if (timeMatch[4]) lookAheadDays = parseInt(timeMatch[4]);
    }

    // Detect mentioned person
    let mentionedPerson: string | null = null;
    const possessiveMatch = text.match(/(?:what(?:'s| is| are)\s+)?(\w[\w\s]*?)'s\s+(?:tasks?|activities|steps|items|work|upcoming)/i);
    const forMatch = text.match(/(?:tasks?|activities|steps|work|upcoming)\s+(?:for|of|assigned to)\s+(\w[\w\s]*?)(?:\s+(?:this|next|that|due|need)|[?.]|$)/i);
    const doesMatch = text.match(/what\s+(?:does|should|will|can)\s+(\w[\w\s]*?)\s+(?:need|have|do|start|complete|own)/i);
    if (possessiveMatch) mentionedPerson = possessiveMatch[1].trim();
    else if (forMatch) mentionedPerson = forMatch[1].trim();
    else if (doesMatch) mentionedPerson = doesMatch[1].trim();

    if (mentionedPerson && (text.includes('task') || text.includes('activities') || text.includes('step') ||
        text.includes('need') || text.includes('start') || text.includes('do') ||
        text.includes('due') || text.includes('upcoming') || text.includes('work') ||
        text.includes('this week') || text.includes('this month'))) {
      return this.handleOwnerTasks(`tasks for ${mentionedPerson}`);
    } else if (isPersonal && (text.includes('activities') || text.includes('task') || text.includes('upcoming') ||
        text.includes('need to') || text.includes('start') || text.includes('do') ||
        text.includes('due') || text.includes('pending') || text.includes('this week') ||
        text.includes('next week') || text.includes('this month') || text.includes('assigned') ||
        text.includes('own') || text.includes('working') || text.includes('responsible'))) {
      return this.handleOwnerTasks(`tasks for ${userName}`);
    } else if (text.includes('upcoming') || text.includes('coming up') || text.includes('what\'s next') ||
        text.includes('whats next') || text.includes('need to be completed') || text.includes('needs to be') ||
        text.includes('activities') || text.includes('pending') ||
        text.includes('scheduled') || text.includes('planned') || text.includes('remaining') ||
        (text.includes('what') && (text.includes('due') || text.includes('deadline') || text.includes('next') || text.includes('left') || text.includes('remain'))) ||
        text.includes('due soon') || text.includes('due this') || text.includes('due next') ||
        text.includes('this week') || text.includes('next week') || text.includes('this month')) {
      return this.handleUpcoming(lookAheadDays);
    } else if (text.includes('how many') || text.includes('count') || text.includes('total') ||
               text.includes('progress') || text.includes('overview') || text.includes('where are we') ||
               text.includes('how are we') || text.includes('status update') || text.includes('update me')) {
      return this.handleSummary();
    } else if (text.includes('blocker') || text.includes('blocking') || text.includes('stuck') || text.includes('at risk')) {
      return this.handleBlockers();
    } else if (text.includes('overdue') || text.includes('late') || text.includes('behind') || text.includes('missed')) {
      return this.handleOverdue();
    } else if (text.includes('summary') || text.includes('leadership') || text.includes('exec') || text.includes('report')) {
      return this.handleSummary();
    } else if (stepId) {
      // If we found a step ID anywhere in the text, show its status
      return this.handleStepQuery(`status ${stepId}`);
    } else {
      return this.textResponse(
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
