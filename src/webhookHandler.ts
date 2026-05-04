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

    // Synonym expansion (mirrors QQIABot)
    text = text.replace(/\bfinished\b/g, 'completed').replace(/\bfinish\b/g, 'complete')
               .replace(/\bstuck\b/g, 'blocked').replace(/\bhold\b/g, 'blocked').replace(/\bon hold\b/g, 'blocked')
               .replace(/\bpending\b/g, 'not started').replace(/\bnot begun\b/g, 'not started').replace(/\bwaiting\b/g, 'not started')
               .replace(/\bin-progress\b/g, 'in progress').replace(/\bstarted\b/g, 'in progress')
               .replace(/\bunderway\b/g, 'in progress').replace(/\bongoing\b/g, 'in progress')
               .replace(/\bassigned to\b/g, 'owner').replace(/\bowned by\b/g, 'owner')
               .replace(/\bdelayed\b/g, 'overdue').replace(/\bpast due\b/g, 'overdue').replace(/\bslipped\b/g, 'overdue')
               .replace(/\bdeps\b/g, 'dependencies').replace(/\bprereqs?\b/g, 'dependencies').replace(/\bprerequisites?\b/g, 'dependencies')
               .replace(/\btimeline\b/g, 'upcoming').replace(/\bschedule\b/g, 'upcoming').replace(/\bcalendar\b/g, 'upcoming')
               .replace(/\bprogress report\b/g, 'summary').replace(/\bstatus report\b/g, 'summary');

    // Fuzzy command matching (first word)
    const firstWord = text.split(' ')[0];
    const fuzzyTarget = this.fuzzyMatchCommand(firstWord);
    if (fuzzyTarget && fuzzyTarget !== firstWord) {
      text = fuzzyTarget + text.slice(firstWord.length);
    }

    try {
      // Greetings & chitchat
      if (/^(hi|hello|hey|howdy|good morning|good afternoon|good evening|yo)\b/.test(text)) {
        return this.textResponse(`👋 Hi ${fromName}! Type **help** to see what I can do, or just ask me about FY27 rollover.`);
      }
      if (/^(thanks|thank you|thx|ty|cheers|appreciate it)\b/.test(text)) {
        return this.textResponse(`You're welcome! Let me know if you need anything else. 😊`);
      }
      if (/^(bye|goodbye|see you|later|cya|good night)\b/.test(text)) {
        return this.textResponse(`See you later, ${fromName}! 👋`);
      }

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

  /** Extract ALL step IDs from text for batch operations (e.g., "1.A, 1.B, and 1.C") */
  private extractAllStepIds(text: string): string[] {
    const ids = new Set<string>();
    const dotPattern = /\b(\d+)\s*\.\s*(\w+(?:\s*\.\s*\d+)?)\b/g;
    let m: RegExpExecArray | null;
    while ((m = dotPattern.exec(text)) !== null) {
      ids.add((m[1] + '.' + m[2]).replace(/\s/g, '').toUpperCase());
    }
    const shortPattern = /\b(\d+)([A-Za-z])\b/g;
    while ((m = shortPattern.exec(text)) !== null) {
      const id = m[1] + '.' + m[2].toUpperCase();
      if (!ids.has(id)) ids.add(id);
    }
    return Array.from(ids);
  }

  // ---- Fuzzy matching ----

  private static KNOWN_COMMANDS = [
    'help', 'dashboard', 'status', 'update', 'mark', 'complete', 'blockers',
    'blocked', 'overdue', 'workstream', 'summary', 'upcoming', 'sync', 'refresh',
    'note', 'tasks', 'dependencies', 'critical', 'show', 'step', 'owner', 'fed',
  ];

  private static levenshtein(a: string, b: string): number {
    const m = a.length, n = b.length;
    const dp = Array.from({ length: m + 1 }, (_, i) =>
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

  private fuzzyMatchCommand(word: string): string | null {
    if (word.length < 3) return null;
    let best: string | null = null;
    let bestDist = 3;
    for (const cmd of WebhookHandler.KNOWN_COMMANDS) {
      const dist = WebhookHandler.levenshtein(word, cmd);
      if (dist === 0) return word; // Exact match
      if (dist < bestDist) { bestDist = dist; best = cmd; }
    }
    return best;
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
    const allIds = this.extractAllStepIds(text);
    if (allIds.length === 0) {
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

    // Single step update
    if (allIds.length === 1) {
      const stepId = allIds[0];
      const updated = await this.dataService.updateStepStatus(stepId, field as any, newStatus, userName, 'webhook');
      if (!updated) {
        return this.textResponse(`Step **${stepId}** not found.`);
      }

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

    // Batch update — multiple step IDs
    const succeeded: string[] = [];
    const failed: string[] = [];
    for (const sid of allIds) {
      const ok = await this.dataService.updateStepStatus(sid, field as any, newStatus, userName, 'webhook');
      if (ok) { succeeded.push(sid); } else { failed.push(sid); }
    }

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

    let msg = `✅ **Batch update**: ${succeeded.length} step(s) set to **${newStatus}** (${track}) by ${userName}.`;
    if (succeeded.length > 0) msg += `\n✔️ Updated: ${succeeded.join(', ')}`;
    if (failed.length > 0) msg += `\n⚠️ Not found: ${failed.join(', ')}`;
    if (newStatus === 'Completed') msg += `\nCompleted date set to ${new Date().toISOString().split('T')[0]}.`;
    msg += syncMsg;

    return this.textResponse(msg);
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

    // For "this week", include items from start of the week (Monday) through end of week
    const startOfWeek = new Date(now);
    const dayOfWeek = startOfWeek.getDay();
    const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    startOfWeek.setDate(startOfWeek.getDate() + mondayOffset);
    startOfWeek.setHours(0, 0, 0, 0);

    const rangeStart = days <= 7 ? startOfWeek : now;

    const active = allSteps.filter(s => {
      if (s.corpStatus === 'Completed' || s.corpStatus === 'N/A') return false;
      const endDate = s.corpEndDate ? new Date(s.corpEndDate) : null;
      return endDate && endDate >= rangeStart && endDate <= cutoff;
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
    // Fallback: show next 10 pending items only (not 20)
    const pending = allSteps.filter(s => s.corpStatus === 'Not Started' || s.corpStatus === 'In Progress')
      .sort((a, b) => {
        const da = a.corpEndDate ? new Date(a.corpEndDate).getTime() : Infinity;
        const db = b.corpEndDate ? new Date(b.corpEndDate).getTime() : Infinity;
        return da - db;
      }).slice(0, 10);
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

  /** Show a person's tasks filtered by upcoming date range */
  private async handleOwnerUpcoming(name: string, days: number): Promise<WebhookResponse> {
    const steps = await this.dataService.getStepsByOwner(name);
    if (steps.length === 0) {
      return this.textResponse(`No steps found for **${name}**.`);
    }
    const now = new Date();
    const cutoff = new Date(now.getTime() + days * 86400000);
    // For <= 7 days, use start of week (Monday)
    const startOfWeek = new Date(now);
    const dayOfWeek = startOfWeek.getDay();
    const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    startOfWeek.setDate(startOfWeek.getDate() + mondayOffset);
    startOfWeek.setHours(0, 0, 0, 0);
    const rangeStart = days <= 7 ? startOfWeek : now;

    const filtered = steps.filter(s => {
      if (s.corpStatus === 'Completed' || s.corpStatus === 'N/A') return false;
      const endDate = s.corpEndDate ? new Date(s.corpEndDate) : null;
      const startDate = s.corpStartDate ? new Date(s.corpStartDate) : null;
      // Include if end date or start date falls within range
      return (endDate && endDate >= rangeStart && endDate <= cutoff) ||
             (startDate && startDate >= rangeStart && startDate <= cutoff);
    }).sort((a, b) => {
      const da = a.corpEndDate ? new Date(a.corpEndDate).getTime() : Infinity;
      const db = b.corpEndDate ? new Date(b.corpEndDate).getTime() : Infinity;
      return da - db;
    });

    if (filtered.length === 0) {
      const label = days <= 7 ? 'this week' : `next ${days} days`;
      // Show the next 5 soonest upcoming items instead of dumping all pending
      const pending = steps.filter(s => s.corpStatus !== 'Completed' && s.corpStatus !== 'N/A')
        .sort((a, b) => {
          const da = a.corpEndDate ? new Date(a.corpEndDate).getTime() : Infinity;
          const db = b.corpEndDate ? new Date(b.corpEndDate).getTime() : Infinity;
          return da - db;
        });
      if (pending.length > 0) {
        const nextFew = pending.slice(0, 5);
        const nextDate = nextFew[0].corpEndDate ? new Date(nextFew[0].corpEndDate).toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) : 'TBD';
        const card = buildStepListCard(
          `📅 ${name} — No items due ${label}. Next soonest (${pending.length} total pending, earliest: ${nextDate})`,
          nextFew, 'Corp');
        return this.cardResponse(card);
      }
      return this.textResponse(`No upcoming steps for **${name}** ${label}. All ${steps.length} steps are completed! ✅`);
    }
    const label = days <= 7 ? 'This Week' : `Next ${days} Days`;
    const card = buildStepListCard(`📅 ${name} — ${label} (${filtered.length} steps)`, filtered, 'Corp');
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

  /** Match a candidate name against workstream or engineering team, return steps or null */
  private async handleTeamQuery(candidate: string): Promise<WebhookResponse | null> {
    const allSteps = await this.dataService.getAllSteps();
    // Clean up candidate: strip "team", "teams", "group" suffix
    const lower = candidate.toLowerCase().replace(/\s+teams?$|\s+groups?$/i, '').trim();
    console.log(`[TeamQuery] candidate="${candidate}" cleaned="${lower}"`);

    // Match against workstream names (fuzzy contains)
    const workstreams = [...new Set(allSteps.map(s => s.workstream))];
    const wsMatch = workstreams.find(ws =>
      ws.toLowerCase().includes(lower) || lower.includes(ws.toLowerCase())
    );
    if (wsMatch) {
      const steps = allSteps.filter(s => s.workstream === wsMatch);
      console.log(`[TeamQuery] Matched workstream "${wsMatch}" with ${steps.length} steps`);
      const card = buildStepListCard(`📦 ${wsMatch} (${steps.length} steps)`, steps, 'Corp');
      return this.cardResponse(card);
    }

    // Match against engineeringDependent values (e.g., "orchestration" → "Y - Orchestration")
    const engTeams = [...new Set(allSteps.map(s => s.engineeringDependent).filter(e => e && e !== 'N' && e !== 'Y'))];
    const engMatch = engTeams.find(et =>
      et.toLowerCase().includes(lower) || lower.includes(et.replace(/^Y\s*-?\s*/i, '').toLowerCase())
    );
    if (engMatch) {
      const steps = allSteps.filter(s => s.engineeringDependent === engMatch);
      const label = engMatch.replace(/^Y\s*-?\s*/i, '').trim();
      const card = buildStepListCard(`🔧 ${label} Team (${steps.length} steps)`, steps, 'Corp');
      return this.cardResponse(card);
    }

    return null; // No team match found
  }

  /** Fallback handler for natural language queries */
  private async handleNaturalLanguage(text: string, userName: string): Promise<WebhookResponse> {
    console.log(`[NL] Input: "${text}"`);

    // Early team/group detection: if text mentions "team" or "group" along with activity words,
    // extract the team name and try team query first
    if ((text.includes('team') || text.includes('group')) &&
        (text.includes('action') || text.includes('task') || text.includes('activit') ||
         text.includes('item') || text.includes('step') || text.includes('work'))) {
      // Extract everything that looks like a team name
      const teamExtract = text
        .replace(/\b(?:what|are|is|the|show|me|list|all|get|for|this|next|week|month|days?|\d+)\b/g, '')
        .replace(/\b(?:action items|tasks?|activities|items|steps|work|upcoming)\b/g, '')
        .replace(/\s+/g, ' ').trim();
      console.log(`[NL] Early team extract: "${teamExtract}"`);
      if (teamExtract) {
        const teamResult = await this.handleTeamQuery(teamExtract);
        if (teamResult) return teamResult;
      }
    }

    const stepId = this.extractStepId(text);
    if (stepId && text.replace(/\s/g, '').length <= stepId.length + 2) {
      return this.handleStepQuery(`status ${stepId}`);
    }
    // Handle updates BEFORE queries — "change status of 15.B to completed" has both "status of" and "completed"
    const allStepIds = this.extractAllStepIds(text);
    const hasUpdateIntent = text.includes('change') || text.includes(' to ') || text.includes('completed') || text.includes('done') ||
        text.includes('in progress') || text.includes('blocked') || text.includes('complete') ||
        text.includes('mark') || text.includes('set');
    if (allStepIds.length > 0 && hasUpdateIntent) {
      let newStatus = 'Completed';
      if (text.includes('in progress') || text.includes('in-progress') || text.includes('started')) newStatus = 'In Progress';
      else if (text.includes('blocked') || text.includes('block')) newStatus = 'Blocked';
      else if (text.includes('not started') || text.includes('reset')) newStatus = 'Not Started';
      const field = text.includes('fed') ? 'fedStatus' : 'corpStatus';
      const track = field === 'fedStatus' ? 'Fed' : 'Corp';

      if (allStepIds.length === 1) {
        const sid = allStepIds[0];
        const step = await this.dataService.updateStepStatus(sid, field as any, newStatus, userName, 'webhook');
        if (!step) return this.textResponse(`Step **${sid}** not found.`);
        let msg = `✅ Step **${sid}** (${track}) updated to **${newStatus}** by ${userName}.`;
        if (newStatus === 'Completed') msg += `\nCompleted date set to ${new Date().toISOString().split('T')[0]}.`;
        return this.textResponse(msg);
      }

      // Batch update
      const succeeded: string[] = [];
      const failed: string[] = [];
      for (const sid of allStepIds) {
        const ok = await this.dataService.updateStepStatus(sid, field as any, newStatus, userName, 'webhook');
        if (ok) { succeeded.push(sid); } else { failed.push(sid); }
      }
      let msg = `✅ **Batch update**: ${succeeded.length} step(s) set to **${newStatus}** (${track}) by ${userName}.`;
      if (succeeded.length > 0) msg += `\n✔️ Updated: ${succeeded.join(', ')}`;
      if (failed.length > 0) msg += `\n⚠️ Not found: ${failed.join(', ')}`;
      if (newStatus === 'Completed') msg += `\nCompleted date set to ${new Date().toISOString().split('T')[0]}.`;
      return this.textResponse(msg);
    }
    // Handle natural language status queries like "what is the status of step 1.A"
    if (stepId && (text.includes('status of') || text.includes('what is') || text.includes('what\'s') ||
        text.includes('whats') || text.includes('check') || text.includes('show me') ||
        text.includes('tell me') || text.includes('details') || text.includes('info'))) {
      return this.handleStepQuery(`status ${stepId}`);
    }

    const isPersonal = /\b(i |i'[a-z]|my |me |mine)\b/.test(text) || text.startsWith('do i ') || text.startsWith('what do i');

    // ---- Status + time-range queries (e.g., "what was completed last month", "what's in progress this week") ----
    const statusFilterMatch = text.match(/\b(completed|in progress|not started|blocked|overdue)\b/);
    const lastTimeMatch = text.match(/\b(last|past|previous)\s*(week|month|quarter|\d+\s*days?)\b/);
    const thisTimeMatch = text.match(/\b(this)\s*(week|month)\b/);
    if (statusFilterMatch) {
      const targetStatus = statusFilterMatch[1] === 'completed' ? 'Completed'
        : statusFilterMatch[1] === 'in progress' ? 'In Progress'
        : statusFilterMatch[1] === 'not started' ? 'Not Started'
        : statusFilterMatch[1] === 'blocked' ? 'Blocked'
        : null;

      const allSteps = await this.dataService.getAllSteps();
      let filtered = targetStatus
        ? allSteps.filter(s => s.corpStatus === targetStatus || s.fedStatus === targetStatus)
        : allSteps;

      // Apply time range filter
      const now = new Date();
      let rangeStart: Date | null = null;
      let rangeEnd: Date | null = null;
      let rangeLabel = '';

      if (lastTimeMatch) {
        const unit = lastTimeMatch[2];
        if (unit === 'week') {
          rangeStart = new Date(now.getTime() - 7 * 86400000);
          rangeEnd = now;
          rangeLabel = 'Last Week';
        } else if (unit === 'month') {
          rangeStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
          rangeEnd = new Date(now.getFullYear(), now.getMonth(), 1);
          rangeLabel = 'Last Month';
        } else if (unit === 'quarter') {
          rangeStart = new Date(now.getTime() - 90 * 86400000);
          rangeEnd = now;
          rangeLabel = 'Last Quarter';
        } else {
          const daysNum = parseInt(unit);
          if (daysNum) {
            rangeStart = new Date(now.getTime() - daysNum * 86400000);
            rangeEnd = now;
            rangeLabel = `Last ${daysNum} Days`;
          }
        }
      } else if (thisTimeMatch) {
        const unit = thisTimeMatch[2];
        if (unit === 'week') {
          const dayOfWeek = now.getDay();
          rangeStart = new Date(now.getTime() - dayOfWeek * 86400000);
          rangeEnd = new Date(rangeStart.getTime() + 7 * 86400000);
          rangeLabel = 'This Week';
        } else if (unit === 'month') {
          rangeStart = new Date(now.getFullYear(), now.getMonth(), 1);
          rangeEnd = new Date(now.getFullYear(), now.getMonth() + 1, 1);
          rangeLabel = 'This Month';
        }
      }

      if (rangeStart && rangeEnd) {
        const rStart = rangeStart.toISOString();
        const rEnd = rangeEnd.toISOString();
        if (targetStatus === 'Completed') {
          // Filter by completed date
          filtered = filtered.filter(s => {
            const d = s.corpCompletedDate || s.lastModified;
            return d && d >= rStart && d < rEnd;
          });
        } else {
          // Filter by due date
          filtered = filtered.filter(s => {
            const d = s.corpEndDate;
            return d && d >= rStart && d < rEnd;
          });
        }
      }

      const title = `📋 ${targetStatus || 'All'} Steps${rangeLabel ? ' — ' + rangeLabel : ''} (${filtered.length})`;
      if (filtered.length === 0) {
        return this.textResponse(`No ${(targetStatus || '').toLowerCase()} steps found${rangeLabel ? ' for ' + rangeLabel.toLowerCase() : ''}.`);
      }
      const card = buildStepListCard(title, filtered.slice(0, 25), 'Corp');
      return this.cardResponse(card);
    }

    // Detect time-window keywords
    const timeMatch = text.match(/(?:this|next)\s+(\d+\s+)?(week|month|sprint)|(?:next|coming|upcoming)\s+(\d+)\s+days?|(?:in the next|within)\s+(\d+)\s+days?/i);
    let lookAheadDays = 7;
    if (timeMatch) {
      if (timeMatch[2] === 'month') lookAheadDays = 30;
      else if (timeMatch[3]) lookAheadDays = parseInt(timeMatch[3]);
      else if (timeMatch[4]) lookAheadDays = parseInt(timeMatch[4]);
    }

    // Strip time phrases from text for clean person/team extraction
    const timeStripped = text
      .replace(/\b(?:for\s+)?(?:the\s+)?(?:this|next)\s+(?:\d+\s+)?(?:weeks?|months?|sprint|days?)\b/g, '')
      .replace(/\b(?:for\s+)?(?:the\s+)?(?:next|coming|upcoming)\s+\d+\s+days?\b/g, '')
      .replace(/\b(?:for\s+)?(?:in the next|within)\s+\d+\s+days?\b/g, '')
      .replace(/\s+for\s*$/, '') // trailing "for" after time removal
      .replace(/\s+/g, ' ')
      .trim();

    // Detect person/team name — try multiple patterns on time-stripped text

    let detectedName: string | null = null;

    // Pattern 1: name BEFORE activity word (possessive, with or without apostrophe)
    // "amit tiwaris activities" → "amit tiwari", "pragyas tasks" → "pragya"
    const beforeActivity = timeStripped.match(
      /(?:what\s+(?:are|is)\s+)?(.+?)(?:'s|s)\s+(?:action items|tasks?|activities|items|steps|work|upcoming)\b/i
    );
    if (beforeActivity) {
      detectedName = beforeActivity[1].trim();
    }

    // Pattern 2: "activities/tasks for [name]" (on time-stripped text so "next week" is gone)
    if (!detectedName) {
      const afterFor = timeStripped.match(
        /(?:action items|tasks?|activities|items|steps|work|upcoming)\s+(?:for|of|assigned to)\s+(?:the\s+)?(.+?)(?:\s+team|\s+group)?(?:[?.]|$)/i
      );
      if (afterFor) {
        detectedName = afterFor[1].replace(/\s+team$|\s+group$/i, '').trim();
      }
    }

    // Pattern 3: "what does [name] need/have/do"
    if (!detectedName) {
      const doesMatch = timeStripped.match(/what\s+(?:does|should|will|can)\s+(\w[\w\s]*?)\s+(?:need|have|do|start|complete|own)/i);
      if (doesMatch) detectedName = doesMatch[1].trim();
    }

    // Pattern 4: "what are [name] action items/tasks/activities" (no possessive, no "for")
    // e.g. "what are MSC team action items", "what are orchestration activities"
    if (!detectedName) {
      const plainName = timeStripped.match(
        /(?:what\s+(?:are|is)\s+|show\s+(?:me\s+)?)(.+?)\s+(?:action items|tasks?|activities|items|steps)\s*[?.]?$/i
      );
      if (plainName) {
        detectedName = plainName[1].replace(/\s+team$|\s+group$/i, '').trim();
      }
    }

    // Pattern 5: "[name] activities/tasks" at start (e.g. "MSC team activities", "orchestration tasks")
    if (!detectedName) {
      const startName = timeStripped.match(
        /^(.+?)\s+(?:action items|tasks?|activities|items|steps)\s*[?.]?$/i
      );
      if (startName) {
        detectedName = startName[1].replace(/\s+team$|\s+group$/i, '').trim();
      }
    }

    // Strip "team"/"group" suffix from detected name for cleaner matching
    if (detectedName) {
      detectedName = detectedName.replace(/\s+team$|\s+group$/i, '').trim();
    }

    if (detectedName && (text.includes('task') || text.includes('activities') || text.includes('step') ||
        text.includes('need') || text.includes('start') || text.includes('do') ||
        text.includes('due') || text.includes('upcoming') || text.includes('work') ||
        text.includes('action item') || text.includes('items') ||
        text.includes('this week') || text.includes('this month') || text.includes('next'))) {
      console.log(`[NL] Detected name: "${detectedName}", timeMatch: ${!!timeMatch}, days: ${lookAheadDays}`);
      // Check if name is a team/workstream first
      const teamResult = await this.handleTeamQuery(detectedName);
      if (teamResult) { console.log(`[NL] Routed to team query`); return teamResult; }
      // Otherwise treat as person — filter by time if time was mentioned
      if (timeMatch) {
        console.log(`[NL] Routed to owner upcoming: ${detectedName}, ${lookAheadDays} days`);
        return this.handleOwnerUpcoming(detectedName, lookAheadDays);
      }
      return this.handleOwnerTasks(`tasks for ${detectedName}`);
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
      return this.buildClarifyingResponse(text, userName);
    }
  }

  /** Build a clarifying response based on keywords in the user's input */
  private buildClarifyingResponse(text: string, userName: string): WebhookResponse {
    const suggestions: string[] = [];

    // Detect what the user might be asking about
    const hasStepRef = /\b\d+\b/.test(text);
    const hasPersonRef = /\b(who|person|owner|assigned|team|member)\b/.test(text);
    const hasStatusRef = /\b(status|progress|done|complete|finish|start|block)\b/.test(text);
    const hasTimeRef = /\b(when|date|due|deadline|week|month|today|tomorrow|time|schedule)\b/.test(text);
    const hasCountRef = /\b(how many|count|number|total|percentage|percent)\b/.test(text);
    const hasListRef = /\b(list|show|all|every|which|what)\b/.test(text);
    const hasUpdateRef = /\b(update|change|set|move|mark|edit|modify)\b/.test(text);

    if (hasStepRef) {
      suggestions.push(`📌 To check a step, try: **status 1.A**`);
      suggestions.push(`📌 To update a step, try: **update 1.A completed**`);
    }
    if (hasPersonRef) {
      suggestions.push(`👤 To see someone's tasks, try: **tasks for [name]**`);
      suggestions.push(`👤 To see your tasks, try: **my tasks**`);
    }
    if (hasStatusRef) {
      suggestions.push(`📊 To see overall progress, try: **dashboard**`);
      suggestions.push(`🚫 To see blocked steps, try: **blockers**`);
    }
    if (hasTimeRef) {
      suggestions.push(`📅 To see what's coming up, try: **upcoming** or **due this week**`);
      suggestions.push(`⏰ To see overdue items, try: **overdue**`);
    }
    if (hasCountRef || hasListRef) {
      suggestions.push(`📈 For a summary, try: **summary**`);
      suggestions.push(`📋 To see all steps in a workstream, try: **workstream [name]**`);
    }
    if (hasUpdateRef) {
      suggestions.push(`✏️ To update status, try: **update [step ID] completed/in progress/blocked**`);
      suggestions.push(`✏️ To batch update, try: **mark 1.A and 1.B as completed**`);
    }

    // If no keywords matched, provide general guidance
    if (suggestions.length === 0) {
      suggestions.push(`📊 **dashboard** — See overall progress`);
      suggestions.push(`📅 **upcoming** — Activities due this week`);
      suggestions.push(`👤 **my tasks** — Your assigned steps`);
      suggestions.push(`📈 **summary** — Leadership summary`);
    }

    // Limit to top 4 most relevant suggestions
    const topSuggestions = suggestions.slice(0, 4);

    return this.textResponse(
      `I'm not sure what you're looking for. Could you clarify?\n\n` +
      `Based on your message, you might want:\n\n` +
      topSuggestions.join('\n') +
      `\n\nOr type **help** for the full command list.`
    );
  }
}
