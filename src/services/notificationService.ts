import { config } from '../config/appConfig';
import { RolloverStep, UserProfile } from '../models/types';
import { DataService } from './dataService';
import { DependencyEngine, WorkstreamStats } from './dependencyEngine';

export interface Notification {
  type: 'deadline_approaching' | 'overdue' | 'predecessor_complete' | 'blocker' | 'weekly_digest' | 'escalation';
  recipientEmail: string;
  recipientTeamsId: string;
  title: string;
  message: string;
  steps: RolloverStep[];
  priority: 'high' | 'medium' | 'low';
}

/**
 * Proactive notification engine for the QQIA Agent.
 * Generates notifications for deadlines, blockers, overdue items, and weekly digests.
 */
export class NotificationService {
  private dataService: DataService;
  private dependencyEngine: DependencyEngine;

  constructor(dataService: DataService, dependencyEngine: DependencyEngine) {
    this.dataService = dataService;
    this.dependencyEngine = dependencyEngine;
  }

  /** Check all steps and generate pending notifications */
  async generateNotifications(): Promise<Notification[]> {
    const notifications: Notification[] = [];
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);

    // Deadline approaching (3 days and 1 day)
    for (const days of config.notifications.deadlineWarningDays) {
      const upcoming = this.dependencyEngine.getUpcomingDeadlines(allSteps, days, 'Corp');
      for (const step of upcoming) {
        const owners = this.getStepOwners(step);
        for (const owner of owners) {
          notifications.push({
            type: 'deadline_approaching',
            recipientEmail: owner,
            recipientTeamsId: '',
            title: `⏰ Deadline in ${days} day(s): Step ${step.id}`,
            message: `**${step.description}** (${step.workstream}) is due in ${days} day(s).\nCorp End Date: ${step.corpEndDate}\nCurrent Status: ${step.corpStatus}`,
            steps: [step],
            priority: days === 1 ? 'high' : 'medium',
          });
        }
      }
    }

    // Overdue steps
    const overdue = this.dependencyEngine.getOverdueSteps(allSteps, 'Corp');
    for (const step of overdue) {
      const daysOverdue = this.daysSince(step.corpEndDate!);
      const owners = this.getStepOwners(step);
      for (const owner of owners) {
        notifications.push({
          type: 'overdue',
          recipientEmail: owner,
          recipientTeamsId: '',
          title: `🚨 OVERDUE (${daysOverdue}d): Step ${step.id}`,
          message: `**${step.description}** is ${daysOverdue} day(s) overdue.\nWas due: ${step.corpEndDate}\nStatus: ${step.corpStatus}`,
          steps: [step],
          priority: daysOverdue > config.notifications.overdueEscalationDays ? 'high' : 'medium',
        });
      }

      // Escalate to PMs if overdue > 3 days
      if (daysOverdue > config.notifications.overdueEscalationDays) {
        const pms = await this.dataService.getUsersByRole('pm');
        for (const pm of pms) {
          notifications.push({
            type: 'escalation',
            recipientEmail: pm.email,
            recipientTeamsId: pm.teamsUserId,
            title: `🔴 ESCALATION: Step ${step.id} overdue ${daysOverdue}d`,
            message: `**${step.description}** owned by ${step.wwicPoc || step.engineeringDri} is ${daysOverdue} days overdue and needs attention.`,
            steps: [step],
            priority: 'high',
          });
        }
      }
    }

    // Blocker notifications
    const blockedSteps = allSteps.filter(s => s.corpStatus === 'Blocked');
    for (const step of blockedSteps) {
      const pms = await this.dataService.getUsersByRole('pm');
      for (const pm of pms) {
        notifications.push({
          type: 'blocker',
          recipientEmail: pm.email,
          recipientTeamsId: pm.teamsUserId,
          title: `🛑 BLOCKED: Step ${step.id}`,
          message: `**${step.description}** is blocked.\nBlockers: ${this.dependencyEngine.getBlockers(step.id).join(', ') || 'See notes'}\nNotes: ${step.referenceNotes}`,
          steps: [step],
          priority: 'high',
        });
      }
    }

    return notifications;
  }

  /** Generate weekly digest for all stakeholders */
  async generateWeeklyDigest(): Promise<Notification[]> {
    const notifications: Notification[] = [];
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);

    const corpStats = this.dependencyEngine.getWorkstreamStats(allSteps, 'Corp');
    const fedStats = this.dependencyEngine.getWorkstreamStats(allSteps, 'Fed');
    const overdue = this.dependencyEngine.getOverdueSteps(allSteps, 'Corp');
    const criticalPath = this.dependencyEngine.getCriticalPath('Corp');
    const blocked = allSteps.filter(s => s.corpStatus === 'Blocked');

    const totalSteps = allSteps.length;
    const completed = allSteps.filter(s => s.corpStatus === 'Completed').length;
    const pct = Math.round((completed / totalSteps) * 100);

    let digest = `# 📊 QQIA Weekly Rollover Digest\n\n`;
    digest += `## Overall Progress: ${completed}/${totalSteps} (${pct}%)\n\n`;

    digest += `### Corp Workstream Summary\n`;
    for (const [ws, stats] of corpStats) {
      const wsPct = Math.round((stats.completed / stats.total) * 100);
      digest += `- **${ws}**: ${stats.completed}/${stats.total} (${wsPct}%) | `;
      digest += `🟢${stats.completed} ⏳${stats.inProgress} 🔴${stats.blocked} ⬜${stats.notStarted}\n`;
    }

    if (overdue.length > 0) {
      digest += `\n### ⚠️ Overdue Steps (${overdue.length})\n`;
      for (const step of overdue.slice(0, 10)) {
        digest += `- **${step.id}** ${step.description} (Due: ${step.corpEndDate}, Owner: ${step.wwicPoc || step.engineeringDri})\n`;
      }
    }

    if (blocked.length > 0) {
      digest += `\n### 🛑 Blocked Steps (${blocked.length})\n`;
      for (const step of blocked.slice(0, 10)) {
        digest += `- **${step.id}** ${step.description} (Owner: ${step.wwicPoc || step.engineeringDri})\n`;
      }
    }

    if (criticalPath.length > 0) {
      digest += `\n### 🔗 Critical Path\n${criticalPath.join(' → ')}\n`;
    }

    // Send to leadership and PMs
    const recipients = [
      ...await this.dataService.getUsersByRole('leadership'),
      ...await this.dataService.getUsersByRole('pm'),
    ];

    for (const user of recipients) {
      notifications.push({
        type: 'weekly_digest',
        recipientEmail: user.email,
        recipientTeamsId: user.teamsUserId,
        title: '📊 QQIA Weekly Rollover Digest',
        message: digest,
        steps: [],
        priority: 'low',
      });
    }

    return notifications;
  }

  /** Generate predecessor completion notifications */
  async notifyPredecessorComplete(completedStepId: string, track: 'Corp' | 'Fed' = 'Corp'): Promise<Notification[]> {
    const notifications: Notification[] = [];
    const allSteps = await this.dataService.getAllSteps();
    this.dependencyEngine.buildGraph(allSteps);

    const unblocked = this.dependencyEngine.getNewlyUnblocked(completedStepId, track);
    for (const stepId of unblocked) {
      const step = await this.dataService.getStep(stepId);
      if (!step) continue;

      const owners = this.getStepOwners(step);
      for (const owner of owners) {
        notifications.push({
          type: 'predecessor_complete',
          recipientEmail: owner,
          recipientTeamsId: '',
          title: `✅ Unblocked: Step ${step.id} ready to start`,
          message: `Predecessor **${completedStepId}** completed. Your step **${step.description}** is now unblocked and ready to start.`,
          steps: [step],
          priority: 'medium',
        });
      }
    }

    return notifications;
  }

  // ---- Helpers ----

  private getStepOwners(step: RolloverStep): string[] {
    return [step.wwicPoc, step.fedPoc, step.engineeringDri, step.engineeringLead]
      .filter(o => o && o !== '-' && o !== '')
      .flatMap(o => o.split(/[,&]/).map(n => n.trim()))
      .filter(n => n.length > 0);
  }

  private daysSince(dateStr: string): number {
    const date = new Date(dateStr);
    const now = new Date();
    return Math.floor((now.getTime() - date.getTime()) / 86400000);
  }
}
