import { Notification } from '../services/notificationService';
import { RolloverStep } from '../models/types';

/**
 * Adaptive Card templates for proactive notifications.
 * These are sent 1:1 to DRIs or posted to channels.
 */

/** Deadline approaching notification card */
export function buildDeadlineCard(notification: Notification): any {
  const step = notification.steps[0];
  if (!step) return buildSimpleNotificationCard(notification);

  const daysText = notification.title.match(/(\d+) day/)?.[1] || '?';
  const urgencyColor = parseInt(daysText) <= 1 ? 'Attention' : 'Warning';

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: urgencyColor === 'Attention' ? 'attention' : 'warning',
        items: [
          { type: 'TextBlock', text: `⏰ Deadline in ${daysText} day(s)`, size: 'Medium', weight: 'Bolder', color: urgencyColor },
        ],
      },
      { type: 'TextBlock', text: `**Step ${step.id}**: ${step.description}`, wrap: true, spacing: 'Medium' },
      {
        type: 'FactSet',
        facts: [
          { title: 'Workstream', value: step.workstream },
          { title: 'Due Date', value: step.corpEndDate || 'TBD' },
          { title: 'Current Status', value: step.corpStatus },
          { title: 'Owner', value: step.wwicPoc || step.engineeringDri || '-' },
        ],
      },
    ],
    actions: [
      { type: 'Action.Submit', title: '✅ Mark Complete', data: { action: 'update_status', stepId: step.id, field: 'corpStatus', newStatus: 'Completed' } },
      { type: 'Action.Submit', title: '🔄 In Progress', data: { action: 'update_status', stepId: step.id, field: 'corpStatus', newStatus: 'In Progress' } },
      { type: 'Action.Submit', title: '🛑 Report Blocker', data: { action: 'update_status', stepId: step.id, field: 'corpStatus', newStatus: 'Blocked' } },
      { type: 'Action.Submit', title: '📋 View Details', data: { action: 'view_step', stepId: step.id } },
    ],
  };
}

/** Overdue step notification card */
export function buildOverdueCard(notification: Notification): any {
  const step = notification.steps[0];
  if (!step) return buildSimpleNotificationCard(notification);

  const daysOverdue = step.corpEndDate
    ? Math.floor((Date.now() - new Date(step.corpEndDate).getTime()) / 86400000)
    : 0;

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'attention',
        items: [
          { type: 'TextBlock', text: `🚨 OVERDUE: ${daysOverdue} day(s) past due`, size: 'Medium', weight: 'Bolder', color: 'Attention' },
        ],
      },
      { type: 'TextBlock', text: `**Step ${step.id}**: ${step.description}`, wrap: true, spacing: 'Medium' },
      {
        type: 'FactSet',
        facts: [
          { title: 'Was Due', value: step.corpEndDate || 'TBD' },
          { title: 'Days Overdue', value: `${daysOverdue}` },
          { title: 'Status', value: step.corpStatus },
          { title: 'Owner', value: step.wwicPoc || step.engineeringDri || '-' },
          { title: 'Workstream', value: step.workstream },
        ],
      },
      ...(step.referenceNotes ? [{ type: 'TextBlock' as const, text: `📝 ${step.referenceNotes}`, wrap: true, isSubtle: true }] : []),
    ],
    actions: [
      { type: 'Action.Submit', title: '✅ Mark Complete', data: { action: 'update_status', stepId: step.id, field: 'corpStatus', newStatus: 'Completed' } },
      { type: 'Action.Submit', title: '📝 Add Note', data: { action: 'add_note_prompt', stepId: step.id } },
    ],
  };
}

/** Predecessor completed / step unblocked card */
export function buildUnblockedCard(notification: Notification): any {
  const step = notification.steps[0];
  if (!step) return buildSimpleNotificationCard(notification);

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'good',
        items: [
          { type: 'TextBlock', text: '✅ Step Unblocked!', size: 'Medium', weight: 'Bolder', color: 'Good' },
        ],
      },
      { type: 'TextBlock', text: notification.message, wrap: true, spacing: 'Medium' },
      { type: 'TextBlock', text: `**Your ${step.id}**: ${step.description}`, wrap: true },
      {
        type: 'FactSet',
        facts: [
          { title: 'Start Date', value: step.corpStartDate || 'TBD' },
          { title: 'End Date', value: step.corpEndDate || 'TBD' },
          { title: 'Workstream', value: step.workstream },
        ],
      },
    ],
    actions: [
      { type: 'Action.Submit', title: '🔄 Start Work', data: { action: 'update_status', stepId: step.id, field: 'corpStatus', newStatus: 'In Progress' } },
      { type: 'Action.Submit', title: '📋 View Details', data: { action: 'view_step', stepId: step.id } },
    ],
  };
}

/** Weekly digest summary card */
export function buildWeeklyDigestCard(
  totalSteps: number,
  completed: number,
  inProgress: number,
  blocked: number,
  overdue: RolloverStep[],
  upcomingDeadlines: RolloverStep[]
): any {
  const pct = Math.round((completed / totalSteps) * 100);

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: '📊 Weekly Rollover Digest', size: 'Large', weight: 'Bolder' },
      { type: 'TextBlock', text: new Date().toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }), isSubtle: true },
      {
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [
            { type: 'TextBlock', text: `${pct}%`, size: 'ExtraLarge', weight: 'Bolder', color: pct >= 75 ? 'Good' : pct >= 50 ? 'Warning' : 'Default' },
            { type: 'TextBlock', text: 'Overall', size: 'Small', isSubtle: true },
          ]},
          { type: 'Column', width: 'auto', items: [
            { type: 'TextBlock', text: `${completed}`, size: 'ExtraLarge', weight: 'Bolder', color: 'Good' },
            { type: 'TextBlock', text: 'Done', size: 'Small', isSubtle: true },
          ]},
          { type: 'Column', width: 'auto', items: [
            { type: 'TextBlock', text: `${inProgress}`, size: 'ExtraLarge', weight: 'Bolder', color: 'Warning' },
            { type: 'TextBlock', text: 'Active', size: 'Small', isSubtle: true },
          ]},
          { type: 'Column', width: 'auto', items: [
            { type: 'TextBlock', text: `${overdue.length}`, size: 'ExtraLarge', weight: 'Bolder', color: 'Attention' },
            { type: 'TextBlock', text: 'Overdue', size: 'Small', isSubtle: true },
          ]},
        ],
      },
      ...(overdue.length > 0 ? [
        { type: 'TextBlock' as const, text: '⚠️ **Overdue Items**', spacing: 'Medium' as const },
        ...overdue.slice(0, 5).map(s => ({
          type: 'TextBlock' as const,
          text: `• **${s.id}** ${s.description} (Due: ${s.corpEndDate})`,
          wrap: true,
          size: 'Small' as const,
        })),
      ] : []),
      ...(upcomingDeadlines.length > 0 ? [
        { type: 'TextBlock' as const, text: '📅 **Due This Week**', spacing: 'Medium' as const },
        ...upcomingDeadlines.slice(0, 5).map(s => ({
          type: 'TextBlock' as const,
          text: `• **${s.id}** ${s.description} (Due: ${s.corpEndDate})`,
          wrap: true,
          size: 'Small' as const,
        })),
      ] : []),
    ],
    actions: [
      { type: 'Action.Submit', title: '📊 Full Dashboard', data: { action: 'dashboard' } },
      { type: 'Action.Submit', title: '🛑 View Blocked', data: { action: 'view_blocked' } },
    ],
  };
}

/** Escalation card for PMs/leadership */
export function buildEscalationCard(notification: Notification): any {
  const step = notification.steps[0];

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'attention',
        items: [
          { type: 'TextBlock', text: '🔴 ESCALATION', size: 'Medium', weight: 'Bolder', color: 'Attention' },
          { type: 'TextBlock', text: notification.message, wrap: true },
        ],
      },
      ...(step ? [{
        type: 'FactSet' as const,
        facts: [
          { title: 'Step', value: `${step.id}: ${step.description}` },
          { title: 'Owner', value: step.wwicPoc || step.engineeringDri || '-' },
          { title: 'Was Due', value: step.corpEndDate || 'TBD' },
          { title: 'Workstream', value: step.workstream },
        ],
      }] : []),
    ],
    actions: [
      ...(step ? [
        { type: 'Action.Submit', title: '📋 View Step', data: { action: 'view_step', stepId: step.id } },
      ] : []),
      { type: 'Action.Submit', title: '📊 Dashboard', data: { action: 'dashboard' } },
    ],
  };
}

/** Simple text notification card (fallback) */
function buildSimpleNotificationCard(notification: Notification): any {
  const priorityStyle = notification.priority === 'high' ? 'attention' :
    notification.priority === 'medium' ? 'warning' : 'default';

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: priorityStyle,
        items: [
          { type: 'TextBlock', text: notification.title, weight: 'Bolder', wrap: true },
          { type: 'TextBlock', text: notification.message, wrap: true },
        ],
      },
    ],
  };
}
