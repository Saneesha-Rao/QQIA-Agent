import { RolloverStep } from '../models/types';
import { WorkstreamStats } from '../services/dependencyEngine';

/**
 * Adaptive Card builders for Teams bot responses.
 * Uses the Adaptive Cards schema (v1.5) for rich interactive messages.
 */

/** Overall rollover progress dashboard card */
export function buildOverallDashboardCard(
  allSteps: RolloverStep[],
  workstreamStats: Map<string, WorkstreamStats>,
  overdue: RolloverStep[],
  blocked: RolloverStep[],
  track: 'Corp' | 'Fed' = 'Corp'
): any {
  const totalSteps = allSteps.length;
  const completed = allSteps.filter(s =>
    (track === 'Corp' ? s.corpStatus : s.fedStatus) === 'Completed'
  ).length;
  const pct = Math.round((completed / totalSteps) * 100);

  const workstreamRows = [];
  for (const [ws, stats] of workstreamStats) {
    const wsPct = stats.total > 0 ? Math.round((stats.completed / stats.total) * 100) : 0;
    workstreamRows.push({
      type: 'ColumnSet',
      columns: [
        { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: ws, size: 'Small' }] },
        { type: 'Column', width: '80px', items: [{ type: 'TextBlock', text: `${stats.completed}/${stats.total}`, size: 'Small', horizontalAlignment: 'Center' }] },
        { type: 'Column', width: '60px', items: [{ type: 'TextBlock', text: `${wsPct}%`, size: 'Small', weight: 'Bolder', color: wsPct === 100 ? 'Good' : wsPct > 50 ? 'Warning' : 'Default', horizontalAlignment: 'Right' }] },
      ],
    });
  }

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'TextBlock',
        text: `📊 FY27 Rollover Dashboard (${track})`,
        size: 'Large',
        weight: 'Bolder',
      },
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column', width: 'auto',
            items: [
              { type: 'TextBlock', text: `${pct}%`, size: 'ExtraLarge', weight: 'Bolder', color: pct === 100 ? 'Good' : pct > 50 ? 'Warning' : 'Attention' },
              { type: 'TextBlock', text: 'Complete', size: 'Small', isSubtle: true },
            ],
          },
          {
            type: 'Column', width: 'auto',
            items: [
              { type: 'TextBlock', text: `${completed}`, size: 'ExtraLarge', weight: 'Bolder', color: 'Good' },
              { type: 'TextBlock', text: 'Done', size: 'Small', isSubtle: true },
            ],
          },
          {
            type: 'Column', width: 'auto',
            items: [
              { type: 'TextBlock', text: `${overdue.length}`, size: 'ExtraLarge', weight: 'Bolder', color: 'Attention' },
              { type: 'TextBlock', text: 'Overdue', size: 'Small', isSubtle: true },
            ],
          },
          {
            type: 'Column', width: 'auto',
            items: [
              { type: 'TextBlock', text: `${blocked.length}`, size: 'ExtraLarge', weight: 'Bolder', color: 'Attention' },
              { type: 'TextBlock', text: 'Blocked', size: 'Small', isSubtle: true },
            ],
          },
        ],
      },
      { type: 'TextBlock', text: 'Workstream Breakdown', weight: 'Bolder', spacing: 'Medium' },
      ...workstreamRows,
    ],
    actions: [
      { type: 'Action.Submit', title: '🔍 View Overdue', data: { action: 'view_overdue', track } },
      { type: 'Action.Submit', title: '🛑 View Blocked', data: { action: 'view_blocked', track } },
      { type: 'Action.Submit', title: '🔗 Critical Path', data: { action: 'critical_path', track } },
    ],
  };
}

/** Step detail card for individual step view */
export function buildStepDetailCard(step: RolloverStep, blockers: string[], blockedBy: string[]): any {
  const statusColor = (s: string) =>
    s === 'Completed' ? 'Good' : s === 'In Progress' ? 'Warning' : s === 'Blocked' ? 'Attention' : 'Default';

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: `Step ${step.id}: ${step.description}`, size: 'Medium', weight: 'Bolder', wrap: true },
      { type: 'TextBlock', text: `Workstream: ${step.workstream}`, isSubtle: true },
      {
        type: 'FactSet',
        facts: [
          { title: 'Corp Status', value: step.corpStatus },
          { title: 'Fed Status', value: step.fedStatus },
          { title: 'Corp Dates', value: `${step.corpStartDate || 'TBD'} → ${step.corpEndDate || 'TBD'}` },
          { title: 'Fed Dates', value: `${step.fedStartDate || 'TBD'} → ${step.fedEndDate || 'TBD'}` },
          { title: 'WWIC POC', value: step.wwicPoc || '-' },
          { title: 'Fed POC', value: step.fedPoc || '-' },
          { title: 'Engineering DRI', value: step.engineeringDri || '-' },
          { title: 'Dependencies', value: step.dependencies.join(', ') || 'None' },
        ],
      },
      ...(blockers.length > 0 ? [{
        type: 'TextBlock' as const,
        text: `🛑 Blocked by: ${blockers.join(', ')}`,
        color: 'Attention' as const,
        wrap: true,
      }] : []),
      ...(blockedBy.length > 0 ? [{
        type: 'TextBlock' as const,
        text: `⏳ Blocking: ${blockedBy.join(', ')}`,
        color: 'Warning' as const,
        wrap: true,
      }] : []),
      ...(step.referenceNotes ? [{
        type: 'TextBlock' as const,
        text: `📝 Notes: ${step.referenceNotes}`,
        wrap: true,
        isSubtle: true,
      }] : []),
    ],
    actions: [
      {
        type: 'Action.Submit', title: '✅ Mark Complete',
        data: { action: 'update_status', stepId: step.id, field: 'corpStatus', newStatus: 'Completed' },
      },
      {
        type: 'Action.Submit', title: '🔄 In Progress',
        data: { action: 'update_status', stepId: step.id, field: 'corpStatus', newStatus: 'In Progress' },
      },
      {
        type: 'Action.Submit', title: '🛑 Mark Blocked',
        data: { action: 'update_status', stepId: step.id, field: 'corpStatus', newStatus: 'Blocked' },
      },
    ],
  };
}

/** Step list card for query results */
export function buildStepListCard(title: string, steps: RolloverStep[], track: 'Corp' | 'Fed' = 'Corp'): any {
  const statusEmoji = (s: string) =>
    s === 'Completed' ? '✅' : s === 'In Progress' ? '🔄' : s === 'Blocked' ? '🛑' : '⬜';

  // Column headers
  const header = {
    type: 'ColumnSet',
    separator: true,
    columns: [
      { type: 'Column', width: '35px', items: [{ type: 'TextBlock', text: 'ID', weight: 'Bolder', size: 'Small' }] },
      { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: 'Description', weight: 'Bolder', size: 'Small' }] },
      { type: 'Column', width: '100px', items: [{ type: 'TextBlock', text: 'Grouping', weight: 'Bolder', size: 'Small' }] },
      { type: 'Column', width: '75px', items: [{ type: 'TextBlock', text: 'Status', weight: 'Bolder', size: 'Small' }] },
      { type: 'Column', width: '75px', items: [{ type: 'TextBlock', text: 'Due', weight: 'Bolder', size: 'Small' }] },
    ],
  };

  const items = steps.slice(0, 15).map(step => {
    const status = track === 'Corp' ? step.corpStatus : step.fedStatus;
    const endDate = track === 'Corp' ? step.corpEndDate : step.fedEndDate;
    return {
      type: 'ColumnSet',
      columns: [
        { type: 'Column', width: '35px', items: [{ type: 'TextBlock', text: step.id, weight: 'Bolder', size: 'Small' }] },
        { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: step.description, size: 'Small', wrap: true }] },
        { type: 'Column', width: '100px', items: [{ type: 'TextBlock', text: step.workstream || '-', size: 'Small', wrap: true, isSubtle: true }] },
        { type: 'Column', width: '75px', items: [{ type: 'TextBlock', text: `${statusEmoji(status)} ${status}`, size: 'Small' }] },
        { type: 'Column', width: '75px', items: [{ type: 'TextBlock', text: endDate || 'TBD', size: 'Small', isSubtle: true }] },
      ],
    };
  });

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: title, size: 'Medium', weight: 'Bolder' },
      { type: 'TextBlock', text: `Showing ${Math.min(steps.length, 15)} of ${steps.length} steps`, isSubtle: true },
      header,
      ...items,
    ],
  };
}

/** Personal task summary card for a specific DRI */
export function buildMyTasksCard(ownerName: string, steps: RolloverStep[]): any {
  const pending = steps.filter(s => s.corpStatus !== 'Completed' && s.corpStatus !== 'N/A');
  const completed = steps.filter(s => s.corpStatus === 'Completed');

  const taskHeader = {
    type: 'ColumnSet',
    separator: true,
    columns: [
      { type: 'Column', width: '35px', items: [{ type: 'TextBlock', text: 'ID', weight: 'Bolder', size: 'Small' }] },
      { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: 'Description', weight: 'Bolder', size: 'Small' }] },
      { type: 'Column', width: '100px', items: [{ type: 'TextBlock', text: 'Grouping', weight: 'Bolder', size: 'Small' }] },
      { type: 'Column', width: '75px', items: [{ type: 'TextBlock', text: 'Status', weight: 'Bolder', size: 'Small' }] },
      { type: 'Column', width: '75px', items: [{ type: 'TextBlock', text: 'Due', weight: 'Bolder', size: 'Small' }] },
    ],
  };

  const taskItems = pending.map(step => ({
    type: 'ColumnSet',
    columns: [
      { type: 'Column', width: '35px', items: [{ type: 'TextBlock', text: step.id, weight: 'Bolder', size: 'Small' }] },
      { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: step.description, size: 'Small', wrap: true }] },
      { type: 'Column', width: '100px', items: [{ type: 'TextBlock', text: step.workstream || '-', size: 'Small', wrap: true, isSubtle: true }] },
      { type: 'Column', width: '75px', items: [{ type: 'TextBlock', text: step.corpStatus, size: 'Small' }] },
      { type: 'Column', width: '75px', items: [{ type: 'TextBlock', text: step.corpEndDate || 'TBD', size: 'Small' }] },
    ],
    selectAction: { type: 'Action.Submit', data: { action: 'view_step', stepId: step.id } },
  }));

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: `📋 Tasks for ${ownerName}`, size: 'Medium', weight: 'Bolder' },
      { type: 'TextBlock', text: `${completed.length} completed, ${pending.length} remaining`, isSubtle: true },
      taskHeader,
      ...taskItems,
    ],
    actions: [
      { type: 'Action.Submit', title: '📊 Full Dashboard', data: { action: 'dashboard' } },
    ],
  };
}
