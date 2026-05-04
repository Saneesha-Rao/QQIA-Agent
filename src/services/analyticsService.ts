/**
 * Analytics Service — shared logic for burndown, workstream health,
 * timeline, and change feed. Used by both web API and bot commands.
 */
import { RolloverStep, Milestone, AuditEntry } from '../models/types';
import { DependencyEngine, WorkstreamStats } from './dependencyEngine';

// ---- Types ----

export interface BurndownPoint {
  date: string;         // ISO date (YYYY-MM-DD)
  completed: number;
  inProgress: number;
  blocked: number;
  notStarted: number;
  total: number;
}

export interface DailySnapshot {
  date: string;
  track: 'Corp' | 'Fed';
  completed: number;
  inProgress: number;
  blocked: number;
  notStarted: number;
  total: number;
}

export interface WorkstreamHealth {
  workstream: string;
  track: string;
  total: number;
  completed: number;
  inProgress: number;
  blocked: number;
  notStarted: number;
  overdue: number;
  completionPct: number;
  health: 'green' | 'yellow' | 'red';   // 🟢 🟡 🔴
  riskScore: number;                      // 0-100, higher = riskier
}

export interface TimelineItem {
  id: string;
  name: string;
  workstream: string;
  start: string | null;
  end: string | null;
  status: string;
  completedDate: string | null;
  dependencies: string[];
  owner: string;
}

export interface ChangeEntry {
  stepId: string;
  field: string;
  previousValue: string;
  newValue: string;
  changedBy: string;
  changedAt: string;
  source: string;
}

// ---- Service ----

export class AnalyticsService {
  /** In-memory daily snapshots (persists for session, seeded on startup) */
  private snapshots: DailySnapshot[] = [];

  /**
   * Record a daily snapshot of current step statuses.
   * Call this once per day (or on startup to seed today's snapshot).
   */
  recordSnapshot(steps: RolloverStep[], track: 'Corp' | 'Fed' = 'Corp'): DailySnapshot {
    const today = new Date().toISOString().split('T')[0];
    // Remove existing snapshot for today+track if any
    this.snapshots = this.snapshots.filter(s => !(s.date === today && s.track === track));

    let completed = 0, inProgress = 0, blocked = 0, notStarted = 0;
    for (const step of steps) {
      const status = track === 'Corp' ? step.corpStatus : step.fedStatus;
      if (status === 'Completed') completed++;
      else if (status === 'In Progress') inProgress++;
      else if (status === 'Blocked') blocked++;
      else if (status !== 'N/A') notStarted++;
    }

    const snap: DailySnapshot = {
      date: today, track, completed, inProgress, blocked,
      notStarted, total: completed + inProgress + blocked + notStarted,
    };
    this.snapshots.push(snap);
    this.snapshots.sort((a, b) => a.date.localeCompare(b.date));
    return snap;
  }

  /**
   * Get burndown data — uses stored snapshots + reconstructs history from audit if available.
   */
  getBurndown(track: 'Corp' | 'Fed' = 'Corp', auditEntries?: AuditEntry[]): BurndownPoint[] {
    const trackSnapshots = this.snapshots.filter(s => s.track === track);

    // If we have audit entries, try to reconstruct daily history
    if (auditEntries && auditEntries.length > 0 && trackSnapshots.length <= 1) {
      return this.reconstructFromAudit(auditEntries, track, trackSnapshots[0]);
    }

    return trackSnapshots.map(s => ({
      date: s.date,
      completed: s.completed,
      inProgress: s.inProgress,
      blocked: s.blocked,
      notStarted: s.notStarted,
      total: s.total,
    }));
  }

  /**
   * Best-effort burndown from audit trail.
   * Works backwards from current state.
   */
  private reconstructFromAudit(
    entries: AuditEntry[], track: 'Corp' | 'Fed', currentSnap?: DailySnapshot
  ): BurndownPoint[] {
    if (!currentSnap) return [];

    const field = track === 'Corp' ? 'corpStatus' : 'fedStatus';
    const statusEntries = entries
      .filter(e => e.field === field)
      .sort((a, b) => new Date(b.changedAt).getTime() - new Date(a.changedAt).getTime());

    if (statusEntries.length === 0) {
      return [{ ...currentSnap }];
    }

    // Group changes by date
    const changesByDate = new Map<string, AuditEntry[]>();
    for (const e of statusEntries) {
      const d = new Date(e.changedAt).toISOString().split('T')[0];
      if (!changesByDate.has(d)) changesByDate.set(d, []);
      changesByDate.get(d)!.push(e);
    }

    // Walk backwards from current snapshot
    const points: BurndownPoint[] = [];
    let { completed, inProgress, blocked, notStarted, total } = currentSnap;
    points.push({ date: currentSnap.date, completed, inProgress, blocked, notStarted, total });

    const sortedDates = Array.from(changesByDate.keys()).sort().reverse();
    for (const date of sortedDates) {
      if (date >= currentSnap.date) continue;
      const changes = changesByDate.get(date)!;
      // Undo each change
      for (const c of changes) {
        this.adjustCount(c.newValue, -1, { completed, inProgress, blocked, notStarted });
        this.adjustCount(c.previousValue, 1, { completed, inProgress, blocked, notStarted });
        // Read back
        const adj = this.adjustCount(c.newValue, 0, { completed, inProgress, blocked, notStarted });
        completed = adj.completed; inProgress = adj.inProgress;
        blocked = adj.blocked; notStarted = adj.notStarted;
      }
      points.unshift({ date, completed, inProgress, blocked, notStarted, total });
    }

    return points;
  }

  private adjustCount(status: string, delta: number, counts: any): any {
    if (status === 'Completed') counts.completed += delta;
    else if (status === 'In Progress') counts.inProgress += delta;
    else if (status === 'Blocked') counts.blocked += delta;
    else if (status !== 'N/A') counts.notStarted += delta;
    // Clamp to 0
    counts.completed = Math.max(0, counts.completed);
    counts.inProgress = Math.max(0, counts.inProgress);
    counts.blocked = Math.max(0, counts.blocked);
    counts.notStarted = Math.max(0, counts.notStarted);
    return counts;
  }

  /**
   * Workstream health analysis with risk scoring.
   */
  getWorkstreamHealth(
    steps: RolloverStep[],
    depEngine: DependencyEngine,
    track: 'Corp' | 'Fed' = 'Corp'
  ): WorkstreamHealth[] {
    const wsStats = depEngine.getWorkstreamStats(steps, track);
    const overdueSteps = depEngine.getOverdueSteps(steps, track);

    // Count overdue per workstream
    const overdueByWs = new Map<string, number>();
    for (const s of overdueSteps) {
      overdueByWs.set(s.workstream, (overdueByWs.get(s.workstream) || 0) + 1);
    }

    const results: WorkstreamHealth[] = [];
    for (const [ws, stat] of wsStats) {
      const overdue = overdueByWs.get(ws) || 0;
      const completionPct = stat.total > 0 ? Math.round((stat.completed / stat.total) * 100) : 0;

      // Risk score: 0-100. Higher = riskier
      // Factors: blocked %, overdue %, not-started % (weighted)
      const blockedPct = stat.total > 0 ? (stat.blocked / stat.total) * 100 : 0;
      const overduePct = stat.total > 0 ? (overdue / stat.total) * 100 : 0;
      const notStartedPct = stat.total > 0 ? (stat.notStarted / stat.total) * 100 : 0;
      const riskScore = Math.min(100, Math.round(
        blockedPct * 1.5 + overduePct * 2.0 + notStartedPct * 0.3
      ));

      // Health: green (>70% complete, no blockers), yellow (40-70% or few issues), red (<40% or major issues)
      let health: 'green' | 'yellow' | 'red';
      if (stat.blocked > 0 || overdue > 2 || completionPct < 40) {
        health = 'red';
      } else if (overdue > 0 || completionPct < 70) {
        health = 'yellow';
      } else {
        health = 'green';
      }

      results.push({
        workstream: ws, track, total: stat.total,
        completed: stat.completed, inProgress: stat.inProgress,
        blocked: stat.blocked, notStarted: stat.notStarted,
        overdue, completionPct, health, riskScore,
      });
    }

    // Sort by risk (highest first)
    results.sort((a, b) => b.riskScore - a.riskScore);
    return results;
  }

  /**
   * Timeline data for Gantt view.
   */
  getTimeline(steps: RolloverStep[], track: 'Corp' | 'Fed' = 'Corp'): TimelineItem[] {
    return steps
      .filter(s => {
        const status = track === 'Corp' ? s.corpStatus : s.fedStatus;
        return status !== 'N/A';
      })
      .map(s => ({
        id: s.id,
        name: s.description || s.id,
        workstream: s.workstream,
        start: track === 'Corp' ? s.corpStartDate : s.fedStartDate,
        end: track === 'Corp' ? s.corpEndDate : s.fedEndDate,
        status: track === 'Corp' ? s.corpStatus : s.fedStatus,
        completedDate: track === 'Corp' ? s.corpCompletedDate || null : null,
        dependencies: s.dependencies || [],
        owner: s.engineeringDri || s.wwicPoc || '',
      }))
      .sort((a, b) => {
        // Sort by workstream, then by start date
        const wsCmp = a.workstream.localeCompare(b.workstream);
        if (wsCmp !== 0) return wsCmp;
        if (!a.start) return 1;
        if (!b.start) return -1;
        return a.start.localeCompare(b.start);
      });
  }

  /**
   * Recent changes feed from audit entries.
   */
  getRecentChanges(auditEntries: AuditEntry[], hours: number = 24): ChangeEntry[] {
    const cutoff = new Date(Date.now() - hours * 3600000);
    return auditEntries
      .filter(e => new Date(e.changedAt) >= cutoff)
      .sort((a, b) => new Date(b.changedAt).getTime() - new Date(a.changedAt).getTime())
      .map(e => ({
        stepId: e.stepId,
        field: e.field,
        previousValue: e.previousValue,
        newValue: e.newValue,
        changedBy: e.changedBy,
        changedAt: e.changedAt,
        source: e.source || 'unknown',
      }));
  }

  /** Format workstream health for Teams text response */
  formatHealthForTeams(healthData: WorkstreamHealth[]): string {
    if (healthData.length === 0) return 'No workstream data available.';
    const emoji = { green: '🟢', yellow: '🟡', red: '🔴' };
    let msg = `## Workstream Health (${healthData[0].track})\n\n`;
    for (const ws of healthData) {
      msg += `${emoji[ws.health]} **${ws.workstream}** — ${ws.completionPct}% complete`;
      msg += ` (${ws.completed}/${ws.total})`;
      if (ws.blocked > 0) msg += ` | ⛔ ${ws.blocked} blocked`;
      if (ws.overdue > 0) msg += ` | ⏰ ${ws.overdue} overdue`;
      msg += `\n`;
    }
    return msg;
  }

  /** Format recent changes for Teams text response */
  formatChangesForTeams(changes: ChangeEntry[]): string {
    if (changes.length === 0) return '📋 No changes in the last 24 hours.';
    let msg = `## 📋 Recent Changes (last 24h)\n\n`;
    const grouped = new Map<string, ChangeEntry[]>();
    for (const c of changes) {
      if (!grouped.has(c.stepId)) grouped.set(c.stepId, []);
      grouped.get(c.stepId)!.push(c);
    }
    for (const [stepId, entries] of grouped) {
      msg += `**${stepId}**:\n`;
      for (const e of entries) {
        msg += `  • ${e.field}: ${e.previousValue || '—'} → **${e.newValue}** (by ${e.changedBy})\n`;
      }
    }
    return msg;
  }
}
