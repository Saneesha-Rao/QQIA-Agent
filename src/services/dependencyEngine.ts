import { RolloverStep } from '../models/types';

export interface DependencyNode {
  stepId: string;
  dependencies: string[];
  dependents: string[];
  corpStatus: string;
  fedStatus: string;
}

/**
 * DAG-based dependency engine for rollover step management.
 * Tracks critical paths, blockers, and auto-unblocking.
 */
export class DependencyEngine {
  private nodes: Map<string, DependencyNode> = new Map();

  /** Build the dependency graph from rollover steps */
  buildGraph(steps: RolloverStep[]): void {
    this.nodes.clear();

    // Initialize nodes
    for (const step of steps) {
      this.nodes.set(step.id, {
        stepId: step.id,
        dependencies: step.dependencies,
        dependents: [],
        corpStatus: step.corpStatus,
        fedStatus: step.fedStatus,
      });
    }

    // Build reverse edges (dependents)
    for (const [id, node] of this.nodes) {
      for (const depId of node.dependencies) {
        const depNode = this.nodes.get(depId);
        if (depNode) {
          depNode.dependents.push(id);
        }
      }
    }
  }

  /** Check if a step has all Corp predecessors completed */
  isCorpUnblocked(stepId: string): boolean {
    const node = this.nodes.get(stepId);
    if (!node) return false;
    return node.dependencies.every(depId => {
      const dep = this.nodes.get(depId);
      return !dep || dep.corpStatus === 'Completed' || dep.corpStatus === 'N/A';
    });
  }

  /** Check if a step has all Fed predecessors completed */
  isFedUnblocked(stepId: string): boolean {
    const node = this.nodes.get(stepId);
    if (!node) return false;
    return node.dependencies.every(depId => {
      const dep = this.nodes.get(depId);
      return !dep || dep.fedStatus === 'Completed' || dep.fedStatus === 'N/A';
    });
  }

  /** Get all steps blocked by a specific step */
  getBlockedBy(stepId: string): string[] {
    const node = this.nodes.get(stepId);
    if (!node) return [];
    return node.dependents.filter(depId => {
      const dep = this.nodes.get(depId);
      return dep && dep.corpStatus !== 'Completed' && dep.corpStatus !== 'N/A';
    });
  }

  /** Get all blocker steps (incomplete predecessors) for a given step */
  getBlockers(stepId: string): string[] {
    const node = this.nodes.get(stepId);
    if (!node) return [];
    return node.dependencies.filter(depId => {
      const dep = this.nodes.get(depId);
      return dep && dep.corpStatus !== 'Completed' && dep.corpStatus !== 'N/A';
    });
  }

  /** Compute critical path using longest path through the DAG */
  getCriticalPath(track: 'Corp' | 'Fed' = 'Corp'): string[] {
    const statusField = track === 'Corp' ? 'corpStatus' : 'fedStatus';
    const incomplete = new Map<string, DependencyNode>();
    for (const [id, node] of this.nodes) {
      if ((node as any)[statusField] !== 'Completed' && (node as any)[statusField] !== 'N/A') {
        incomplete.set(id, node);
      }
    }

    // Find all paths from sources (no incomplete deps) to sinks (no incomplete dependents)
    const sources: string[] = [];
    for (const [id, node] of incomplete) {
      const hasIncompleteDep = node.dependencies.some(d => incomplete.has(d));
      if (!hasIncompleteDep) sources.push(id);
    }

    let longestPath: string[] = [];
    const dfs = (current: string, path: string[]) => {
      path.push(current);
      const node = incomplete.get(current);
      if (!node) return;

      const nextSteps = node.dependents.filter(d => incomplete.has(d));
      if (nextSteps.length === 0) {
        if (path.length > longestPath.length) {
          longestPath = [...path];
        }
      } else {
        for (const next of nextSteps) {
          dfs(next, path);
        }
      }
      path.pop();
    };

    for (const source of sources) {
      dfs(source, []);
    }
    return longestPath;
  }

  /** Get steps that are newly unblocked after a step is completed */
  getNewlyUnblocked(completedStepId: string, track: 'Corp' | 'Fed' = 'Corp'): string[] {
    const node = this.nodes.get(completedStepId);
    if (!node) return [];

    const isUnblocked = track === 'Corp' ? this.isCorpUnblocked.bind(this) : this.isFedUnblocked.bind(this);
    return node.dependents.filter(depId => isUnblocked(depId));
  }

  /** Get progress stats by workstream */
  getWorkstreamStats(steps: RolloverStep[], track: 'Corp' | 'Fed' = 'Corp'): Map<string, WorkstreamStats> {
    const stats = new Map<string, WorkstreamStats>();
    for (const step of steps) {
      const status = track === 'Corp' ? step.corpStatus : step.fedStatus;
      if (!stats.has(step.workstream)) {
        stats.set(step.workstream, { total: 0, completed: 0, inProgress: 0, blocked: 0, notStarted: 0 });
      }
      const ws = stats.get(step.workstream)!;
      ws.total++;
      if (status === 'Completed') ws.completed++;
      else if (status === 'In Progress') ws.inProgress++;
      else if (status === 'Blocked') ws.blocked++;
      else ws.notStarted++;
    }
    return stats;
  }

  /** Validate the graph for cycles (should be a valid DAG) */
  validateDAG(): { valid: boolean; cycles: string[][] } {
    const visited = new Set<string>();
    const inStack = new Set<string>();
    const cycles: string[][] = [];

    const dfs = (nodeId: string, path: string[]): boolean => {
      if (inStack.has(nodeId)) {
        const cycleStart = path.indexOf(nodeId);
        cycles.push(path.slice(cycleStart));
        return true;
      }
      if (visited.has(nodeId)) return false;

      visited.add(nodeId);
      inStack.add(nodeId);
      path.push(nodeId);

      const node = this.nodes.get(nodeId);
      if (node) {
        for (const depId of node.dependents) {
          dfs(depId, path);
        }
      }

      path.pop();
      inStack.delete(nodeId);
      return false;
    };

    for (const nodeId of this.nodes.keys()) {
      if (!visited.has(nodeId)) {
        dfs(nodeId, []);
      }
    }

    return { valid: cycles.length === 0, cycles };
  }

  /** Get upcoming deadline steps within N days */
  getUpcomingDeadlines(steps: RolloverStep[], withinDays: number, track: 'Corp' | 'Fed' = 'Corp'): RolloverStep[] {
    const now = new Date();
    const cutoff = new Date(now.getTime() + withinDays * 86400000);

    return steps.filter(step => {
      const status = track === 'Corp' ? step.corpStatus : step.fedStatus;
      const endDate = track === 'Corp' ? step.corpEndDate : step.fedEndDate;
      if (status === 'Completed' || status === 'N/A' || !endDate) return false;

      const deadline = new Date(endDate);
      return deadline >= now && deadline <= cutoff;
    });
  }

  /** Get overdue steps */
  getOverdueSteps(steps: RolloverStep[], track: 'Corp' | 'Fed' = 'Corp'): RolloverStep[] {
    const now = new Date();
    return steps.filter(step => {
      const status = track === 'Corp' ? step.corpStatus : step.fedStatus;
      const endDate = track === 'Corp' ? step.corpEndDate : step.fedEndDate;
      if (status === 'Completed' || status === 'N/A' || !endDate) return false;
      return new Date(endDate) < now;
    });
  }
}

export interface WorkstreamStats {
  total: number;
  completed: number;
  inProgress: number;
  blocked: number;
  notStarted: number;
}
