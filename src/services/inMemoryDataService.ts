import { RolloverStep, Milestone, AuditEntry, UserProfile } from '../models/types';

/**
 * In-memory data store for local development without Cosmos DB.
 * Implements the same interface as DataService but stores everything in Maps.
 */
export class InMemoryDataService {
  private steps: Map<string, RolloverStep> = new Map();
  private milestones: Map<string, Milestone> = new Map();
  private audit: AuditEntry[] = [];
  private users: Map<string, UserProfile> = new Map();

  async initialize(): Promise<void> {
    console.log('✅ In-memory data store initialized');
  }

  // ---- Steps ----

  async upsertStep(step: RolloverStep): Promise<void> {
    this.steps.set(step.id, { ...step });
  }

  async getStep(stepId: string): Promise<RolloverStep | null> {
    return this.steps.get(stepId) || null;
  }

  async getAllSteps(): Promise<RolloverStep[]> {
    return [...this.steps.values()].sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
  }

  async getStepsByWorkstream(workstream: string): Promise<RolloverStep[]> {
    return [...this.steps.values()]
      .filter(s => s.workstream === workstream)
      .sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
  }

  async getStepsByStatus(status: string, track: 'Corp' | 'Fed' = 'Corp'): Promise<RolloverStep[]> {
    return [...this.steps.values()]
      .filter(s => (track === 'Corp' ? s.corpStatus : s.fedStatus) === status)
      .sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
  }

  async getStepsByOwner(ownerName: string): Promise<RolloverStep[]> {
    const lower = ownerName.toLowerCase();
    return [...this.steps.values()].filter(s =>
      s.wwicPoc.toLowerCase().includes(lower) ||
      s.fedPoc.toLowerCase().includes(lower) ||
      s.engineeringDri.toLowerCase().includes(lower) ||
      s.engineeringLead.toLowerCase().includes(lower)
    );
  }

  /** Search steps by keyword across description, engineeringDependent, owner, and workstream fields */
  async searchSteps(keyword: string): Promise<RolloverStep[]> {
    const lower = keyword.toLowerCase();
    return [...this.steps.values()].filter(s =>
      s.description.toLowerCase().includes(lower) ||
      s.engineeringDependent.toLowerCase().includes(lower) ||
      s.wwicPoc.toLowerCase().includes(lower) ||
      s.fedPoc.toLowerCase().includes(lower) ||
      s.engineeringDri.toLowerCase().includes(lower) ||
      s.engineeringLead.toLowerCase().includes(lower) ||
      s.workstream.toLowerCase().includes(lower) ||
      s.referenceNotes.toLowerCase().includes(lower)
    ).sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
  }

  async updateStepStatus(
    stepId: string,
    field: 'corpStatus' | 'fedStatus',
    newStatus: string,
    updatedBy: string,
    source: 'bot' | 'excel' | 'automation' | 'webhook'
  ): Promise<RolloverStep | null> {
    const step = this.steps.get(stepId);
    if (!step) return null;

    const previousValue = (step as any)[field];
    (step as any)[field] = newStatus;
    step.lastModified = new Date().toISOString();
    step.lastModifiedBy = updatedBy;
    step.lastModifiedSource = source;

    if (newStatus === 'Completed' && field === 'corpStatus') {
      step.corpCompletedDate = new Date().toISOString().split('T')[0];
    }

    this.steps.set(stepId, step);
    await this.logAudit(stepId, field, previousValue, newStatus, updatedBy, source);
    return step;
  }

  // ---- Milestones ----

  async upsertMilestone(milestone: Milestone): Promise<void> {
    this.milestones.set(milestone.id, { ...milestone });
  }

  async getAllMilestones(): Promise<Milestone[]> {
    return [...this.milestones.values()];
  }

  // ---- Audit ----

  async logAudit(
    stepId: string,
    field: string,
    previousValue: string,
    newValue: string,
    changedBy: string,
    source: 'bot' | 'excel' | 'automation' | 'webhook',
    reason?: string
  ): Promise<void> {
    this.audit.push({
      id: `audit-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      stepId,
      field,
      previousValue,
      newValue,
      changedBy,
      changedAt: new Date().toISOString(),
      source,
      reason,
    });
  }

  async getAuditForStep(stepId: string): Promise<AuditEntry[]> {
    return this.audit.filter(a => a.stepId === stepId).reverse();
  }

  async getRecentChanges(since: string, limit: number = 50): Promise<AuditEntry[]> {
    return this.audit
      .filter(a => a.changedAt >= since)
      .sort((a, b) => b.changedAt.localeCompare(a.changedAt))
      .slice(0, limit);
  }

  async getAllAudit(limit: number = 200): Promise<AuditEntry[]> {
    return this.audit
      .sort((a, b) => b.changedAt.localeCompare(a.changedAt))
      .slice(0, limit);
  }

  // ---- Users ----

  async upsertUser(user: UserProfile): Promise<void> {
    this.users.set(user.id, { ...user });
  }

  async getUserByEmail(email: string): Promise<UserProfile | null> {
    return [...this.users.values()].find(u => u.email === email) || null;
  }

  async getUsersByRole(role: string): Promise<UserProfile[]> {
    return [...this.users.values()].filter(u => u.role === role);
  }

  async getDRIsForStep(stepId: string): Promise<UserProfile[]> {
    return [...this.users.values()].filter(u => u.ownedSteps.includes(stepId));
  }

  // ---- Bulk ----

  async bulkUpsertSteps(steps: RolloverStep[]): Promise<number> {
    for (const step of steps) {
      this.steps.set(step.id, { ...step });
    }
    return steps.length;
  }

  async bulkUpsertMilestones(milestones: Milestone[]): Promise<number> {
    for (const ms of milestones) {
      this.milestones.set(ms.id, { ...ms });
    }
    return milestones.length;
  }

  // ---- Stats ----

  getStepCount(): number { return this.steps.size; }
  getAuditCount(): number { return this.audit.length; }
}
