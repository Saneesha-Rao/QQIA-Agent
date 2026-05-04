import { CosmosClient, Database, Container } from '@azure/cosmos';
import { config } from '../config/appConfig';
import { RolloverStep, Milestone, KeyBusinessDate, AuditEntry, UserProfile, RaidEntry } from '../models/types';

/**
 * Data access layer for Cosmos DB operations.
 * Manages CRUD for steps, milestones, audit logs, and user profiles.
 */
export class DataService {
  private client: CosmosClient;
  private db!: Database;
  private stepsContainer!: Container;
  private milestonesContainer!: Container;
  private auditContainer!: Container;
  private usersContainer!: Container;

  constructor() {
    this.client = new CosmosClient({
      endpoint: config.cosmos.endpoint,
      key: config.cosmos.key,
    });
  }

  /** Initialize database and containers */
  async initialize(): Promise<void> {
    const { database } = await this.client.databases.createIfNotExists({ id: config.cosmos.database });
    this.db = database;

    const containers = [
      { id: config.cosmos.containers.steps, partitionKey: '/workstream' },
      { id: config.cosmos.containers.milestones, partitionKey: '/category' },
      { id: config.cosmos.containers.audit, partitionKey: '/stepId' },
      { id: config.cosmos.containers.users, partitionKey: '/role' },
    ];

    for (const c of containers) {
      await this.db.containers.createIfNotExists({ id: c.id, partitionKey: { paths: [c.partitionKey] } });
    }

    this.stepsContainer = this.db.container(config.cosmos.containers.steps);
    this.milestonesContainer = this.db.container(config.cosmos.containers.milestones);
    this.auditContainer = this.db.container(config.cosmos.containers.audit);
    this.usersContainer = this.db.container(config.cosmos.containers.users);
  }

  // ---- Steps CRUD ----

  async upsertStep(step: RolloverStep): Promise<void> {
    await this.stepsContainer.items.upsert(step);
  }

  async getStep(stepId: string): Promise<RolloverStep | null> {
    const query = `SELECT * FROM c WHERE c.id = @id`;
    const { resources } = await this.stepsContainer.items
      .query({ query, parameters: [{ name: '@id', value: stepId }] })
      .fetchAll();
    return resources[0] || null;
  }

  async getAllSteps(): Promise<RolloverStep[]> {
    const { resources } = await this.stepsContainer.items
      .query('SELECT * FROM c ORDER BY c.id')
      .fetchAll();
    return resources;
  }

  async getStepsByWorkstream(workstream: string): Promise<RolloverStep[]> {
    const query = 'SELECT * FROM c WHERE c.workstream = @ws ORDER BY c.id';
    const { resources } = await this.stepsContainer.items
      .query({ query, parameters: [{ name: '@ws', value: workstream }] })
      .fetchAll();
    return resources;
  }

  async getStepsByStatus(status: string, track: 'Corp' | 'Fed' = 'Corp'): Promise<RolloverStep[]> {
    const field = track === 'Corp' ? 'corpStatus' : 'fedStatus';
    const query = `SELECT * FROM c WHERE c.${field} = @status ORDER BY c.id`;
    const { resources } = await this.stepsContainer.items
      .query({ query, parameters: [{ name: '@status', value: status }] })
      .fetchAll();
    return resources;
  }

  async getStepsByOwner(ownerName: string): Promise<RolloverStep[]> {
    const query = `SELECT * FROM c WHERE CONTAINS(c.wwicPoc, @name) OR CONTAINS(c.fedPoc, @name) OR CONTAINS(c.engineeringDri, @name) OR CONTAINS(c.engineeringLead, @name)`;
    const { resources } = await this.stepsContainer.items
      .query({ query, parameters: [{ name: '@name', value: ownerName }] })
      .fetchAll();
    return resources;
  }

  /** Search steps by keyword across description, engineeringDependent, owner, and workstream fields */
  async searchSteps(keyword: string): Promise<RolloverStep[]> {
    const query = `SELECT * FROM c WHERE CONTAINS(LOWER(c.description), @kw) OR CONTAINS(LOWER(c.engineeringDependent), @kw) OR CONTAINS(LOWER(c.wwicPoc), @kw) OR CONTAINS(LOWER(c.fedPoc), @kw) OR CONTAINS(LOWER(c.engineeringDri), @kw) OR CONTAINS(LOWER(c.engineeringLead), @kw) OR CONTAINS(LOWER(c.workstream), @kw) OR CONTAINS(LOWER(c.referenceNotes), @kw) ORDER BY c.id`;
    const { resources } = await this.stepsContainer.items
      .query({ query, parameters: [{ name: '@kw', value: keyword.toLowerCase() }] })
      .fetchAll();
    return resources;
  }

  async updateStepStatus(
    stepId: string,
    field: 'corpStatus' | 'fedStatus',
    newStatus: string,
    updatedBy: string,
    source: 'bot' | 'excel' | 'automation' | 'webhook'
  ): Promise<RolloverStep | null> {
    const step = await this.getStep(stepId);
    if (!step) return null;

    const previousValue = (step as any)[field];
    (step as any)[field] = newStatus;
    step.lastModified = new Date().toISOString();
    step.lastModifiedBy = updatedBy;
    step.lastModifiedSource = source;

    if (newStatus === 'Completed') {
      const completedField = field === 'corpStatus' ? 'corpCompletedDate' : 'fedStartDate';
      if (field === 'corpStatus') step.corpCompletedDate = new Date().toISOString().split('T')[0];
    }

    await this.upsertStep(step);
    await this.logAudit(stepId, field, previousValue, newStatus, updatedBy, source);
    return step;
  }

  // ---- Milestones CRUD ----

  async upsertMilestone(milestone: Milestone): Promise<void> {
    await this.milestonesContainer.items.upsert(milestone);
  }

  async getAllMilestones(): Promise<Milestone[]> {
    const { resources } = await this.milestonesContainer.items
      .query('SELECT * FROM c ORDER BY c.id')
      .fetchAll();
    return resources;
  }

  // ---- Audit Log ----

  async logAudit(
    stepId: string,
    field: string,
    previousValue: string,
    newValue: string,
    changedBy: string,
    source: 'bot' | 'excel' | 'automation' | 'webhook',
    reason?: string
  ): Promise<void> {
    const entry: AuditEntry = {
      id: `audit-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      stepId,
      field,
      previousValue,
      newValue,
      changedBy,
      changedAt: new Date().toISOString(),
      source,
      reason,
    };
    await this.auditContainer.items.create(entry);
  }

  async getAuditForStep(stepId: string): Promise<AuditEntry[]> {
    const query = 'SELECT * FROM c WHERE c.stepId = @stepId ORDER BY c.changedAt DESC';
    const { resources } = await this.auditContainer.items
      .query({ query, parameters: [{ name: '@stepId', value: stepId }] })
      .fetchAll();
    return resources;
  }

  async getRecentChanges(since: string, limit: number = 50): Promise<AuditEntry[]> {
    const query = 'SELECT TOP @limit * FROM c WHERE c.changedAt >= @since ORDER BY c.changedAt DESC';
    const { resources } = await this.auditContainer.items
      .query({ query, parameters: [{ name: '@since', value: since }, { name: '@limit', value: limit }] })
      .fetchAll();
    return resources;
  }

  async getAllAudit(limit: number = 200): Promise<AuditEntry[]> {
    const query = 'SELECT TOP @limit * FROM c ORDER BY c.changedAt DESC';
    const { resources } = await this.auditContainer.items
      .query({ query, parameters: [{ name: '@limit', value: limit }] })
      .fetchAll();
    return resources;
  }

  // ---- Users ----

  async upsertUser(user: UserProfile): Promise<void> {
    await this.usersContainer.items.upsert(user);
  }

  async getUserByEmail(email: string): Promise<UserProfile | null> {
    const query = 'SELECT * FROM c WHERE c.email = @email';
    const { resources } = await this.usersContainer.items
      .query({ query, parameters: [{ name: '@email', value: email }] })
      .fetchAll();
    return resources[0] || null;
  }

  async getUsersByRole(role: string): Promise<UserProfile[]> {
    const query = 'SELECT * FROM c WHERE c.role = @role';
    const { resources } = await this.usersContainer.items
      .query({ query, parameters: [{ name: '@role', value: role }] })
      .fetchAll();
    return resources;
  }

  async getDRIsForStep(stepId: string): Promise<UserProfile[]> {
    const query = 'SELECT * FROM c WHERE ARRAY_CONTAINS(c.ownedSteps, @stepId)';
    const { resources } = await this.usersContainer.items
      .query({ query, parameters: [{ name: '@stepId', value: stepId }] })
      .fetchAll();
    return resources;
  }

  // ---- Bulk Operations ----

  async bulkUpsertSteps(steps: RolloverStep[]): Promise<number> {
    let count = 0;
    for (const step of steps) {
      await this.upsertStep(step);
      count++;
    }
    return count;
  }

  async bulkUpsertMilestones(milestones: Milestone[]): Promise<number> {
    let count = 0;
    for (const ms of milestones) {
      await this.upsertMilestone(ms);
      count++;
    }
    return count;
  }
}
