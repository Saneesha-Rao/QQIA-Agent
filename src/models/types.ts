import { StepStatus } from '../config/appConfig';

/** Core rollover step from FY27_Rollover sheet */
export interface RolloverStep {
  id: string;               // e.g. "1.A", "2.B"
  workstream: string;       // Grouping column (e.g. "System Rollover")
  description: string;
  corpStartDate: string | null;
  corpEndDate: string | null;
  corpStatus: StepStatus;
  corpCompletedDate: string | null;
  fedStartDate: string | null;
  fedEndDate: string | null;
  fedStatus: StepStatus;
  setupValidation: string;
  engineeringDependent: string;
  wwicPoc: string;
  fedPoc: string;
  engineeringDri: string;
  engineeringLead: string;
  dependencies: string[];   // Array of step IDs from StepDependent
  adoLink: string;
  referenceNotes: string;
  fy26CorpStart: string | null;
  fy26CorpEnd: string | null;
  fy26FedStart: string | null;
  fy26FedEnd: string | null;
  lastModified: string;
  lastModifiedBy: string;
  lastModifiedSource: 'bot' | 'excel' | 'automation' | 'webhook' | 'com_synced';
}

/** High-level milestone from HighLevelMilestones sheet */
export interface Milestone {
  id: string;
  category: string;
  milestone: string;
  corpDate: string | null;
  corpStatus: StepStatus;
  fedDate: string | null;
  fedStatus: StepStatus;
  comments: string;
  fy25CorpDate: string | null;
  fy25FedDate: string | null;
}

/** Key business date from KeyBusinessDates sheet */
export interface KeyBusinessDate {
  id: string;
  category: string;
  driTeam: string;
  milestone: string;
  owner: string;
  startDate: string | null;
  endDate: string | null;
  processTime: string;
  status: StepStatus;
  timelineLock: string;
  fy25StartDate: string | null;
  fy25EndDate: string | null;
  fy25ProcessTime: string;
}

/** RAID log entry */
export interface RaidEntry {
  id: string;
  date: string;
  description: string;
  mitigation: string;
  nextSteps: string;
  owner: string;
  dueDate: string | null;
}

/** Audit log entry for tracking all changes */
export interface AuditEntry {
  id: string;
  stepId: string;
  field: string;
  previousValue: string;
  newValue: string;
  changedBy: string;
  changedAt: string;
  source: 'bot' | 'excel' | 'automation' | 'webhook';
  reason?: string;
}

/** User role mapping */
export interface UserProfile {
  id: string;              // Azure AD object ID
  email: string;
  displayName: string;
  role: 'dri' | 'pm' | 'leadership' | 'admin';
  ownedSteps: string[];    // Step IDs this user owns
  teamsUserId: string;     // For proactive messaging
  notificationPreferences: {
    deadlineAlerts: boolean;
    weeklyDigest: boolean;
    blockerAlerts: boolean;
    predecessorAlerts: boolean;
  };
}
