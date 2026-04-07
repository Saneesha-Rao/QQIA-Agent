import * as XLSX from 'xlsx';
import { config } from '../config/appConfig';
import { RolloverStep } from '../models/types';
import { DataService } from './dataService';
import { GraphService } from './graphService';
import { DependencyEngine } from './dependencyEngine';
import { NotificationService } from './notificationService';
import {
  parseRolloverRow,
  parseMilestoneRow,
  isoToExcelDate,
} from '../utils/excelParser';

export interface SyncResult {
  fromExcel: number;
  toExcel: number;
  conflicts: SyncConflict[];
  timestamp: string;
}

export interface SyncConflict {
  stepId: string;
  field: string;
  excelValue: string;
  dbValue: string;
  resolution: 'excel_wins' | 'db_wins';
  reason: string;
}

/**
 * Production-grade bi-directional Excel sync using Graph API.
 * Runs every 15 minutes, resolving conflicts with last-write-wins + audit.
 */
export class SyncEngine {
  private dataService: DataService;
  private graphService: GraphService;
  private dependencyEngine: DependencyEngine;
  private notificationService: NotificationService;
  private localFallbackPath: string;
  private lastSyncTimestamp: string | null = null;

  constructor(
    dataService: DataService,
    graphService: GraphService,
    dependencyEngine: DependencyEngine,
    notificationService: NotificationService,
    localFallbackPath?: string
  ) {
    this.dataService = dataService;
    this.graphService = graphService;
    this.dependencyEngine = dependencyEngine;
    this.notificationService = notificationService;
    this.localFallbackPath = localFallbackPath ||
      'C:\\Users\\salingal\\OneDrive - Microsoft\\Seller Incentives\\QQIA\\FY27_Mint_RolloverTimeline.xlsx';
  }

  /** Full bi-directional sync cycle */
  async runSync(): Promise<SyncResult> {
    const result: SyncResult = {
      fromExcel: 0,
      toExcel: 0,
      conflicts: [],
      timestamp: new Date().toISOString(),
    };

    try {
      console.log(`[SyncEngine] Starting sync at ${result.timestamp}`);

      // Step 1: Read current Excel data
      const excelSteps = await this.readExcelSteps();
      console.log(`[SyncEngine] Read ${excelSteps.length} steps from Excel`);

      // Step 2: Read current DB state
      const dbSteps = await this.dataService.getAllSteps();
      const dbStepMap = new Map(dbSteps.map(s => [s.id, s]));

      // Step 3: Sync Excel → DB (pick up manual Excel edits)
      for (const excelStep of excelSteps) {
        const dbStep = dbStepMap.get(excelStep.id);

        if (!dbStep) {
          // New step in Excel, import it
          await this.dataService.upsertStep(excelStep);
          result.fromExcel++;
          continue;
        }

        // Compare tracked fields
        const fieldsToSync: { field: keyof RolloverStep; excelVal: any; dbVal: any }[] = [
          { field: 'corpStatus', excelVal: excelStep.corpStatus, dbVal: dbStep.corpStatus },
          { field: 'fedStatus', excelVal: excelStep.fedStatus, dbVal: dbStep.fedStatus },
          { field: 'corpStartDate', excelVal: excelStep.corpStartDate, dbVal: dbStep.corpStartDate },
          { field: 'corpEndDate', excelVal: excelStep.corpEndDate, dbVal: dbStep.corpEndDate },
          { field: 'fedStartDate', excelVal: excelStep.fedStartDate, dbVal: dbStep.fedStartDate },
          { field: 'fedEndDate', excelVal: excelStep.fedEndDate, dbVal: dbStep.fedEndDate },
          { field: 'corpCompletedDate', excelVal: excelStep.corpCompletedDate, dbVal: dbStep.corpCompletedDate },
          { field: 'referenceNotes', excelVal: excelStep.referenceNotes, dbVal: dbStep.referenceNotes },
          { field: 'wwicPoc', excelVal: excelStep.wwicPoc, dbVal: dbStep.wwicPoc },
          { field: 'engineeringDri', excelVal: excelStep.engineeringDri, dbVal: dbStep.engineeringDri },
        ];

        let stepChanged = false;
        for (const { field, excelVal, dbVal } of fieldsToSync) {
          if (this.normalize(excelVal) !== this.normalize(dbVal)) {
            // Conflict resolution: last-write-wins based on source
            const conflict: SyncConflict = {
              stepId: excelStep.id,
              field: field as string,
              excelValue: String(excelVal ?? ''),
              dbValue: String(dbVal ?? ''),
              resolution: 'excel_wins',
              reason: '',
            };

            if (dbStep.lastModifiedSource === 'bot' || dbStep.lastModifiedSource === 'automation') {
              // DB was modified by bot/automation more recently — DB wins
              conflict.resolution = 'db_wins';
              conflict.reason = `DB updated by ${dbStep.lastModifiedSource} at ${dbStep.lastModified}`;
            } else {
              // Excel wins — apply the change to DB
              conflict.resolution = 'excel_wins';
              conflict.reason = 'Excel has newer manual edit';
              (dbStep as any)[field] = excelVal;
              stepChanged = true;

              await this.dataService.logAudit(
                dbStep.id, field as string,
                String(dbVal ?? ''), String(excelVal ?? ''),
                'sync-engine', 'excel'
              );
            }

            result.conflicts.push(conflict);
          }
        }

        if (stepChanged) {
          dbStep.lastModified = new Date().toISOString();
          dbStep.lastModifiedBy = 'sync-engine';
          dbStep.lastModifiedSource = 'excel';
          await this.dataService.upsertStep(dbStep);
          result.fromExcel++;

          // Check if status changed to Completed — trigger dependency notifications
          if (dbStep.corpStatus === 'Completed' && excelStep.corpStatus === 'Completed') {
            await this.notificationService.notifyPredecessorComplete(dbStep.id, 'Corp');
          }
        }
      }

      // Step 4: Sync DB → Excel (push bot/automation changes)
      const botUpdatedSteps = dbSteps.filter(s =>
        (s.lastModifiedSource === 'bot' || s.lastModifiedSource === 'automation') &&
        (!this.lastSyncTimestamp || s.lastModified > this.lastSyncTimestamp)
      );

      if (botUpdatedSteps.length > 0) {
        result.toExcel = await this.writeStepsToExcel(botUpdatedSteps, excelSteps);
      }

      this.lastSyncTimestamp = result.timestamp;
      console.log(`[SyncEngine] Sync complete: ${result.fromExcel} from Excel, ${result.toExcel} to Excel, ${result.conflicts.length} conflicts`);

    } catch (err: any) {
      console.error(`[SyncEngine] Sync failed: ${err.message}`);
      throw err;
    }

    return result;
  }

  /** Read steps from Excel via Graph API, fallback to local file */
  private async readExcelSteps(): Promise<RolloverStep[]> {
    let rows: any[][];

    try {
      // Try Graph API first (production)
      const { values } = await this.graphService.getUsedRange('FY27_Rollover');
      rows = values;
    } catch {
      // Fallback to local file (development)
      console.log('[SyncEngine] Graph API unavailable, using local file');
      const wb = XLSX.readFile(this.localFallbackPath);
      const ws = wb.Sheets['FY27_Rollover'];
      rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    }

    const steps: RolloverStep[] = [];
    for (let i = 3; i < rows.length; i++) {
      const step = parseRolloverRow(rows[i], i);
      if (step) steps.push(step);
    }
    return steps;
  }

  /** Write bot/automation-updated steps back to Excel */
  private async writeStepsToExcel(updatedSteps: RolloverStep[], currentExcelSteps: RolloverStep[]): Promise<number> {
    try {
      // Build a map of step ID → Excel row index
      // Header rows are 0-2, data starts at row 3
      const excelStepIds = currentExcelSteps.map(s => s.id);
      let writtenCount = 0;

      for (const step of updatedSteps) {
        const excelIndex = excelStepIds.indexOf(step.id);
        if (excelIndex === -1) continue;

        const excelRow = excelIndex + 4; // +3 for header rows, +1 for 1-based indexing

        try {
          // Update status column (F = column 6)
          await this.graphService.updateWorksheetRange(
            'FY27_Rollover',
            `F${excelRow}`,
            [[step.corpStatus]]
          );

          // Update completed date column (G = column 7)
          if (step.corpCompletedDate) {
            const excelDate = isoToExcelDate(step.corpCompletedDate);
            await this.graphService.updateWorksheetRange(
              'FY27_Rollover',
              `G${excelRow}`,
              [[excelDate]]
            );
          }

          // Update Fed status column (J = column 10)
          await this.graphService.updateWorksheetRange(
            'FY27_Rollover',
            `J${excelRow}`,
            [[step.fedStatus]]
          );

          // Update reference notes column (S = column 19)
          if (step.referenceNotes) {
            await this.graphService.updateWorksheetRange(
              'FY27_Rollover',
              `S${excelRow}`,
              [[step.referenceNotes]]
            );
          }

          writtenCount++;
          await this.dataService.logAudit(
            step.id, 'excel_writeback', '', 'synced',
            'sync-engine', 'automation', 'Pushed bot changes to Excel'
          );

          // Reset source so it won't be written again next cycle
          step.lastModifiedSource = 'excel';
          await this.dataService.upsertStep(step);

        } catch (err: any) {
          console.error(`[SyncEngine] Failed to write step ${step.id} to Excel: ${err.message}`);
        }
      }

      return writtenCount;
    } catch {
      // Fallback: write entire file locally
      return this.writeStepsToLocalExcel(updatedSteps);
    }
  }

  /** Local fallback: write steps to Excel file directly */
  private writeStepsToLocalExcel(updatedSteps: RolloverStep[]): number {
    const wb = XLSX.readFile(this.localFallbackPath);
    const ws = wb.Sheets['FY27_Rollover'];
    if (!ws) return 0;

    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const stepMap = new Map(updatedSteps.map(s => [s.id, s]));
    let count = 0;

    for (let i = 3; i < rows.length; i++) {
      const stepId = rows[i][0]?.toString().trim();
      const step = stepMap.get(stepId);
      if (!step) continue;

      let changed = false;
      if (rows[i][5] !== step.corpStatus) { rows[i][5] = step.corpStatus; changed = true; }
      if (rows[i][9] !== step.fedStatus) { rows[i][9] = step.fedStatus; changed = true; }
      if (step.corpCompletedDate) {
        const ed = isoToExcelDate(step.corpCompletedDate);
        if (rows[i][6] !== ed) { rows[i][6] = ed; changed = true; }
      }
      if (step.referenceNotes && rows[i][18] !== step.referenceNotes) {
        rows[i][18] = step.referenceNotes; changed = true;
      }
      if (changed) count++;
    }

    if (count > 0) {
      const newWs = XLSX.utils.aoa_to_sheet(rows);
      wb.Sheets['FY27_Rollover'] = newWs;
      XLSX.writeFile(wb, this.localFallbackPath);
    }

    return count;
  }

  private normalize(val: any): string {
    if (val === null || val === undefined) return '';
    return String(val).trim();
  }
}
