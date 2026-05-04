import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';
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
  private sourcePath: string;   // OneDrive source (may be locked)
  private localWorkingPath: string;  // Local working copy in data/
  private lastSyncTimestamp: string | null = null;
  private hasPendingLocalWrites: boolean = false;  // True when local file has unsynced changes

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
    this.sourcePath = localFallbackPath ||
      'C:\\Users\\salingal\\OneDrive - Microsoft\\Seller Incentives\\QQIA\\FY27_Mint_RolloverTimeline.xlsx';
    this.localWorkingPath = path.join(process.cwd(), 'data', 'FY27_Mint_RolloverTimeline.xlsx');
  }

  /** Refresh local working copy from OneDrive source if newer */
  private refreshLocalCopy(): void {
    try {
      // Don't overwrite local file if we have pending writes that haven't been pushed
      if (this.hasPendingLocalWrites) {
        // Try to push pending changes first
        if (this.pushToSource()) {
          this.hasPendingLocalWrites = false;
          console.log('[SyncEngine] Pushed pending local changes to OneDrive before refresh');
        } else {
          console.log('[SyncEngine] Skipping refresh — local file has pending writes that could not be pushed');
          return;
        }
      }

      if (!fs.existsSync(this.sourcePath)) return;
      const srcStat = fs.statSync(this.sourcePath);
      const localExists = fs.existsSync(this.localWorkingPath);
      const localStat = localExists ? fs.statSync(this.localWorkingPath) : null;
      if (!localExists || srcStat.mtimeMs > (localStat?.mtimeMs || 0)) {
        fs.copyFileSync(this.sourcePath, this.localWorkingPath);
        console.log('[SyncEngine] Refreshed local copy from OneDrive source');
      }
    } catch (err: any) {
      console.warn(`[SyncEngine] Could not refresh local copy: ${err.message}`);
    }
  }

  /** Push local working copy back to OneDrive source. Returns true if successful. */
  private pushToSource(): boolean {
    try {
      fs.copyFileSync(this.localWorkingPath, this.sourcePath);
      console.log('[SyncEngine] Pushed local copy back to OneDrive source');
      return true;
    } catch (err: any) {
      console.warn(`[SyncEngine] Could not push to OneDrive (file may be open): ${err.message}`);
      return false;
    }
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

            if (dbStep.lastModifiedSource === 'bot' || dbStep.lastModifiedSource === 'automation' || dbStep.lastModifiedSource === 'webhook' || dbStep.lastModifiedSource === 'com_synced') {
              // DB was modified by bot/automation/webhook/COM more recently — DB wins
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
      // Include steps modified by bot AND steps where db won a conflict
      const dbWinsStepIds = new Set(
        result.conflicts.filter(c => c.resolution === 'db_wins').map(c => c.stepId)
      );
      const botUpdatedSteps = dbSteps.filter(s =>
        dbWinsStepIds.has(s.id) ||
        (
          (s.lastModifiedSource === 'bot' || s.lastModifiedSource === 'automation' || s.lastModifiedSource === 'webhook') &&
          (!this.lastSyncTimestamp || s.lastModified > this.lastSyncTimestamp)
        )
      );

      // Reset com_synced steps to 'excel' now that SyncEngine has seen them
      // (COM already wrote to Excel, no need to push again)
      for (const s of dbSteps) {
        if (s.lastModifiedSource === 'com_synced') {
          s.lastModifiedSource = 'excel';
          await this.dataService.upsertStep(s);
        }
      }

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
      // Refresh local copy from OneDrive source first (picks up manual edits)
      this.refreshLocalCopy();
      console.log('[SyncEngine] Graph API unavailable, using local file');
      const wb = XLSX.readFile(this.localWorkingPath);
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
    // Try Graph API first
    let graphAvailable = true;
    try {
      const excelStepIds = currentExcelSteps.map(s => s.id);
      let writtenCount = 0;

      for (const step of updatedSteps) {
        const excelIndex = excelStepIds.indexOf(step.id);
        if (excelIndex === -1) continue;

        const excelRow = excelIndex + 4; // +3 for header rows, +1 for 1-based indexing

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
        step.lastModifiedSource = 'com_synced';
        await this.dataService.upsertStep(step);
      }

      return writtenCount;
    } catch (err: any) {
      console.log(`[SyncEngine] Graph API write failed, falling back to local file: ${err.message}`);
      graphAvailable = false;
    }

    // Fallback: write to local Excel file
    const count = await this.writeStepsToLocalExcel(updatedSteps);

    // Try to push updated local file back to OneDrive source
    if (count > 0) {
      this.hasPendingLocalWrites = true;
      if (this.pushToSource()) {
        this.hasPendingLocalWrites = false;
      }
    }

    // Reset source for written steps — use 'com_synced' to protect
    // them from being reverted by the next sync if the file is stale
    for (const step of updatedSteps) {
      step.lastModifiedSource = 'com_synced';
      await this.dataService.upsertStep(step);
    }

    return count;
  }

  /** Local fallback: write steps to Excel file using ExcelJS (preserves formatting) */
  private async writeStepsToLocalExcel(updatedSteps: RolloverStep[]): Promise<number> {
    console.log(`[SyncEngine] writeStepsToLocalExcel: ${updatedSteps.length} steps to write`);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(this.localWorkingPath);
    const ws = wb.getWorksheet('FY27_Rollover');
    if (!ws) { console.log('[SyncEngine] Sheet FY27_Rollover not found!'); return 0; }

    // Build map of step IDs to their ExcelJS row numbers (1-based)
    // Data starts at row 4 (rows 1-3 are headers)
    const stepMap = new Map(updatedSteps.map(s => [s.id, s]));
    let count = 0;

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber <= 3) return; // skip header rows
      const stepId = row.getCell(1).text?.trim(); // Column A
      const step = stepMap.get(stepId);
      if (!step) return;

      let changed = false;

      // Column F (6) = corpStatus
      const currentStatus = row.getCell(6).text?.trim() || '';
      if (this.normalize(currentStatus) !== this.normalize(step.corpStatus)) {
        row.getCell(6).value = step.corpStatus;
        console.log(`[SyncEngine] ${stepId}: corpStatus "${currentStatus}" → "${step.corpStatus}"`);
        changed = true;
      }

      // Column J (10) = fedStatus
      const currentFedStatus = row.getCell(10).text?.trim() || '';
      if (this.normalize(currentFedStatus) !== this.normalize(step.fedStatus)) {
        row.getCell(10).value = step.fedStatus;
        console.log(`[SyncEngine] ${stepId}: fedStatus "${currentFedStatus}" → "${step.fedStatus}"`);
        changed = true;
      }

      // Column G (7) = corpCompletedDate
      if (step.corpCompletedDate) {
        const excelDate = isoToExcelDate(step.corpCompletedDate);
        const currentVal = row.getCell(7).value;
        if (currentVal !== excelDate) {
          row.getCell(7).value = excelDate;
          console.log(`[SyncEngine] ${stepId}: completedDate "${currentVal}" → "${excelDate}"`);
          changed = true;
        }
      }

      // Column S (19) = referenceNotes
      if (step.referenceNotes) {
        const currentNotes = row.getCell(19).text?.trim() || '';
        if (this.normalize(currentNotes) !== this.normalize(step.referenceNotes)) {
          row.getCell(19).value = step.referenceNotes;
          changed = true;
        }
      }

      if (changed) count++;
    });

    console.log(`[SyncEngine] Writing ${count} changed steps to ${this.localWorkingPath}`);
    if (count > 0) {
      await wb.xlsx.writeFile(this.localWorkingPath);
      console.log(`[SyncEngine] File written successfully`);
    }

    return count;
  }

  private normalize(val: any): string {
    if (val === null || val === undefined) return '';
    return String(val).trim();
  }
}
