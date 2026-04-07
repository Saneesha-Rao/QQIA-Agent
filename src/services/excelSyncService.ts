import * as XLSX from 'xlsx';
import { config } from '../config/appConfig';
import { RolloverStep, Milestone, KeyBusinessDate, RaidEntry } from '../models/types';
import { DataService } from './dataService';
import {
  parseRolloverRow,
  parseMilestoneRow,
  parseKeyBusinessDateRow,
  parseRaidRow,
  isoToExcelDate,
} from '../utils/excelParser';

/**
 * Bi-directional sync service between the Excel file and Cosmos DB.
 * 
 * For local development: reads/writes Excel file directly.
 * For production: uses Microsoft Graph API to access SharePoint/OneDrive.
 */
export class ExcelSyncService {
  private dataService: DataService;
  private localFilePath: string;

  constructor(dataService: DataService, localFilePath?: string) {
    this.dataService = dataService;
    this.localFilePath = localFilePath ||
      'C:\\Users\\salingal\\OneDrive - Microsoft\\Seller Incentives\\QQIA\\FY27_Mint_RolloverTimeline.xlsx';
  }

  /** Import all data from Excel into Cosmos DB (initial load) */
  async importFromExcel(): Promise<ImportResult> {
    const wb = XLSX.readFile(this.localFilePath);
    const result: ImportResult = { steps: 0, milestones: 0, keyDates: 0, raidEntries: 0 };

    // Import FY27_Rollover steps
    const rolloverSheet = wb.Sheets['FY27_Rollover'];
    if (rolloverSheet) {
      const rows: any[][] = XLSX.utils.sheet_to_json(rolloverSheet, { header: 1, defval: '' });
      for (let i = 3; i < rows.length; i++) { // Skip header rows (0-2)
        const step = parseRolloverRow(rows[i], i);
        if (step) {
          await this.dataService.upsertStep(step);
          result.steps++;
        }
      }
    }

    // Import HighLevelMilestones
    const msSheet = wb.Sheets['HighLevelMilestones'];
    if (msSheet) {
      const rows: any[][] = XLSX.utils.sheet_to_json(msSheet, { header: 1, defval: '' });
      for (let i = 2; i < rows.length; i++) {
        const ms = parseMilestoneRow(rows[i], i);
        if (ms) {
          await this.dataService.upsertMilestone(ms);
          result.milestones++;
        }
      }
    }

    console.log(`Import complete: ${result.steps} steps, ${result.milestones} milestones`);
    return result;
  }

  /** Sync changes from Cosmos DB back to Excel (write-back) */
  async syncToExcel(): Promise<number> {
    const wb = XLSX.readFile(this.localFilePath);
    const ws = wb.Sheets['FY27_Rollover'];
    if (!ws) return 0;

    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const steps = await this.dataService.getAllSteps();
    const stepMap = new Map(steps.map(s => [s.id, s]));
    let updatedCount = 0;

    for (let i = 3; i < rows.length; i++) {
      const stepId = rows[i][0]?.toString().trim();
      if (!stepId) continue;

      const dbStep = stepMap.get(stepId);
      if (!dbStep) continue;

      // Only write back if the bot/automation changed it more recently
      if (dbStep.lastModifiedSource === 'bot' || dbStep.lastModifiedSource === 'automation') {
        let changed = false;

        // Update status columns
        if (rows[i][5] !== dbStep.corpStatus) { rows[i][5] = dbStep.corpStatus; changed = true; }
        if (rows[i][9] !== dbStep.fedStatus) { rows[i][9] = dbStep.fedStatus; changed = true; }

        // Update completed date
        if (dbStep.corpCompletedDate) {
          const excelDate = isoToExcelDate(dbStep.corpCompletedDate);
          if (rows[i][6] !== excelDate) { rows[i][6] = excelDate; changed = true; }
        }

        // Update reference notes
        if (dbStep.referenceNotes && rows[i][18] !== dbStep.referenceNotes) {
          rows[i][18] = dbStep.referenceNotes;
          changed = true;
        }

        if (changed) updatedCount++;
      }
    }

    if (updatedCount > 0) {
      const newWs = XLSX.utils.aoa_to_sheet(rows);
      wb.Sheets['FY27_Rollover'] = newWs;
      XLSX.writeFile(wb, this.localFilePath);
      console.log(`Synced ${updatedCount} step(s) back to Excel`);
    }

    return updatedCount;
  }

  /** Sync changes from Excel into Cosmos DB (pick up manual Excel edits) */
  async syncFromExcel(): Promise<number> {
    const wb = XLSX.readFile(this.localFilePath);
    const ws = wb.Sheets['FY27_Rollover'];
    if (!ws) return 0;

    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    let updatedCount = 0;

    for (let i = 3; i < rows.length; i++) {
      const excelStep = parseRolloverRow(rows[i], i);
      if (!excelStep) continue;

      const dbStep = await this.dataService.getStep(excelStep.id);
      if (!dbStep) {
        // New step in Excel, import it
        await this.dataService.upsertStep(excelStep);
        updatedCount++;
        continue;
      }

      // Compare and update if Excel has newer data
      let changed = false;
      const fieldsToCheck: (keyof RolloverStep)[] = [
        'corpStatus', 'fedStatus', 'corpCompletedDate', 'referenceNotes',
        'corpStartDate', 'corpEndDate', 'fedStartDate', 'fedEndDate',
      ];

      for (const field of fieldsToCheck) {
        if (excelStep[field] !== dbStep[field]) {
          // Excel changed - if DB was last modified by Excel or is older, accept the change
          if (dbStep.lastModifiedSource === 'excel' || !dbStep.lastModified) {
            (dbStep as any)[field] = excelStep[field];
            changed = true;
            await this.dataService.logAudit(
              dbStep.id, field, (dbStep as any)[field]?.toString() || '',
              (excelStep as any)[field]?.toString() || '', 'excel-sync', 'excel'
            );
          }
        }
      }

      if (changed) {
        dbStep.lastModified = new Date().toISOString();
        dbStep.lastModifiedBy = 'excel-sync';
        dbStep.lastModifiedSource = 'excel';
        await this.dataService.upsertStep(dbStep);
        updatedCount++;
      }
    }

    return updatedCount;
  }

  /** Full bi-directional sync: Excel → DB, then DB → Excel */
  async fullSync(): Promise<{ fromExcel: number; toExcel: number }> {
    const fromExcel = await this.syncFromExcel();
    const toExcel = await this.syncToExcel();
    console.log(`Full sync: ${fromExcel} from Excel, ${toExcel} to Excel`);
    return { fromExcel, toExcel };
  }

  /** Get summary of current Excel data without importing */
  readExcelSummary(): ExcelSummary {
    const wb = XLSX.readFile(this.localFilePath);
    const summary: ExcelSummary = { sheets: [], totalSteps: 0 };

    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      summary.sheets.push({ name, rowCount: rows.length });
    }

    const rolloverSheet = wb.Sheets['FY27_Rollover'];
    if (rolloverSheet) {
      const rows: any[][] = XLSX.utils.sheet_to_json(rolloverSheet, { header: 1, defval: '' });
      for (let i = 3; i < rows.length; i++) {
        if (rows[i][0]?.toString().trim()) summary.totalSteps++;
      }
    }

    return summary;
  }
}

interface ImportResult {
  steps: number;
  milestones: number;
  keyDates: number;
  raidEntries: number;
}

interface ExcelSummary {
  sheets: { name: string; rowCount: number }[];
  totalSteps: number;
}
