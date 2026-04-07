import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';
import { config } from '../config/appConfig';
import { RolloverStep, Milestone, KeyBusinessDate, RaidEntry } from '../models/types';
import { DataService } from './dataService';
import { GraphService } from './graphService';
import {
  parseRolloverRow,
  parseMilestoneRow,
  parseKeyBusinessDateRow,
  parseRaidRow,
  isoToExcelDate,
} from '../utils/excelParser';

/**
 * Bi-directional sync service between the Excel file and the data store.
 * 
 * Mode 1 (Graph API — production):
 *   Reads/writes via Microsoft Graph Excel REST API.
 *   No file locks, works concurrently with other users.
 * 
 * Mode 2 (Local file — development):
 *   Uses a local working copy under qqia-agent/data/ to avoid OneDrive locks.
 *   Periodically attempts to push changes back to the OneDrive source.
 */
export class ExcelSyncService {
  private dataService: DataService;
  private graphService: GraphService | null = null;
  private useGraphApi: boolean = false;
  /** The OneDrive/SharePoint source file (used for local fallback) */
  private sourceFilePath: string;
  /** Local writable working copy (under qqia-agent/data/) */
  private localFilePath: string;

  constructor(dataService: DataService, graphService?: GraphService, sourceFilePath?: string) {
    this.dataService = dataService;
    this.sourceFilePath = sourceFilePath ||
      'C:\\Users\\salingal\\OneDrive - Microsoft\\Seller Incentives\\QQIA\\FY27_Mint_RolloverTimeline.xlsx';

    // Create local working copy path under project data/ directory
    const projectRoot = path.resolve(__dirname, '..', '..');
    const dataDir = path.join(projectRoot, 'data');
    if (!fs.existsSync(dataDir)) {
      fs.mkdirSync(dataDir, { recursive: true });
    }
    this.localFilePath = path.join(dataDir, 'FY27_Mint_RolloverTimeline.xlsx');

    // Copy source to local working copy at startup
    this.refreshLocalCopy();
  }

  /** Enable Graph API mode after GraphService is initialized */
  enableGraphApi(graphService: GraphService): void {
    this.graphService = graphService;
    this.useGraphApi = true;
    console.log('📡 Excel sync: Graph API mode enabled (SharePoint direct read/write)');
  }

  // ---- Local file helpers ----

  /** Copy OneDrive source file to local working copy */
  private refreshLocalCopy(): void {
    try {
      if (!fs.existsSync(this.sourceFilePath)) {
        console.warn(`Source Excel not found at: ${this.sourceFilePath}`);
        return;
      }
      const srcStat = fs.statSync(this.sourceFilePath);
      const localExists = fs.existsSync(this.localFilePath);
      const localStat = localExists ? fs.statSync(this.localFilePath) : null;

      if (!localExists || srcStat.mtimeMs > (localStat?.mtimeMs || 0)) {
        fs.copyFileSync(this.sourceFilePath, this.localFilePath);
        console.log(`📋 Copied Excel source → local working copy`);
      }
    } catch (err: any) {
      console.warn(`Could not refresh local copy: ${err.message}`);
    }
  }

  /** Try to push local working copy back to the OneDrive source */
  private pushToSource(): boolean {
    try {
      fs.copyFileSync(this.localFilePath, this.sourceFilePath);
      console.log(`📤 Pushed changes back to OneDrive Excel`);
      return true;
    } catch (err: any) {
      console.warn(`📤 OneDrive file locked, changes saved locally. Will retry on next sync.`);
      return false;
    }
  }

  // ---- Import (Excel → Data Store) ----

  /** Import all data from Excel into the data store (initial load) */
  async importFromExcel(): Promise<ImportResult> {
    const result: ImportResult = { steps: 0, milestones: 0, keyDates: 0, raidEntries: 0 };

    if (this.useGraphApi && this.graphService) {
      return this.importFromExcelViaGraph(result);
    }
    return this.importFromExcelLocal(result);
  }

  /** Import via Graph API — reads directly from SharePoint */
  private async importFromExcelViaGraph(result: ImportResult): Promise<ImportResult> {
    try {
      // Read FY27_Rollover sheet
      const { values: rolloverRows } = await this.graphService!.getUsedRange('FY27_Rollover');
      for (let i = 3; i < rolloverRows.length; i++) {
        const step = parseRolloverRow(rolloverRows[i], i);
        if (step) {
          await this.dataService.upsertStep(step);
          result.steps++;
        }
      }

      // Read HighLevelMilestones sheet
      try {
        const { values: msRows } = await this.graphService!.getUsedRange('HighLevelMilestones');
        for (let i = 2; i < msRows.length; i++) {
          const ms = parseMilestoneRow(msRows[i], i);
          if (ms) {
            await this.dataService.upsertMilestone(ms);
            result.milestones++;
          }
        }
      } catch {
        console.warn('HighLevelMilestones sheet not found via Graph API');
      }

      console.log(`📡 Graph import: ${result.steps} steps, ${result.milestones} milestones`);
      return result;
    } catch (err: any) {
      console.warn(`Graph import failed (${err.message}), falling back to local file`);
      return this.importFromExcelLocal(result);
    }
  }

  /** Import from local file */
  private async importFromExcelLocal(result: ImportResult): Promise<ImportResult> {
    const wb = XLSX.readFile(this.localFilePath);

    const rolloverSheet = wb.Sheets['FY27_Rollover'];
    if (rolloverSheet) {
      const rows: any[][] = XLSX.utils.sheet_to_json(rolloverSheet, { header: 1, defval: '' });
      for (let i = 3; i < rows.length; i++) {
        const step = parseRolloverRow(rows[i], i);
        if (step) {
          await this.dataService.upsertStep(step);
          result.steps++;
        }
      }
    }

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

    console.log(`📁 Local import: ${result.steps} steps, ${result.milestones} milestones`);
    return result;
  }

  // ---- Write-back (Data Store → Excel) ----

  /** Sync changes from the data store back to Excel */
  async syncToExcel(): Promise<number> {
    if (this.useGraphApi && this.graphService) {
      return this.syncToExcelViaGraph();
    }
    return this.syncToExcelLocal();
  }

  /** Write-back via Graph API — updates individual cells in SharePoint */
  private async syncToExcelViaGraph(): Promise<number> {
    const steps = await this.dataService.getAllSteps();
    const botUpdated = steps.filter(s =>
      s.lastModifiedSource === 'bot' || s.lastModifiedSource === 'automation'
    );
    if (botUpdated.length === 0) return 0;

    // Read current Excel to find row positions
    const { values: rows } = await this.graphService!.getUsedRange('FY27_Rollover');
    const idToRow = new Map<string, number>();
    for (let i = 3; i < rows.length; i++) {
      const id = rows[i][0]?.toString().trim();
      if (id) idToRow.set(id, i + 1); // +1 for 1-based Excel row numbering
    }

    let updatedCount = 0;
    for (const step of botUpdated) {
      const excelRow = idToRow.get(step.id);
      if (!excelRow) continue;

      try {
        // Update Corp status (column F)
        await this.graphService!.updateWorksheetRange(
          'FY27_Rollover', `F${excelRow}`, [[step.corpStatus]]
        );

        // Update Corp completed date (column G)
        if (step.corpCompletedDate) {
          const ed = isoToExcelDate(step.corpCompletedDate);
          await this.graphService!.updateWorksheetRange(
            'FY27_Rollover', `G${excelRow}`, [[ed]]
          );
        }

        // Update Fed status (column J)
        if (step.fedStatus) {
          await this.graphService!.updateWorksheetRange(
            'FY27_Rollover', `J${excelRow}`, [[step.fedStatus]]
          );
        }

        // Update reference notes (column S)
        if (step.referenceNotes) {
          await this.graphService!.updateWorksheetRange(
            'FY27_Rollover', `S${excelRow}`, [[step.referenceNotes]]
          );
        }

        updatedCount++;
        console.log(`📡 Graph: updated step ${step.id} in SharePoint Excel`);

        // Reset source so it won't re-write on next cycle
        step.lastModifiedSource = 'excel';
        await this.dataService.upsertStep(step);
      } catch (err: any) {
        console.error(`📡 Graph: failed to update step ${step.id}: ${err.message}`);
      }
    }

    console.log(`📡 Graph sync: ${updatedCount} step(s) written to SharePoint Excel`);
    return updatedCount;
  }

  /** Write-back via local file */
  private async syncToExcelLocal(): Promise<number> {
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

      if (dbStep.lastModifiedSource === 'bot' || dbStep.lastModifiedSource === 'automation') {
        let changed = false;

        if (rows[i][5] !== dbStep.corpStatus) { rows[i][5] = dbStep.corpStatus; changed = true; }
        if (rows[i][9] !== dbStep.fedStatus) { rows[i][9] = dbStep.fedStatus; changed = true; }

        if (dbStep.corpCompletedDate) {
          const excelDate = isoToExcelDate(dbStep.corpCompletedDate);
          if (rows[i][6] !== excelDate) { rows[i][6] = excelDate; changed = true; }
        }

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
      console.log(`📁 Local sync: ${updatedCount} step(s) written to local Excel copy`);
      this.pushToSource();
    }

    return updatedCount;
  }

  // ---- Read-back (Excel → Data Store) ----

  /** Sync changes from Excel into the data store (pick up manual edits) */
  async syncFromExcel(): Promise<number> {
    if (this.useGraphApi && this.graphService) {
      return this.syncFromExcelViaGraph();
    }
    return this.syncFromExcelLocal();
  }

  /** Read-back via Graph API */
  private async syncFromExcelViaGraph(): Promise<number> {
    try {
      const { values: rows } = await this.graphService!.getUsedRange('FY27_Rollover');
      return this.processExcelRows(rows);
    } catch (err: any) {
      console.warn(`Graph read-back failed (${err.message}), falling back to local`);
      return this.syncFromExcelLocal();
    }
  }

  /** Read-back from local file */
  private async syncFromExcelLocal(): Promise<number> {
    this.refreshLocalCopy();
    const wb = XLSX.readFile(this.localFilePath);
    const ws = wb.Sheets['FY27_Rollover'];
    if (!ws) return 0;
    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    return this.processExcelRows(rows);
  }

  /** Common logic: compare Excel rows with data store and apply changes */
  private async processExcelRows(rows: any[][]): Promise<number> {
    let updatedCount = 0;

    for (let i = 3; i < rows.length; i++) {
      const excelStep = parseRolloverRow(rows[i], i);
      if (!excelStep) continue;

      const dbStep = await this.dataService.getStep(excelStep.id);
      if (!dbStep) {
        await this.dataService.upsertStep(excelStep);
        updatedCount++;
        continue;
      }

      let changed = false;
      const fieldsToCheck: (keyof RolloverStep)[] = [
        'corpStatus', 'fedStatus', 'corpCompletedDate', 'referenceNotes',
        'corpStartDate', 'corpEndDate', 'fedStartDate', 'fedEndDate',
      ];

      for (const field of fieldsToCheck) {
        if (excelStep[field] !== dbStep[field]) {
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

  // ---- Full Sync ----

  /** Full bi-directional sync: Excel → DB, then DB → Excel */
  async fullSync(): Promise<{ fromExcel: number; toExcel: number }> {
    if (!this.useGraphApi) this.refreshLocalCopy();
    const fromExcel = await this.syncFromExcel();
    const toExcel = await this.syncToExcel();
    console.log(`Full sync: ${fromExcel} from Excel, ${toExcel} to Excel (mode: ${this.useGraphApi ? 'Graph API' : 'local'})`);
    return { fromExcel, toExcel };
  }

  /** Get summary of current Excel data */
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
