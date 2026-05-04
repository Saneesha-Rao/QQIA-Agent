import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';
import { config } from '../config/appConfig';
import { RolloverStep, Milestone, KeyBusinessDate, RaidEntry } from '../models/types';
import { DataService } from './dataService';
import { GraphService } from './graphService';
import { DelegatedGraphService } from './delegatedGraphService';
import {
  parseRolloverRow,
  parseMilestoneRow,
  parseKeyBusinessDateRow,
  parseRaidRow,
  isoToExcelDate,
  normalizeStatus,
  parseDependencies,
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
  private delegatedGraph: DelegatedGraphService | null = null;
  private delegatedDriveId: string = '';
  private delegatedItemId: string = '';
  private useGraphApi: boolean = false;
  private useDelegatedGraph: boolean = false;
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

  /** Enable delegated Graph mode (device code auth — user's own permissions) */
  enableDelegatedGraph(delegatedGraph: DelegatedGraphService, driveId: string, itemId: string): void {
    this.delegatedGraph = delegatedGraph;
    this.delegatedDriveId = driveId;
    this.delegatedItemId = itemId;
    this.useDelegatedGraph = true;
    console.log('📡 Excel sync: Delegated Graph mode enabled (real-time sync)');
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

    if (this.useDelegatedGraph && this.delegatedGraph) {
      return this.importFromExcelViaDelegated(result);
    }
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

  /** Import via delegated Graph API — reads using user's own credentials */
  private async importFromExcelViaDelegated(result: ImportResult): Promise<ImportResult> {
    try {
      const dg = this.delegatedGraph!;
      const { values: rolloverRows } = await dg.getUsedRange(this.delegatedDriveId, this.delegatedItemId, 'FY27_Rollover');
      for (let i = 3; i < rolloverRows.length; i++) {
        const step = parseRolloverRow(rolloverRows[i], i);
        if (step) {
          await this.dataService.upsertStep(step);
          result.steps++;
        }
      }

      try {
        const { values: msRows } = await dg.getUsedRange(this.delegatedDriveId, this.delegatedItemId, 'HighLevelMilestones');
        for (let i = 2; i < msRows.length; i++) {
          const ms = parseMilestoneRow(msRows[i], i);
          if (ms) {
            await this.dataService.upsertMilestone(ms);
            result.milestones++;
          }
        }
      } catch {
        console.warn('HighLevelMilestones sheet not found via delegated Graph');
      }

      console.log(`📡 Delegated import: ${result.steps} steps, ${result.milestones} milestones`);
      return result;
    } catch (err: any) {
      console.warn(`Delegated import failed (${err.message}), falling back to local file`);
      return this.importFromExcelLocal(result);
    }
  }

  /** Import from local file (xlsx or CSV fallback) */
  private async importFromExcelLocal(result: ImportResult): Promise<ImportResult> {
    // Try xlsx first, then fall back to CSV
    if (fs.existsSync(this.localFilePath)) {
      return this.importFromExcelFile(result);
    }

    // CSV fallback — look for the bundled CSV in data/
    const csvPath = path.join(path.dirname(this.localFilePath), 'FY27_Rollover_Dataverse.csv');
    if (fs.existsSync(csvPath)) {
      return this.importFromCSV(csvPath, result);
    }

    console.warn('⚠️ No Excel or CSV file found for import');
    return result;
  }

  /** Import from .xlsx file */
  private async importFromExcelFile(result: ImportResult): Promise<ImportResult> {
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

  /** Import from CSV file (used when xlsx not available, e.g. cloud deploy) */
  private async importFromCSV(csvPath: string, result: ImportResult): Promise<ImportResult> {
    const csvContent = fs.readFileSync(csvPath, 'utf-8');
    const wb = XLSX.read(csvContent, { type: 'string' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // CSV has a header row at index 0, data starts at index 1
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const stepId = row[0]?.toString().trim();
      if (!stepId || stepId === 'StepId' || stepId === '') continue;

      const step: RolloverStep = {
        id: stepId,
        workstream: (row[1] || '').toString().trim(),
        description: (row[2] || '').toString().trim(),
        corpStartDate: (row[3] || '').toString().trim() || null,
        corpEndDate: (row[4] || '').toString().trim() || null,
        corpStatus: normalizeStatus(row[5]) as any,
        corpCompletedDate: (row[6] || '').toString().trim() || null,
        fedStartDate: (row[7] || '').toString().trim() || null,
        fedEndDate: (row[8] || '').toString().trim() || null,
        fedStatus: normalizeStatus(row[9]) as any,
        setupValidation: (row[10] || '').toString().trim(),
        engineeringDependent: (row[11] || '').toString().trim(),
        wwicPoc: (row[12] || '').toString().trim(),
        fedPoc: (row[13] || '').toString().trim(),
        engineeringDri: (row[14] || '').toString().trim(),
        engineeringLead: (row[15] || '').toString().trim(),
        dependencies: parseDependencies(row[16]?.toString()),
        adoLink: (row[17] || '').toString().trim(),
        referenceNotes: (row[18] || '').toString().trim(),
        fy26CorpStart: null,
        fy26CorpEnd: null,
        fy26FedStart: null,
        fy26FedEnd: null,
        lastModified: new Date().toISOString(),
        lastModifiedBy: 'system',
        lastModifiedSource: 'excel',
      };
      await this.dataService.upsertStep(step);
      result.steps++;
    }

    console.log(`📁 CSV import: ${result.steps} steps`);
    return result;
  }

  // ---- Write-back (Data Store → Excel) ----

  /** Sync changes from the data store back to Excel */
  async syncToExcel(): Promise<number> {
    if (this.useDelegatedGraph && this.delegatedGraph) {
      return this.syncToExcelViaDelegated();
    }
    if (this.useGraphApi && this.graphService) {
      return this.syncToExcelViaGraph();
    }
    return this.syncToExcelLocal();
  }

  /** Write-back via Graph API — updates individual cells in SharePoint */
  private async syncToExcelViaGraph(): Promise<number> {
    const steps = await this.dataService.getAllSteps();
    const botUpdated = steps.filter(s =>
      s.lastModifiedSource === 'bot' || s.lastModifiedSource === 'automation' || s.lastModifiedSource === 'webhook'
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

        // Reset source — use 'com_synced' to protect from stale reads
        step.lastModifiedSource = 'com_synced';
        await this.dataService.upsertStep(step);
      } catch (err: any) {
        console.error(`📡 Graph: failed to update step ${step.id}: ${err.message}`);
      }
    }

    console.log(`📡 Graph sync: ${updatedCount} step(s) written to SharePoint Excel`);
    return updatedCount;
  }

  /** Write-back via delegated Graph API — uses user's own credentials */
  private async syncToExcelViaDelegated(): Promise<number> {
    const dg = this.delegatedGraph!;
    const steps = await this.dataService.getAllSteps();
    const botUpdated = steps.filter(s =>
      s.lastModifiedSource === 'bot' || s.lastModifiedSource === 'automation' || s.lastModifiedSource === 'webhook'
    );
    if (botUpdated.length === 0) return 0;

    // Read current Excel to find row positions
    const { values: rows } = await dg.getUsedRange(this.delegatedDriveId, this.delegatedItemId, 'FY27_Rollover');
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
        await dg.updateWorksheetRange(
          this.delegatedDriveId, this.delegatedItemId,
          'FY27_Rollover', `F${excelRow}`, [[step.corpStatus]]
        );

        // Update Corp completed date (column G)
        if (step.corpCompletedDate) {
          await dg.updateWorksheetRange(
            this.delegatedDriveId, this.delegatedItemId,
            'FY27_Rollover', `G${excelRow}`, [[step.corpCompletedDate]]
          );
        }

        // Update Fed status (column J)
        if (step.fedStatus) {
          await dg.updateWorksheetRange(
            this.delegatedDriveId, this.delegatedItemId,
            'FY27_Rollover', `J${excelRow}`, [[step.fedStatus]]
          );
        }

        // Update reference notes (column S)
        if (step.referenceNotes) {
          await dg.updateWorksheetRange(
            this.delegatedDriveId, this.delegatedItemId,
            'FY27_Rollover', `S${excelRow}`, [[step.referenceNotes]]
          );
        }

        updatedCount++;
        console.log(`📡 Delegated: updated step ${step.id} in SharePoint Excel`);

        // Reset source — use 'com_synced' to protect from stale reads
        step.lastModifiedSource = 'com_synced';
        await this.dataService.upsertStep(step);
      } catch (err: any) {
        console.error(`📡 Delegated: failed to update step ${step.id}: ${err.message}`);
      }
    }

    console.log(`📡 Delegated sync: ${updatedCount} step(s) written to SharePoint Excel`);
    return updatedCount;
  }

  /** Write-back via local file */
  private async syncToExcelLocal(): Promise<number> {
    const steps = await this.dataService.getAllSteps();
    const botUpdated = steps.filter(s =>
      s.lastModifiedSource === 'bot' || s.lastModifiedSource === 'automation' || s.lastModifiedSource === 'webhook'
    );
    if (botUpdated.length === 0) return 0;

    const stepMap = new Map(botUpdated.map(s => [s.id, s]));

    // Use ExcelJS for writes to avoid SheetJS file corruption/bloat
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(this.localFilePath);
    const ws = wb.getWorksheet('FY27_Rollover');
    if (!ws) return 0;

    let updatedCount = 0;
    ws.eachRow((row, rowNum) => {
      if (rowNum <= 3) return; // skip header rows
      const stepId = row.getCell(1).text?.trim();
      if (!stepId) return;

      const dbStep = stepMap.get(stepId);
      if (!dbStep) return;

      let changed = false;

      if (row.getCell(6).text !== dbStep.corpStatus) {
        row.getCell(6).value = dbStep.corpStatus;
        changed = true;
      }
      if (row.getCell(10).text !== dbStep.fedStatus) {
        row.getCell(10).value = dbStep.fedStatus;
        changed = true;
      }
      if (dbStep.corpCompletedDate) {
        const excelDate = isoToExcelDate(dbStep.corpCompletedDate);
        const currentVal = row.getCell(7).value;
        if (currentVal !== excelDate) {
          row.getCell(7).value = excelDate;
          changed = true;
        }
      }
      if (dbStep.referenceNotes && row.getCell(19).text !== dbStep.referenceNotes) {
        row.getCell(19).value = dbStep.referenceNotes;
        changed = true;
      }

      if (changed) updatedCount++;
    });

    if (updatedCount > 0) {
      await wb.xlsx.writeFile(this.localFilePath);
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
