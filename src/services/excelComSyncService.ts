import { execSync } from 'child_process';
import * as path from 'path';
import { config } from '../config/appConfig';
import { RolloverStep } from '../models/types';
import { DataService } from './dataService';
import { isoToExcelDate } from '../utils/excelParser';

/**
 * Excel sync via COM automation — uses the running Excel desktop instance
 * to read/write the OneDrive-synced file without file lock issues.
 * 
 * Calls a PowerShell 5.1 script that uses COM to access Excel.
 * Requires: Excel must be running (or installable) on the machine.
 */
export class ExcelComSyncService {
  private dataService: DataService;
  private scriptPath: string;

  constructor(dataService: DataService) {
    this.dataService = dataService;
    this.scriptPath = path.resolve(__dirname, '..', '..', 'scripts', 'excel-sync.ps1');
  }

  /** Check if this sync method is available (Excel installed) */
  get isAvailable(): boolean {
    try {
      const result = execSync(
        'powershell.exe -NoProfile -Command "Get-Process excel -ErrorAction SilentlyContinue | Select-Object -First 1 Id"',
        { timeout: 5000, encoding: 'utf-8' }
      );
      return result.trim().length > 0;
    } catch {
      return false;
    }
  }

  /** Run the PowerShell sync script and return parsed JSON result */
  private runScript(args: string): any {
    const cmd = `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "${this.scriptPath}" ${args}`;
    try {
      const output = execSync(cmd, {
        timeout: 60000,
        encoding: 'utf-8',
        windowsHide: true,
      });
      return JSON.parse(output.trim());
    } catch (err: any) {
      // Try to parse stderr/stdout for JSON error
      const stdout = err.stdout?.toString().trim();
      if (stdout) {
        try { return JSON.parse(stdout); } catch { /* not JSON */ }
      }
      throw new Error(`Excel COM script failed: ${err.message}`);
    }
  }

  /** Write a single step status to the SharePoint Excel file */
  async writeStep(stepId: string, corpStatus: string, fedStatus?: string,
                  corpCompletedDate?: string, referenceNotes?: string): Promise<boolean> {
    const args = [
      '-Action write',
      `-StepId "${stepId}"`,
    ];
    if (corpStatus) args.push(`-CorpStatus "${corpStatus}"`);
    if (fedStatus) args.push(`-FedStatus "${fedStatus}"`);
    if (corpCompletedDate) args.push(`-CorpCompletedDate "${corpCompletedDate}"`);
    if (referenceNotes) args.push(`-ReferenceNotes "${referenceNotes.replace(/"/g, '""')}"`);

    const result = this.runScript(args.join(' '));
    if (result.updated > 0) {
      console.log(`[Excel COM] Updated ${stepId} → ${corpStatus} (row ${result.row})`);
      return true;
    }
    console.warn(`[Excel COM] Step ${stepId} not found: ${result.error}`);
    return false;
  }

  /** Sync all bot/automation changes to the Excel file */
  async syncToExcel(): Promise<number> {
    const steps = await this.dataService.getAllSteps();
    const botUpdated = steps.filter(s =>
      s.lastModifiedSource === 'bot' || s.lastModifiedSource === 'automation' || s.lastModifiedSource === 'webhook'
    );

    if (botUpdated.length === 0) return 0;

    let successCount = 0;
    for (const step of botUpdated) {
      try {
        const completedDate = step.corpCompletedDate || '';
        const success = await this.writeStep(
          step.id,
          step.corpStatus,
          step.fedStatus || '',
          completedDate,
          step.referenceNotes || ''
        );
        if (success) {
          successCount++;
          // Mark as 'com_synced' — NOT 'excel' — so SyncEngine won't revert
          // the change when it reads the (potentially stale) file
          step.lastModifiedSource = 'com_synced';
          await this.dataService.upsertStep(step);
        }
      } catch (err: any) {
        console.error(`[Excel COM] Failed to sync ${step.id}: ${err.message}`);
      }
    }

    console.log(`[Excel COM] Synced ${successCount}/${botUpdated.length} step(s) to SharePoint Excel`);
    return successCount;
  }

  /** Read all steps from the Excel file */
  async readFromExcel(): Promise<any[]> {
    const result = this.runScript('-Action read');
    if (result.error) throw new Error(result.error);
    console.log(`[Excel COM] Read ${result.count} steps from SharePoint Excel`);
    return result.rows || [];
  }
}
