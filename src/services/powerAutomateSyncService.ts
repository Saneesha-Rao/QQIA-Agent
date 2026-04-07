import * as https from 'https';
import { config } from '../config/appConfig';
import { RolloverStep } from '../models/types';
import { DataService } from './dataService';
import { parseRolloverRow, isoToExcelDate } from '../utils/excelParser';

/** Helper: POST JSON to a URL using https module (avoids corp proxy issues with fetch) */
function httpsPostJson(url: string, body: object): Promise<{ status: number; body: string }> {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const data = JSON.stringify(body);
    const options: https.RequestOptions = {
      hostname: parsed.hostname,
      port: parseInt(parsed.port) || 443,
      path: parsed.pathname + parsed.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(data),
      },
    };

    const req = https.request(options, (res) => {
      let responseBody = '';
      res.on('data', (chunk) => responseBody += chunk);
      res.on('end', () => resolve({ status: res.statusCode || 0, body: responseBody }));
    });

    req.on('error', reject);
    req.setTimeout(120000, () => { req.destroy(new Error('Request timeout')); });
    req.write(data);
    req.end();
  });
}

/**
 * Power Automate-based Excel sync service.
 * 
 * Uses two Power Automate flows (HTTP triggers) to read/write the SharePoint
 * Excel file without needing an App Registration or Graph API credentials.
 * The flows run under the user's own M365 connection — no admin consent needed.
 * 
 * Read flow:  Returns all rows from FY27_Rollover sheet as JSON
 * Write flow: Accepts step updates and writes them to specific cells
 */
export class PowerAutomateSyncService {
  private dataService: DataService;
  private readFlowUrl: string;
  private writeFlowUrl: string;

  constructor(dataService: DataService) {
    this.dataService = dataService;
    this.readFlowUrl = config.powerAutomate.readFlowUrl;
    this.writeFlowUrl = config.powerAutomate.writeFlowUrl;
  }

  get isConfigured(): boolean {
    return !!(this.readFlowUrl && this.writeFlowUrl);
  }

  /** Read all steps from SharePoint Excel via Power Automate */
  async readFromExcel(): Promise<RolloverStep[]> {
    if (!this.readFlowUrl) throw new Error('PA_READ_FLOW_URL not configured');

    console.log('[PA Sync] Reading Excel from SharePoint via Power Automate...');
    const response = await httpsPostJson(this.readFlowUrl, { action: 'read' });

    if (response.status !== 200) {
      throw new Error(`Power Automate read flow failed: ${response.status}`);
    }

    // Office Script returns a JSON string — may need double-parse
    let data: any;
    try {
      data = JSON.parse(response.body);
      if (typeof data === 'string') data = JSON.parse(data);
    } catch {
      throw new Error(`Failed to parse PA read response: ${response.body.substring(0, 200)}`);
    }

    const rows: any[][] = data.rows || data.value || [];
    const steps: RolloverStep[] = [];

    for (let i = 0; i < rows.length; i++) {
      const row = Array.isArray(rows[i]) ? rows[i] : Object.values(rows[i]);
      const step = parseRolloverRow(row, i + 3);
      if (step) steps.push(step);
    }

    console.log(`[PA Sync] Read ${steps.length} steps from SharePoint`);
    return steps;
  }

  /** Write step updates to SharePoint Excel via Power Automate */
  async writeToExcel(updates: StepUpdate[]): Promise<number> {
    if (!this.writeFlowUrl) throw new Error('PA_WRITE_FLOW_URL not configured');
    if (updates.length === 0) return 0;

    console.log(`[PA Sync] Writing ${updates.length} update(s) to SharePoint via Power Automate...`);
    let successCount = 0;

    // Send each update individually (Office Script handles one step at a time)
    for (const update of updates) {
      try {
        const response = await httpsPostJson(this.writeFlowUrl, {
          action: 'write',
          stepId: update.stepId,
          corpStatus: update.corpStatus,
          fedStatus: update.fedStatus,
          corpCompletedDate: update.corpCompletedDate,
          referenceNotes: update.referenceNotes,
        });

        if (response.status !== 200) {
          console.warn(`[PA Sync] Write failed for ${update.stepId}: ${response.status}`);
          continue;
        }

        const result = JSON.parse(response.body);
        if (result.updated > 0) {
          successCount++;
          console.log(`[PA Sync] Updated step ${update.stepId} in SharePoint Excel`);
        } else {
          console.warn(`[PA Sync] Step ${update.stepId} not found in Excel: ${result.error || 'unknown'}`);
        }
      } catch (err: any) {
        console.error(`[PA Sync] Error writing ${update.stepId}: ${err.message}`);
      }
    }

    console.log(`[PA Sync] Successfully wrote ${successCount}/${updates.length} update(s) to SharePoint`);
    return successCount;
  }

  /** Sync bot/automation changes to SharePoint Excel */
  async syncToExcel(): Promise<number> {
    const steps = await this.dataService.getAllSteps();
    const botUpdated = steps.filter(s =>
      s.lastModifiedSource === 'bot' || s.lastModifiedSource === 'automation'
    );

    if (botUpdated.length === 0) return 0;

    const updates: StepUpdate[] = botUpdated.map(step => ({
      stepId: step.id,
      corpStatus: step.corpStatus,
      fedStatus: step.fedStatus || '',
      corpCompletedDate: step.corpCompletedDate
        ? isoToExcelDate(step.corpCompletedDate)?.toString() || ''
        : '',
      referenceNotes: step.referenceNotes || '',
    }));

    const count = await this.writeToExcel(updates);

    // Reset source so they won't re-write on next cycle
    for (const step of botUpdated) {
      step.lastModifiedSource = 'excel';
      await this.dataService.upsertStep(step);
    }

    return count;
  }

  /** Sync SharePoint Excel changes into the data store */
  async syncFromExcel(): Promise<number> {
    const excelSteps = await this.readFromExcel();
    let updatedCount = 0;

    for (const excelStep of excelSteps) {
      const dbStep = await this.dataService.getStep(excelStep.id);
      if (!dbStep) {
        await this.dataService.upsertStep(excelStep);
        updatedCount++;
        continue;
      }

      // Only accept Excel changes if DB wasn't modified by bot/automation
      if (dbStep.lastModifiedSource === 'bot' || dbStep.lastModifiedSource === 'automation') {
        continue;
      }

      let changed = false;
      const fields: (keyof RolloverStep)[] = [
        'corpStatus', 'fedStatus', 'corpCompletedDate', 'referenceNotes',
        'corpStartDate', 'corpEndDate', 'fedStartDate', 'fedEndDate',
      ];

      for (const field of fields) {
        if (excelStep[field] !== dbStep[field]) {
          (dbStep as any)[field] = excelStep[field];
          changed = true;
        }
      }

      if (changed) {
        dbStep.lastModified = new Date().toISOString();
        dbStep.lastModifiedBy = 'pa-sync';
        dbStep.lastModifiedSource = 'excel';
        await this.dataService.upsertStep(dbStep);
        updatedCount++;
      }
    }

    console.log(`[PA Sync] Imported ${updatedCount} changed step(s) from SharePoint`);
    return updatedCount;
  }

  /** Full bi-directional sync */
  async fullSync(): Promise<{ fromExcel: number; toExcel: number }> {
    const fromExcel = await this.syncFromExcel();
    const toExcel = await this.syncToExcel();
    console.log(`[PA Sync] Full sync: ${fromExcel} from Excel, ${toExcel} to Excel`);
    return { fromExcel, toExcel };
  }
}

export interface StepUpdate {
  stepId: string;
  corpStatus: string;
  fedStatus: string;
  corpCompletedDate: string;
  referenceNotes: string;
}
