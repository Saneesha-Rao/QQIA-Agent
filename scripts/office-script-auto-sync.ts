/**
 * QQIA Sync — Office Script for Excel Online
 * 
 * This script receives bot step data as a JSON string parameter
 * and updates the FY27_Rollover sheet accordingly.
 * 
 * Called by Power Automate (Automation Templates) which reads
 * data from the qqia-bot-sync.xlsx staging file.
 * 
 * SETUP:
 * 1. Paste this into Automate > New Script > save as "QQIA Sync"
 * 2. Click "Automation Templates" in the Automate ribbon
 * 3. Pick "Run script from another workbook" or create a custom flow:
 *    - Trigger: Recurrence (e.g., every 1 hour) or Manual
 *    - Action 1: Excel Online > List rows in qqia-bot-sync.xlsx, sheet BotData
 *    - Action 2: For each row, or Run Script on this workbook
 */
function main(
  workbook: ExcelScript.Workbook,
  stepId: string,
  corpStatus: string,
  fedStatus?: string,
  corpCompletedDate?: string,
  referenceNotes?: string
) {
  const sheet = workbook.getWorksheet("FY27_Rollover");
  if (!sheet) { console.log("Sheet not found!"); return; }

  const range = sheet.getUsedRange();
  if (!range) { console.log("Sheet is empty!"); return; }
  const values = range.getValues();

  // Find the step row
  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]?.toString().trim() === stepId) {
      targetRow = i;
      break;
    }
  }
  if (targetRow === -1) {
    console.log(`Step ${stepId} not found`);
    return;
  }

  // Auto-detect status column by scanning headers
  let statusCol = -1;
  const headers = values[0];
  for (let c = 0; c < headers.length; c++) {
    const h = headers[c]?.toString().toLowerCase() || "";
    if (h.includes("status") && !h.includes("fed") && !h.includes("setup")) {
      statusCol = c;
      break;
    }
  }

  // Fallback: find column by checking sample data for known status values
  if (statusCol === -1) {
    for (let c = 3; c < Math.min(headers.length, 12); c++) {
      for (let r = 3; r < Math.min(values.length, 10); r++) {
        const v = values[r][c]?.toString() || "";
        if (["Not Started", "In Progress", "Completed", "Blocked", "N/A"].includes(v)) {
          statusCol = c;
          break;
        }
      }
      if (statusCol !== -1) break;
    }
  }

  if (statusCol === -1) {
    console.log("Could not find Status column!");
    return;
  }

  const currentStatus = values[targetRow][statusCol]?.toString().trim() || "";

  if (corpStatus && corpStatus !== currentStatus) {
    sheet.getCell(targetRow, statusCol).setValue(corpStatus);
    console.log(`${stepId}: ${currentStatus} → ${corpStatus}`);
  } else {
    console.log(`${stepId}: no change (already ${currentStatus})`);
  }
}
