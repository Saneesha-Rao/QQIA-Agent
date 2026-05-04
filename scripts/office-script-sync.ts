/**
 * QQIA Sync — STANDALONE Office Script for Excel Online
 * 
 * HOW IT WORKS:
 * 1. The bot writes qqia-bot-sync.xlsx to your OneDrive QQIA folder
 * 2. You open qqia-bot-sync.xlsx in Excel Online, select all data, copy
 * 3. In the main workbook, paste into a sheet called "BotUpdates"
 * 4. Run this script — it reads BotUpdates and pushes changes to FY27_Rollover
 * 
 * NOTE: Reads only the columns it needs to stay under Office Script data limits.
 */
function main(workbook: ExcelScript.Workbook) {
  const updateSheet = workbook.getWorksheet("BotUpdates");
  if (!updateSheet) {
    console.log("Create a sheet called 'BotUpdates' and paste data from qqia-bot-sync.xlsx");
    return;
  }

  const updateRange = updateSheet.getUsedRange();
  if (!updateRange) { console.log("BotUpdates is empty."); return; }
  const updateRows = updateRange.getValues();
  if (updateRows.length <= 1) { console.log("No data rows in BotUpdates."); return; }

  const dataSheet = workbook.getWorksheet("FY27_Rollover");
  if (!dataSheet) { console.log("FY27_Rollover not found!"); return; }

  // Read only the header row (row 1) to find the status column — avoids full-sheet read
  const lastCol = dataSheet.getUsedRange()?.getColumnCount() || 20;
  const headerRange = dataSheet.getRangeByIndexes(0, 0, 1, lastCol);
  const headers = headerRange.getValues()[0];

  // Auto-detect Corp Status column
  let statusCol = -1;
  for (let c = 0; c < headers.length; c++) {
    const h = headers[c]?.toString().toLowerCase() || "";
    if (h.includes("status") && !h.includes("fed") && !h.includes("setup")) {
      statusCol = c;
      break;
    }
  }
  if (statusCol === -1) {
    // Fallback: scan a small block of data rows for status-like values
    const sampleRange = dataSheet.getRangeByIndexes(3, 3, 7, Math.min(lastCol - 3, 9));
    const sampleVals = sampleRange.getValues();
    for (let c = 0; c < sampleVals[0].length; c++) {
      for (let r = 0; r < sampleVals.length; r++) {
        const v = sampleVals[r][c]?.toString() || "";
        if (["Not Started", "In Progress", "Completed", "Blocked", "N/A"].includes(v)) {
          statusCol = c + 3; break;
        }
      }
      if (statusCol !== -1) break;
    }
  }
  if (statusCol === -1) { console.log("Could not find Status column!"); return; }
  console.log("Status column: " + statusCol + " (" + headers[statusCol] + ")");

  // Read only column A (StepId) and the status column — much smaller than full sheet
  const lastRow = dataSheet.getUsedRange()?.getRowCount() || 200;
  const stepIdCol = dataSheet.getRangeByIndexes(0, 0, lastRow, 1).getValues();
  const statusVals = dataSheet.getRangeByIndexes(0, statusCol, lastRow, 1).getValues();

  // Build StepId → row map
  const stepRows: { [key: string]: number } = {};
  for (let i = 1; i < stepIdCol.length; i++) {
    const id = stepIdCol[i][0]?.toString().trim();
    if (id) stepRows[id] = i;
  }

  // BotUpdates columns: StepId | CorpStatus | FedStatus | CompletedDate | Notes | Source | Modified
  let updated = 0;
  let skipped = 0;
  for (let i = 1; i < updateRows.length; i++) {
    const stepId = updateRows[i][0]?.toString().trim();
    const botStatus = updateRows[i][1]?.toString().trim();
    if (!stepId || !botStatus) continue;

    const row = stepRows[stepId];
    if (row === undefined) continue;

    const current = statusVals[row][0]?.toString().trim() || "";
    if (botStatus !== current) {
      dataSheet.getCell(row, statusCol).setValue(botStatus);
      console.log(stepId + ": " + current + " → " + botStatus);
      updated++;
    } else {
      skipped++;
    }
  }

  // Clear BotUpdates and write summary
  updateRange.clear(ExcelScript.ClearApplyTo.contents);
  updateSheet.getRange("A1").setValue(
    "Last sync: " + new Date().toISOString() + " — " + updated + " updated, " + skipped + " unchanged"
  );

  console.log("Done: " + updated + " updated, " + skipped + " unchanged");
}
