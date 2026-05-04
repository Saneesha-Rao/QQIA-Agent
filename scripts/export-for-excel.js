/**
 * Generates a ready-to-run Office Script that syncs QQIA Agent changes
 * into Excel Online. Just copy the output, paste into Automate > Script editor, and Run.
 * 
 * Usage: node scripts/export-for-excel.js
 * 
 * Requires: bot server running on localhost:3978
 */

const BOT_URL = 'http://localhost:3978/api/steps/json';

async function main() {
  let data;
  try {
    const resp = await fetch(BOT_URL);
    data = await resp.json();
  } catch (err) {
    console.error('ERROR: Cannot reach bot at', BOT_URL);
    console.error('Make sure the server is running: npx tsc && node dist/index.js');
    process.exit(1);
  }

  const steps = data.steps || [];

  // Get ALL steps with their current status (so the script can compare)
  const allSteps = steps.map(s => ({
    id: s.id,
    corpStatus: s.corpStatus || '',
    fedStatus: s.fedStatus || '',
    corpCompletedDate: s.corpCompletedDate || '',
  }));

  if (allSteps.length === 0) {
    console.error('No steps found from bot.');
    process.exit(1);
  }

  // Now read the local Excel file to find the status column index
  // (We'll hardcode the column mapping based on the sheet layout)

  // Generate the Office Script
  const script = `function main(workbook: ExcelScript.Workbook) {
  // Auto-generated on ${new Date().toISOString()}
  // Source: QQIA Agent bot data

  const botData: { id: string; corpStatus: string }[] = ${JSON.stringify(
    allSteps.map(s => ({ id: s.id, corpStatus: s.corpStatus })),
    null, 4
  )};

  const sheet = workbook.getWorksheet("FY27_Rollover");
  if (!sheet) { console.log("Sheet FY27_Rollover not found!"); return; }

  const range = sheet.getUsedRange();
  const values = range.getValues();

  // Find which column has "Status" or contains status values
  // From the sheet: A=Step, B=Grouping, C=Description, D=StartDate, E=EndDate, F=Status...
  // We'll auto-detect by scanning the header row
  let statusCol = -1;
  const headers = values[0];
  for (let c = 0; c < headers.length; c++) {
    const h = headers[c]?.toString().toLowerCase() || "";
    if (h.includes("status") && !h.includes("fed")) {
      statusCol = c;
      break;
    }
  }
  if (statusCol === -1) {
    // Fallback: try common positions
    for (let c = 3; c < Math.min(headers.length, 10); c++) {
      const sample = values[3]?.[c]?.toString() || "";
      if (["Not Started", "In Progress", "Completed", "Blocked"].includes(sample)) {
        statusCol = c;
        break;
      }
    }
  }
  if (statusCol === -1) {
    console.log("Could not find Status column! Check your sheet layout.");
    return;
  }
  console.log("Status column index: " + statusCol + " (header: " + headers[statusCol] + ")");

  // Build StepId -> row map
  const stepRows: { [key: string]: number } = {};
  for (let i = 1; i < values.length; i++) {
    const id = values[i][0]?.toString().trim();
    if (id) stepRows[id] = i;
  }

  let updated = 0;
  let skipped = 0;
  for (const bot of botData) {
    const row = stepRows[bot.id];
    if (row === undefined) continue;

    const current = values[row][statusCol]?.toString().trim() || "";
    if (bot.corpStatus && bot.corpStatus !== current) {
      sheet.getCell(row, statusCol).setValue(bot.corpStatus);
      console.log(bot.id + ": " + current + " → " + bot.corpStatus);
      updated++;
    } else {
      skipped++;
    }
  }

  console.log("Done: " + updated + " updated, " + skipped + " unchanged");
}`;

  console.log(script);
  console.error(`\n=== Copy ALL the output above, paste into Excel Online Automate > Script editor, click Run ===`);
}

main();
