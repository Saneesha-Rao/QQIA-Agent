/**
 * Office Script for Excel Online — paste this into Excel > Automate > New Script
 * This script is called by Power Automate to READ all steps.
 * 
 * Output: { rows: string[][] } — all data rows from FY27_Rollover sheet
 */
function main(workbook: ExcelScript.Workbook): { rows: string[][] } {
  const sheet = workbook.getWorksheet("FY27_Rollover");
  if (!sheet) return { rows: [] };

  const range = sheet.getUsedRange();
  const values = range.getValues();

  // Skip header rows (first 3 rows), return data as strings
  const rows: string[][] = [];
  for (let i = 3; i < values.length; i++) {
    const row = values[i].map(cell => cell?.toString() || "");
    // Only include rows that have a Step ID
    if (row[0] && row[0].trim()) {
      rows.push(row);
    }
  }

  return { rows };
}
