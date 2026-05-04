/**
 * Office Script for Excel Online — paste this into Excel > Automate > New Script
 * This script is called by Power Automate to update step statuses.
 * 
 * Input: { stepId, corpStatus, fedStatus, corpCompletedDate, referenceNotes }
 * Output: { updated: number, error?: string }
 */
function main(
  workbook: ExcelScript.Workbook,
  stepId: string,
  corpStatus: string,
  fedStatus: string,
  corpCompletedDate: string,
  referenceNotes: string
): { updated: number; error?: string } {
  const sheet = workbook.getWorksheet("FY27_Rollover");
  if (!sheet) return { updated: 0, error: "Sheet FY27_Rollover not found" };

  const range = sheet.getUsedRange();
  const values = range.getValues();

  // Find the row with matching Step ID (column A, index 0)
  let targetRow = -1;
  for (let i = 3; i < values.length; i++) {
    const cellVal = values[i][0]?.toString().trim();
    if (cellVal === stepId) {
      targetRow = i;
      break;
    }
  }

  if (targetRow === -1) return { updated: 0, error: `Step ${stepId} not found` };

  // Column mapping (0-indexed):
  // A=0: Step ID
  // B=1: Category  
  // C=2: Step Description
  // D=3: Owner
  // E=4: Corp Status
  // F=5: Corp Start Date
  // G=6: Corp End Date
  // H=7: Corp Completed Date
  // I=8: Fed Status
  // J=9: Reference/Notes

  if (corpStatus) {
    sheet.getCell(targetRow, 4).setValue(corpStatus);
  }
  if (fedStatus) {
    sheet.getCell(targetRow, 8).setValue(fedStatus);
  }
  if (corpCompletedDate) {
    sheet.getCell(targetRow, 7).setValue(corpCompletedDate);
  }
  if (referenceNotes) {
    sheet.getCell(targetRow, 9).setValue(referenceNotes);
  }

  return { updated: 1 };
}
