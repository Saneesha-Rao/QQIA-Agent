const XLSX = require('xlsx');
const fs = require('fs');
const xlsxPath = 'C:/Users/salingal/OneDrive - Microsoft/Seller Incentives/QQIA/FY27_Mint_RolloverTimeline.xlsx';
const wb = XLSX.readFile(xlsxPath);
const ws = wb.Sheets['FY27_Rollover'];
const data = XLSX.utils.sheet_to_json(ws, { range: 2 });

function excelDate(serial) {
  if (!serial || serial === '-' || typeof serial !== 'number') return '';
  const d = new Date((serial - 25569) * 86400000);
  return d.toISOString().split('T')[0];
}

function esc(val) {
  const s = String(val || '').replace(/"/g, '""');
  return '"' + s + '"';
}

const headers = [
  'StepId','Workstream','Description','CorpStartDate','CorpEndDate',
  'CorpStatus','CorpCompletedDate','FedStartDate','FedEndDate','FedStatus',
  'SetupValidation','EngineeringDependent','WWIC_POC','Fed_POC',
  'EngineeringDRI','EngineeringLead','ADOLink','ReferenceNotes'
];

const rows = data.filter(r => r['Step']).map(r => [
  esc(r['Step']),
  esc(r['Grouping'] || ''),
  esc(r['Description'] || ''),
  esc(excelDate(r['StartDate'])),
  esc(excelDate(r['EndDate'])),
  esc(r['Status'] || 'Not Started'),
  esc(excelDate(r['Completed Date'])),
  esc(excelDate(r['FedStartDate'])),
  esc(excelDate(r['FedEndDate'])),
  esc(r['Fed Status'] || ''),
  esc(r['Setup validation'] || ''),
  esc(r['EngineeringDependent'] || ''),
  esc(r['WWIC POC'] || ''),
  esc(r['Fed POC'] || ''),
  esc(r['Engineering DRI'] || ''),
  esc(r['Engineering Lead'] || ''),
  esc(r['ADO_Link'] || ''),
  esc('')
].join(','));

const csv = [headers.join(','), ...rows].join('\n');
fs.writeFileSync('data/FY27_Rollover_Dataverse.csv', csv);
console.log(`Created CSV: ${rows.length} rows, ${headers.length} columns`);
console.log('File: data/FY27_Rollover_Dataverse.csv');
