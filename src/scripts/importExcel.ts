import * as XLSX from 'xlsx';
import {
  parseRolloverRow,
  parseMilestoneRow,
  parseKeyBusinessDateRow,
  parseRaidRow,
  extractOwners,
} from '../utils/excelParser';
import { RolloverStep, Milestone } from '../models/types';

/**
 * Standalone script to import Excel data and display a summary.
 * Run with: npm run import-excel
 */

const EXCEL_PATH = process.argv[2] ||
  'C:\\Users\\salingal\\OneDrive - Microsoft\\Seller Incentives\\QQIA\\FY27_Mint_RolloverTimeline.xlsx';

function main() {
  console.log(`📂 Reading: ${EXCEL_PATH}\n`);
  const wb = XLSX.readFile(EXCEL_PATH);

  // Parse FY27_Rollover steps
  const rolloverSheet = wb.Sheets['FY27_Rollover'];
  const rolloverRows: any[][] = XLSX.utils.sheet_to_json(rolloverSheet, { header: 1, defval: '' });
  const steps: RolloverStep[] = [];

  for (let i = 3; i < rolloverRows.length; i++) {
    const step = parseRolloverRow(rolloverRows[i], i);
    if (step) steps.push(step);
  }

  console.log(`✅ Parsed ${steps.length} rollover steps\n`);

  // Status breakdown
  const statusCount: Record<string, number> = {};
  const fedStatusCount: Record<string, number> = {};
  const workstreamCount: Record<string, { total: number; completed: number }> = {};

  for (const step of steps) {
    statusCount[step.corpStatus] = (statusCount[step.corpStatus] || 0) + 1;
    fedStatusCount[step.fedStatus] = (fedStatusCount[step.fedStatus] || 0) + 1;

    if (!workstreamCount[step.workstream]) {
      workstreamCount[step.workstream] = { total: 0, completed: 0 };
    }
    workstreamCount[step.workstream].total++;
    if (step.corpStatus === 'Completed') workstreamCount[step.workstream].completed++;
  }

  console.log('📊 Corp Status Breakdown:');
  for (const [status, count] of Object.entries(statusCount)) {
    console.log(`   ${status}: ${count}`);
  }

  console.log('\n📊 Fed Status Breakdown:');
  for (const [status, count] of Object.entries(fedStatusCount)) {
    console.log(`   ${status}: ${count}`);
  }

  console.log('\n📦 Workstream Summary:');
  for (const [ws, counts] of Object.entries(workstreamCount)) {
    const pct = Math.round((counts.completed / counts.total) * 100);
    console.log(`   ${ws}: ${counts.completed}/${counts.total} (${pct}%)`);
  }

  // Owner mapping
  const owners = extractOwners(steps);
  console.log(`\n👥 Unique Owners/DRIs: ${owners.size}`);
  for (const [name, stepIds] of owners) {
    console.log(`   ${name}: ${stepIds.length} step(s)`);
  }

  // Dependency analysis
  const stepsWithDeps = steps.filter(s => s.dependencies.length > 0);
  console.log(`\n🔗 Steps with dependencies: ${stepsWithDeps.length}`);

  const stepIds = new Set(steps.map(s => s.id));
  const brokenDeps: string[] = [];
  for (const step of steps) {
    for (const dep of step.dependencies) {
      if (dep.includes('-')) continue; // Range deps like "1.A - 1.G"
      if (!stepIds.has(dep)) brokenDeps.push(`${step.id} → ${dep}`);
    }
  }
  if (brokenDeps.length > 0) {
    console.log(`\n⚠️ Broken dependency references (${brokenDeps.length}):`);
    for (const bd of brokenDeps.slice(0, 10)) console.log(`   ${bd}`);
  }

  // Parse milestones
  const msSheet = wb.Sheets['HighLevelMilestones'];
  const msRows: any[][] = XLSX.utils.sheet_to_json(msSheet, { header: 1, defval: '' });
  const milestones: Milestone[] = [];
  for (let i = 2; i < msRows.length; i++) {
    const ms = parseMilestoneRow(msRows[i], i);
    if (ms) milestones.push(ms);
  }
  console.log(`\n🏁 High-Level Milestones: ${milestones.length}`);
  for (const ms of milestones) {
    console.log(`   [${ms.corpStatus}] ${ms.category}: ${ms.milestone} (${ms.corpDate || 'TBD'})`);
  }

  console.log('\n✅ Import validation complete!');
}

main();
