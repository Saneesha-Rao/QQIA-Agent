import { RolloverStep, Milestone, KeyBusinessDate, RaidEntry } from '../models/types';

/**
 * Excel date serial number to JS Date conversion.
 * Excel uses epoch 1900-01-01 with a known leap year bug (treats 1900 as leap year).
 */
export function excelDateToISO(serial: number | string | null): string | null {
  if (serial === null || serial === undefined || serial === '' || serial === 'TBD' || serial === '-') {
    return null;
  }
  const num = typeof serial === 'string' ? parseFloat(serial) : serial;
  if (isNaN(num) || num < 1) return typeof serial === 'string' ? serial : null;

  // Excel epoch: 1900-01-01, but Excel incorrectly treats 1900 as leap year
  // so dates after Feb 28, 1900 are off by 1 day
  const excelEpoch = new Date(1900, 0, 1);
  const dayOffset = num > 59 ? num - 2 : num - 1; // Adjust for Excel's 1900 leap year bug
  const date = new Date(excelEpoch.getTime() + dayOffset * 86400000);
  return date.toISOString().split('T')[0];
}

/** Convert ISO date string back to Excel serial number */
export function isoToExcelDate(isoDate: string | null): number | null {
  if (!isoDate) return null;
  const date = new Date(isoDate);
  if (isNaN(date.getTime())) return null;
  const excelEpoch = new Date(1900, 0, 1);
  const diffDays = Math.round((date.getTime() - excelEpoch.getTime()) / 86400000);
  return diffDays > 59 ? diffDays + 2 : diffDays + 1;
}

/** Parse dependencies string into array of step IDs */
export function parseDependencies(depString: string | null): string[] {
  if (!depString || depString.trim() === '') return [];
  return depString
    .split(/[,;]/)
    .map(d => d.trim())
    .filter(d => d.length > 0 && d !== '-');
}

/** Parse FY27_Rollover sheet row into a RolloverStep */
export function parseRolloverRow(row: any[], index: number): RolloverStep | null {
  const stepId = row[0]?.toString().trim();
  if (!stepId || stepId === 'Step' || stepId === '') return null;

  return {
    id: stepId,
    workstream: (row[1] || '').toString().trim(),
    description: (row[2] || '').toString().trim(),
    corpStartDate: excelDateToISO(row[3]),
    corpEndDate: excelDateToISO(row[4]),
    corpStatus: normalizeStatus(row[5]),
    corpCompletedDate: excelDateToISO(row[6]),
    fedStartDate: excelDateToISO(row[7]),
    fedEndDate: excelDateToISO(row[8]),
    fedStatus: normalizeStatus(row[9]),
    setupValidation: (row[10] || '').toString().trim(),
    engineeringDependent: (row[11] || '').toString().trim(),
    wwicPoc: (row[12] || '').toString().trim(),
    fedPoc: (row[13] || '').toString().trim(),
    engineeringDri: (row[14] || '').toString().trim(),
    engineeringLead: (row[15] || '').toString().trim(),
    dependencies: parseDependencies(row[16]?.toString()),
    adoLink: (row[17] || '').toString().trim(),
    referenceNotes: (row[18] || '').toString().trim(),
    fy26CorpStart: excelDateToISO(row[20]),
    fy26CorpEnd: excelDateToISO(row[21]),
    fy26FedStart: excelDateToISO(row[23]),
    fy26FedEnd: excelDateToISO(row[24]),
    lastModified: new Date().toISOString(),
    lastModifiedBy: 'system',
    lastModifiedSource: 'excel',
  };
}

/** Parse KeyBusinessDates row */
export function parseKeyBusinessDateRow(row: any[]): KeyBusinessDate | null {
  const category = (row[0] || '').toString().trim();
  if (!category || category === 'Category' || category === 'Copied from') return null;
  const milestone = (row[2] || '').toString().trim();
  if (!milestone) return null;

  return {
    id: `kbd-${category}-${milestone}`.replace(/[^a-zA-Z0-9-]/g, '_').substring(0, 100),
    category,
    driTeam: (row[1] || '').toString().trim(),
    milestone,
    owner: (row[3] || '').toString().trim(),
    startDate: excelDateToISO(row[4]),
    endDate: excelDateToISO(row[5]),
    processTime: (row[6] || '').toString().trim(),
    status: normalizeStatus(row[7]),
    timelineLock: (row[8] || '').toString().trim(),
    fy25StartDate: excelDateToISO(row[9]),
    fy25EndDate: excelDateToISO(row[10]),
    fy25ProcessTime: (row[11] || '').toString().trim(),
  };
}

/** Parse HighLevelMilestones row */
export function parseMilestoneRow(row: any[], index: number): Milestone | null {
  const category = (row[0] || '').toString().trim();
  if (!category || category === '' || category === 'Category') return null;

  return {
    id: `ms-${index}`,
    category,
    milestone: (row[1] || '').toString().trim(),
    corpDate: excelDateToISO(row[2]),
    corpStatus: normalizeStatus(row[3]),
    fedDate: excelDateToISO(row[4]),
    fedStatus: normalizeStatus(row[5]),
    comments: (row[6] || '').toString().trim(),
    fy25CorpDate: excelDateToISO(row[8]),
    fy25FedDate: excelDateToISO(row[9]),
  };
}

/** Parse RAID Log row */
export function parseRaidRow(row: any[], index: number): RaidEntry | null {
  const desc = (row[1] || '').toString().trim();
  if (!desc) return null;

  return {
    id: `raid-${index}`,
    date: excelDateToISO(row[0]) || '',
    description: desc,
    mitigation: (row[2] || '').toString().trim(),
    nextSteps: (row[3] || '').toString().trim(),
    owner: (row[4] || '').toString().trim(),
    dueDate: excelDateToISO(row[5]),
  };
}

/** Normalize status values from Excel to our canonical set */
function normalizeStatus(value: any): any {
  if (!value) return 'Not Started';
  const s = value.toString().trim().toLowerCase();
  if (s === 'completed' || s === 'complete' || s === 'done') return 'Completed';
  if (s === 'in progress' || s === 'in-progress' || s === 'started') return 'In Progress';
  if (s === 'blocked' || s === 'on hold') return 'Blocked';
  if (s === 'n/a' || s === 'na' || s === 'not applicable') return 'N/A';
  if (s === 'not started' || s === 'not set' || s === '') return 'Not Started';
  return value.toString().trim();
}

/** Extract all unique owners/DRIs from steps for user mapping */
export function extractOwners(steps: RolloverStep[]): Map<string, string[]> {
  const ownerMap = new Map<string, string[]>();
  for (const step of steps) {
    const people = [step.wwicPoc, step.fedPoc, step.engineeringDri, step.engineeringLead]
      .filter(p => p && p !== '-' && p !== '');
    for (const person of people) {
      // Handle "Person1, Person2" entries
      const names = person.split(/[,&]/).map(n => n.trim()).filter(n => n);
      for (const name of names) {
        if (!ownerMap.has(name)) ownerMap.set(name, []);
        ownerMap.get(name)!.push(step.id);
      }
    }
  }
  return ownerMap;
}
