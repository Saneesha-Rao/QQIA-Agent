const docx = require('docx');
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType } = docx;

function heading(text, level) {
  return new Paragraph({ heading: level, spacing: { before: 300, after: 100 }, children: [new TextRun({ text, bold: true })] });
}
function para(text) {
  return new Paragraph({ spacing: { after: 80 }, children: [new TextRun(text)] });
}
function boldPara(text) {
  return new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text, bold: true })] });
}
function bullet(text) {
  return new Paragraph({ bullet: { level: 0 }, spacing: { after: 40 }, children: [new TextRun(text)] });
}
function code(text) {
  return new Paragraph({ spacing: { after: 60 }, indent: { left: 720 }, children: [new TextRun({ text, font: 'Consolas', size: 20, color: '2E7D32' })] });
}
function tableRow(cells, isHeader) {
  return new TableRow({ children: cells.map(c => new TableCell({
    width: { size: Math.floor(100 / cells.length), type: WidthType.PERCENTAGE },
    children: [new Paragraph({ children: [new TextRun({ text: c, bold: !!isHeader, size: isHeader ? 22 : 20, color: isHeader ? 'FFFFFF' : '000000' })] })],
    shading: isHeader ? { fill: '1F4E79' } : undefined,
  }))});
}
function simpleTable(headers, rows) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [tableRow(headers, true), ...rows.map(r => tableRow(r, false))],
  });
}

const doc = new Document({
  styles: { default: { document: { run: { font: 'Segoe UI', size: 22 } } } },
  sections: [{
    properties: { page: { margin: { top: 1000, bottom: 1000, left: 1200, right: 1200 } } },
    children: [
      // Title
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: 'QQIA Agent', size: 48, bold: true, color: '1F4E79' })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 50 }, children: [new TextRun({ text: 'User Guide', size: 36, bold: true, color: '1F4E79' })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: 'FY27 Mint Rollover Tracking Bot for Microsoft Teams', size: 24, italics: true, color: '555555' })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: 'April 2026 | v1.0', size: 20, color: '888888' })] }),

      // What Is
      heading('What Is the QQIA Agent?', HeadingLevel.HEADING_1),
      para('The QQIA Agent is a Microsoft Teams bot that helps the FY27 Mint Rollover team track, update, and coordinate over 182 rollover steps across 15 workstreams. It works alongside the existing Excel tracker \u2014 any updates you make in Teams automatically sync back to the shared Excel file, and vice versa.'),
      para('You don\'t need to install anything. Just message the bot in Teams like you would a colleague.'),

      // Getting Started
      heading('Getting Started', HeadingLevel.HEADING_1),
      heading('Finding the Bot', HeadingLevel.HEADING_2),
      bullet('Open Microsoft Teams'),
      bullet('In the chat pane, search for "QQIA Agent"'),
      bullet('Start a 1:1 chat with the bot'),
      bullet('The bot will greet you and show available commands'),
      para('Tip: You can also @mention the bot in a team channel to use it publicly.'),
      heading('First Message', HeadingLevel.HEADING_2),
      para('Type help to see all available commands:'),
      code('help'),

      // Viewing Status
      heading('Viewing Status & Dashboards', HeadingLevel.HEADING_1),

      heading('Overall Dashboard', HeadingLevel.HEADING_2),
      para('See a visual overview of the entire rollover \u2014 progress bars, workstream breakdown, alerts:'),
      code('dashboard'),
      para('The dashboard card shows completed, in progress, blocked, and not started counts, per-workstream progress, overdue alerts, and quick-action buttons.'),

      heading('My Tasks', HeadingLevel.HEADING_2),
      para('See all steps assigned to you (matched by your Teams display name):'),
      code('my tasks'),

      heading('Check a Specific Step', HeadingLevel.HEADING_2),
      para('Look up any step by its ID (e.g., 1.A, 3.B.1, 7.D):'),
      code('status 1.A'),
      para('This shows the step\'s description, dates, Corp/Fed status, DRI, dependencies, and notes.'),

      heading('View a Workstream', HeadingLevel.HEADING_2),
      para('See all steps within a workstream:'),
      code('workstream System Rollover'),
      code('workstream Orchestration'),
      code('ws Quota Issuance'),
      para('Shortcut: You can use "ws" instead of "workstream".'),

      heading('See Someone Else\'s Tasks', HeadingLevel.HEADING_2),
      para('Check what steps are assigned to a specific person:'),
      code('tasks for Jim R'),
      code('owner Saneesha'),
      para('Use the name as it appears in the Excel tracker (WWIC POC, Fed POC, or Engineering DRI columns).'),

      // Blockers
      heading('Blockers, Overdue & Critical Path', HeadingLevel.HEADING_1),
      heading('View All Blocked Steps', HeadingLevel.HEADING_2),
      code('blockers'),
      heading('View Overdue Steps', HeadingLevel.HEADING_2),
      para('Steps past their end date that aren\'t yet completed:'),
      code('overdue'),
      heading('Critical Path', HeadingLevel.HEADING_2),
      para('See the longest chain of dependent steps \u2014 the sequence that determines the earliest the rollover can complete:'),
      code('critical path'),

      // Updating
      heading('Updating Step Status', HeadingLevel.HEADING_1),

      heading('Corp Track Updates', HeadingLevel.HEADING_2),
      simpleTable(['Action', 'Command'], [
        ['Mark complete', 'update 1.A completed'],
        ['Mark in progress', 'update 1.A in progress'],
        ['Mark blocked', 'update 1.A blocked'],
        ['Reset to not started', 'update 1.A not started'],
      ]),

      heading('Fed Track Updates', HeadingLevel.HEADING_2),
      para('Prefix any update with "fed" to update the Fed status instead of Corp:'),
      code('fed update 1.A completed'),
      code('fed mark 3.B in progress'),

      heading('Add a Note to a Step', HeadingLevel.HEADING_2),
      para('Attach a comment or context to any step:'),
      code('note 1.A Waiting on ADO work item approval from Finance'),
      code('add note 3.B.1 Discussed with Jim - will complete by Thursday'),

      // Leadership
      heading('Leadership & Summary Views', HeadingLevel.HEADING_1),
      para('Get a high-level overview suitable for leadership updates:'),
      code('summary'),
      code('exec summary'),
      code('leadership update'),
      para('Provides: overall completion %, blocked/overdue counts, workstream-by-workstream progress, key risks, and upcoming deadlines.'),

      // Corp vs Fed
      heading('Corp vs Fed Tracking', HeadingLevel.HEADING_1),
      para('Every step has independent Corp and Fed statuses. By default, commands show Corp track data.'),
      simpleTable(['To See...', 'Command'], [
        ['Corp dashboard', 'dashboard'],
        ['Fed dashboard', 'fed   or   fed status'],
        ['Update Corp status', 'update 1.A completed'],
        ['Update Fed status', 'fed update 1.A completed'],
      ]),

      // Notifications
      heading('Automatic Notifications', HeadingLevel.HEADING_1),
      para('The bot proactively sends you messages \u2014 you don\'t need to check manually:'),
      simpleTable(['Notification', 'When', 'Who Gets It'], [
        ['Deadline approaching', '3 days and 1 day before due', 'Step DRI/POC'],
        ['Overdue alert', 'Step passes its due date', 'DRI/POC + PM'],
        ['Predecessor completed', 'Blocking step finishes', 'DRI of unblocked step'],
        ['Step unblocked', 'All predecessors are done', 'Step DRI/POC'],
        ['Weekly digest', 'Every Monday at 8:00 AM', 'All active DRIs/POCs'],
        ['Escalation', 'Step overdue by 3+ days', 'PM + Leadership'],
      ]),
      para('Notifications appear as 1:1 messages from the bot. No action needed to opt in \u2014 if you\'re a DRI or POC on any step, you\'ll get relevant alerts automatically.'),

      // Excel Sync
      heading('Excel Sync', HeadingLevel.HEADING_1),
      para('The QQIA Agent stays synchronized with the shared FY27_Mint_RolloverTimeline.xlsx file on SharePoint:'),
      bullet('Auto-sync every 15 minutes \u2014 changes in Teams appear in Excel and vice versa'),
      bullet('You can still update Excel directly \u2014 the bot will pick up your changes'),
      bullet('No data loss \u2014 if both are edited, the most recent change wins and both versions are logged'),
      para('Trigger a manual sync anytime:'),
      code('sync'),

      // Natural Language
      heading('Natural Language Queries', HeadingLevel.HEADING_1),
      para('Don\'t remember the exact command? The bot understands natural language too:'),
      simpleTable(['What You Type', 'What Happens'], [
        ['How many steps are done?', 'Shows summary counts'],
        ['What\'s due this week?', 'Lists steps due in the next 7 days'],
        ['Who owns step 3.B?', 'Shows step details including DRI'],
      ]),

      // Quick Reference
      heading('Quick Reference Card', HeadingLevel.HEADING_1),
      simpleTable(['Command', 'Description'], [
        ['help', 'Show all commands'],
        ['dashboard', 'Overall rollover progress'],
        ['my tasks', 'Steps assigned to you'],
        ['status 1.A', 'View step 1.A details'],
        ['update 1.A completed', 'Mark 1.A as completed (Corp)'],
        ['update 1.A in progress', 'Mark 1.A as in progress'],
        ['update 1.A blocked', 'Mark 1.A as blocked'],
        ['fed update 1.A completed', 'Mark 1.A as completed (Fed)'],
        ['note 1.A your note here', 'Add a note to step 1.A'],
        ['workstream Orchestration', 'View Orchestration workstream'],
        ['tasks for Jim R', 'View Jim R\'s steps'],
        ['blockers', 'All blocked steps'],
        ['overdue', 'All overdue steps'],
        ['critical path', 'Critical dependency chain'],
        ['summary', 'Executive / leadership summary'],
        ['fed', 'Fed track dashboard'],
        ['sync', 'Trigger Excel sync now'],
      ]),

      // FAQ
      heading('Frequently Asked Questions', HeadingLevel.HEADING_1),

      boldPara('Q: Do I need to install anything?'),
      para('No. The bot runs in Microsoft Teams. Just search for "QQIA Agent" and start chatting.'),

      boldPara('Q: Will my updates show up in the Excel tracker?'),
      para('Yes. The bot syncs with the Excel file every 15 minutes. Your updates appear in the shared spreadsheet automatically.'),

      boldPara('Q: Can I still update the Excel file directly?'),
      para('Yes. The sync is bi-directional. Update either Teams or Excel \u2014 both stay current.'),

      boldPara('Q: What if the bot doesn\'t recognize my name for "my tasks"?'),
      para('Your Teams display name must match the name in the Excel tracker columns (WWIC POC, Fed POC, or Engineering DRI). Try using "tasks for [exact name]" with the name from the tracker.'),

      boldPara('Q: Can I use the bot in a team channel?'),
      para('Yes. @mention the bot in any channel: @QQIA Agent dashboard. The response will be visible to everyone.'),

      boldPara('Q: What step ID format should I use?'),
      para('Step IDs follow the format from the tracker: 1.A, 3.B, 7.D.1, etc. You can find them by browsing a workstream.'),

      boldPara('Q: Who can update step statuses?'),
      para('DRIs and POCs can update their own steps. PMs can update any step. Leadership has view-only access.'),

      boldPara('Q: What if I make a mistake in an update?'),
      para('All changes are audit-logged. Simply update the step again to the correct status, or ask a PM to correct it.'),

      // For PMs
      heading('For PMs & Admins', HeadingLevel.HEADING_1),
      bullet('PMs can update any step, not just their own'),
      bullet('PMs can override dependency blocks when needed'),
      bullet('External systems (ADO, pipelines) can auto-update via webhooks'),
      bullet('See the Power Automate Integration Guide for automation setup'),

      // Support
      heading('Support', HeadingLevel.HEADING_1),
      bullet('Bot Admin: Saneesha (salingal)'),
      bullet('Teams Channel: Post in the QQIA Rollover channel'),
      bullet('GitHub Issues: github.com/Saneesha-Rao/QQIA-Agent/issues'),

      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400 }, children: [new TextRun({ text: 'Last updated: April 2026 | QQIA Agent v1.0', size: 18, color: '888888', italics: true })] }),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('docs/QQIA-Agent-User-Guide.docx', buf);
  console.log('Word doc created: docs/QQIA-Agent-User-Guide.docx (' + buf.length + ' bytes)');
}).catch(err => {
  console.error('Error:', err.message);
});
