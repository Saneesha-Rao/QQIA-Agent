const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const pptx = new pptxgen();

const C = {
  purple: '6264A7', darkPurple: '4B4D8A', dark: '252423', white: 'FFFFFF',
  gray: 'F5F5F5', green: '13A10E', red: 'D13438', orange: 'F7630C',
  blue: '0078D4', muted: '8A8886', light: 'E8E6E3',
};
const FONT = 'Segoe UI';

pptx.author = 'Saneesha';
pptx.title = 'QQIA Agent — Building an AI Agent in Enterprise';
pptx.subject = 'FY27 Mint Rollover';

// ============================================================
// SLIDE 1 — TITLE
// ============================================================
let s = pptx.addSlide();
s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.purple } });
s.addText('QQIA Agent', { x: 0.8, y: 1.2, w: 8.5, h: 1.1, fontSize: 48, bold: true, color: C.white, fontFace: FONT });
s.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 2.45, w: 1.6, h: 0.05, fill: { color: C.white } });
s.addText('Building an AI Agent for FY27 Mint Rollover\nLearnings on Tools, Platforms & Enterprise Constraints', {
  x: 0.8, y: 2.7, w: 8.5, h: 0.9, fontSize: 19, color: 'D0D0FF', fontFace: FONT, lineSpacingMultiple: 1.3,
});
s.addText('Seller Incentives  ·  April 2026', {
  x: 0.8, y: 4.6, w: 8.5, h: 0.4, fontSize: 14, color: 'A0A0D0', fontFace: FONT,
});

// Helper
function hdr(slide, title, subtitle) {
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: 1.05, fill: { color: C.purple } });
  slide.addText(title, { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 24, bold: true, color: C.white, fontFace: FONT });
  if (subtitle) slide.addText(subtitle, { x: 0.5, y: 0.57, w: 9, h: 0.33, fontSize: 13, color: 'D0D0FF', fontFace: FONT });
}

// ============================================================
// SLIDE 2 — THE CHALLENGE (story hook — team framing)
// ============================================================
s = pptx.addSlide();
s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.dark } });

s.addText('"181 activities. 17 workstreams. 29 people.\nOne shared Excel file as the source of truth."', {
  x: 0.8, y: 0.7, w: 8.5, h: 1.2, fontSize: 22, italic: true, color: C.white, fontFace: FONT, lineSpacingMultiple: 1.4,
});
s.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 2.1, w: 1.2, h: 0.04, fill: { color: C.purple } });

const challenges = [
  'No real-time visibility — you have to open the spreadsheet',
  'No dependency tracking — blocked items go unnoticed',
  'No alerts — overdue steps discovered in weekly meetings',
  'Leadership wants dashboards, not raw Excel rows',
];
challenges.forEach((c, i) => {
  s.addText(`→  ${c}`, { x: 0.8, y: 2.35 + i * 0.5, w: 8.5, h: 0.45, fontSize: 16, color: 'C0C0C0', fontFace: FONT });
});

s.addText('Question: Could an intelligent Teams agent solve this — reading the Excel, answering questions, tracking dependencies, and keeping everyone on the same page?', {
  x: 0.8, y: 4.6, w: 8.5, h: 0.65, fontSize: 14, bold: true, color: C.white, fontFace: FONT, lineSpacingMultiple: 1.25,
});

// ============================================================
// SLIDE 3 — WHAT WE BUILT
// ============================================================
s = pptx.addSlide();
hdr(s, 'What We Built', 'TypeScript Bot Framework SDK v4  ·  28 source files  ·  181 steps tracked');

const built = [
  ['💬', 'Natural Language', '"Update step 1.A to completed"\n"Show me overdue items"\n"Tasks for Pragya"'],
  ['📊', 'Smart Dashboards', 'Progress by workstream, blockers,\ncritical path, overdue alerts —\nall as Adaptive Cards in Teams'],
  ['🔗', 'Dependency Engine', 'Completing a step auto-detects\nunblocked downstream work.\nDAG-based critical path analysis.'],
  ['📥', 'Excel Sync', 'Reads from the source Excel.\nWrites back status changes.\nMultiple fallback strategies.'],
];

built.forEach((b, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.8;
  const y = 1.25 + row * 1.75;
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x, y, w: 4.5, h: 1.55, fill: { color: C.gray }, rectRadius: 0.08 });
  s.addText(b[0], { x: x + 0.15, y: y + 0.1, w: 0.5, h: 0.5, fontSize: 22 });
  s.addText(b[1], { x: x + 0.6, y: y + 0.1, w: 3.6, h: 0.35, fontSize: 14, bold: true, color: C.purple, fontFace: FONT });
  s.addText(b[2], { x: x + 0.6, y: y + 0.48, w: 3.7, h: 1.0, fontSize: 11, color: C.muted, fontFace: FONT, lineSpacingMultiple: 1.25 });
});

s.addText('15+ commands  ·  Adaptive Cards  ·  Typo correction  ·  Dual-track (Corp & Fed)  ·  Proactive alerts', {
  x: 0.5, y: 4.85, w: 9.0, h: 0.35, fontSize: 11.5, color: C.muted, fontFace: FONT, align: 'center',
});

// ============================================================
// SLIDE 4 — THE JOURNEY TIMELINE (from v1, with story framing)
// ============================================================
s = pptx.addSlide();
hdr(s, 'The Journey — What We Tried', 'Every approach, what happened, and where it landed');

const timeline = [
  { p: '1', label: 'Build the Bot', detail: 'TypeScript, Bot Framework, Excel parsing, 28 files', status: '✅ Success', color: C.green },
  { p: '2', label: 'Azure Deployment', detail: 'Bicep, App Service, Bot Service, GitHub Actions CI/CD', status: '❌ Blocked', color: C.red },
  { p: '3', label: 'Graph API Sync', detail: 'SharePoint REST API for Excel read/write', status: '❌ Blocked', color: C.red },
  { p: '4', label: 'Power Automate', detail: 'HTTP triggers + Office Script + Excel connector', status: '⚠️ Partial', color: C.orange },
  { p: '5', label: 'Excel COM Automation', detail: 'PowerShell 5.1 + running Excel instance', status: '✅ Works locally', color: C.green },
  { p: '6', label: 'Copilot Studio', detail: 'Knowledge sources, Dataverse, Agent Flows, Tools', status: '⚠️ In Progress', color: C.orange },
  { p: '7', label: 'GitHub Codespaces', detail: 'Free cloud hosting with public URL', status: '✅ Works', color: C.green },
  { p: '8', label: 'Teams Tab (Web UI)', detail: 'Chat interface embedded as a Teams Tab', status: '✅ Live', color: C.green },
];

timeline.forEach((t, i) => {
  const y = 1.2 + i * 0.5;
  s.addShape(pptx.shapes.OVAL, { x: 0.4, y: y + 0.06, w: 0.3, h: 0.3, fill: { color: C.purple } });
  s.addText(t.p, { x: 0.4, y: y + 0.06, w: 0.3, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: FONT });
  if (i < timeline.length - 1) s.addShape(pptx.shapes.LINE, { x: 0.55, y: y + 0.36, w: 0, h: 0.2, line: { color: C.purple, width: 1.5 } });
  s.addText(t.label, { x: 0.85, y, w: 2.65, h: 0.4, fontSize: 12.5, bold: true, color: C.dark, fontFace: FONT, valign: 'middle' });
  s.addText(t.detail, { x: 3.5, y, w: 4.2, h: 0.4, fontSize: 11, color: C.muted, fontFace: FONT, valign: 'middle' });
  s.addText(t.status, { x: 7.8, y, w: 1.8, h: 0.4, fontSize: 11, bold: true, color: t.color, fontFace: FONT, valign: 'middle', align: 'right' });
});

// ============================================================
// SLIDE 5 — THE WALL (story conflict — educational)
// ============================================================
s = pptx.addSlide();
hdr(s, 'Enterprise Compliance: What Blocked Us', 'These aren\'t bugs — they\'re intentional security measures');

const walls = [
  ['Conditional Access', 'Token protection policy', 'Blocks az login, Teams Toolkit, any programmatic Azure auth', 'Azure, TTK'],
  ['DLP Policy', 'Personal Developer env', 'Blocks HTTP triggers in Power Automate, suspends flows', 'Power Automate'],
  ['App Registration', 'Secret/cert restriction', 'Cannot create client secrets or upload certificates', 'Graph API'],
  ['OneDrive Sync', 'File lock (EBUSY)', 'Sync daemon locks Excel — Node.js writes fail', 'Local write-back'],
  ['Corporate Proxy', 'CDN restrictions', 'alcdn.msauth.net (MSAL) blocked', 'MSAL.js'],
  ['Tenant Settings', 'Feature disabled', 'Outgoing Webhooks disabled at org level', 'Webhooks'],
];

s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 1.18, w: 9.2, h: 0.38, fill: { color: C.darkPurple } });
s.addText('Policy', { x: 0.5, y: 1.18, w: 1.8, h: 0.38, fontSize: 11, bold: true, color: C.white, fontFace: FONT, valign: 'middle' });
s.addText('Specific Restriction', { x: 2.3, y: 1.18, w: 2.0, h: 0.38, fontSize: 11, bold: true, color: C.white, fontFace: FONT, valign: 'middle' });
s.addText('Impact', { x: 4.3, y: 1.18, w: 4.0, h: 0.38, fontSize: 11, bold: true, color: C.white, fontFace: FONT, valign: 'middle' });
s.addText('Blocked', { x: 8.3, y: 1.18, w: 1.3, h: 0.38, fontSize: 11, bold: true, color: C.white, fontFace: FONT, valign: 'middle', align: 'center' });

walls.forEach((w, i) => {
  const y = 1.56 + i * 0.52;
  const bg = i % 2 === 0 ? C.gray : C.white;
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.52, fill: { color: bg } });
  s.addText(w[0], { x: 0.5, y, w: 1.8, h: 0.52, fontSize: 10.5, bold: true, color: C.dark, fontFace: FONT, valign: 'middle' });
  s.addText(w[1], { x: 2.3, y, w: 2.0, h: 0.52, fontSize: 10, color: C.muted, fontFace: FONT, valign: 'middle' });
  s.addText(w[2], { x: 4.3, y, w: 4.0, h: 0.52, fontSize: 10, color: C.muted, fontFace: FONT, valign: 'middle' });
  s.addText(w[3], { x: 8.3, y, w: 1.3, h: 0.52, fontSize: 10, bold: true, color: C.red, fontFace: FONT, valign: 'middle', align: 'center' });
});

s.addText('💡 Understanding these policies upfront is critical. Most teams discover them mid-project — costing weeks of rework.', {
  x: 0.4, y: 4.8, w: 9.2, h: 0.45, fontSize: 12, italic: true, color: C.blue, fontFace: FONT,
});

// ============================================================
// SLIDE 6 — THE PIVOTS (story resolution — team language)
// ============================================================
s = pptx.addSlide();
hdr(s, 'How We Adapted', 'Every blocker led to a creative workaround');

const pivots = [
  { problem: 'Can\'t deploy to Azure?', solution: 'GitHub Codespaces', detail: 'Free cloud hosting with a public URL. The bot runs in a Codespace — no Azure subscription needed. ~60 hrs/month on free tier.', color: C.blue },
  { problem: 'Can\'t register a Teams bot?', solution: 'Teams Tab (Web UI)', detail: 'A full chat interface built in HTML/JS, added as a Website Tab in the Teams channel. Looks and feels native to Teams.', color: C.purple },
  { problem: 'Can\'t sync Excel via cloud APIs?', solution: 'Excel COM + Office Scripts', detail: 'PowerShell 5.1 talks to the running Excel instance locally. Office Scripts call the bot\'s JSON API. Download button as fallback.', color: C.green },
  { problem: 'Copilot Studio can\'t query tables?', solution: 'Excel Online Connector Tool', detail: '"Get a row" tool does precise lookups by step ID. Knowledge source handles general Q&A. Best of both worlds.', color: C.orange },
];

pivots.forEach((p, i) => {
  const y = 1.15 + i * 0.9;
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.78, fill: { color: C.gray }, rectRadius: 0.06 });
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.4, y, w: 0.08, h: 0.78, fill: { color: p.color }, rectRadius: 0.04 });
  s.addText(p.problem, { x: 0.7, y, w: 2.6, h: 0.3, fontSize: 11, italic: true, color: C.muted, fontFace: FONT });
  s.addText(p.solution, { x: 0.7, y: y + 0.3, w: 2.6, h: 0.4, fontSize: 15, bold: true, color: p.color, fontFace: FONT });
  s.addText(p.detail, { x: 3.5, y: y + 0.05, w: 5.9, h: 0.68, fontSize: 11.5, color: C.dark, fontFace: FONT, valign: 'middle', lineSpacingMultiple: 1.2 });
});

s.addText('Takeaway: Standard paths may be blocked — but there\'s almost always an alternative if the team stays flexible.', {
  x: 0.4, y: 4.8, w: 9.2, h: 0.4, fontSize: 12, italic: true, color: C.green, fontFace: FONT,
});

// ============================================================
// SLIDE 7 — PLATFORM DEEP-DIVES (detail from v1, condensed)
// ============================================================
s = pptx.addSlide();
hdr(s, 'Platform-by-Platform: What We Learned');

// Azure
s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.3, y: 1.15, w: 4.55, h: 1.55, fill: { color: 'FDE7E9' }, rectRadius: 0.06 });
s.addText('☁️  Azure Deployment', { x: 0.5, y: 1.2, w: 4.1, h: 0.3, fontSize: 13, bold: true, color: C.red, fontFace: FONT });
s.addText('Tried: az login, Teams Toolkit, Bicep, GitHub Actions\nResult: All blocked by Conditional Access token protection\nLesson: Verify Azure auth works before building CI/CD', {
  x: 0.5, y: 1.55, w: 4.2, h: 1.0, fontSize: 10.5, color: C.dark, fontFace: FONT, lineSpacingMultiple: 1.35 });

// Graph API
s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 5.15, y: 1.15, w: 4.55, h: 1.55, fill: { color: 'FDE7E9' }, rectRadius: 0.06 });
s.addText('🔑  Graph API', { x: 5.35, y: 1.2, w: 4.1, h: 0.3, fontSize: 13, bold: true, color: C.red, fontFace: FONT });
s.addText('Tried: Client secret, certificate, MSAL.js (SPA PKCE)\nResult: No secrets allowed; CDN blocked; admin consent needed\nLesson: Graph requires IT involvement in locked-down tenants', {
  x: 5.35, y: 1.55, w: 4.2, h: 1.0, fontSize: 10.5, color: C.dark, fontFace: FONT, lineSpacingMultiple: 1.35 });

// Power Automate
s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.3, y: 2.85, w: 4.55, h: 1.55, fill: { color: 'FFF3E0' }, rectRadius: 0.06 });
s.addText('⚡  Power Automate', { x: 0.5, y: 2.9, w: 4.1, h: 0.3, fontSize: 13, bold: true, color: C.orange, fontFace: FONT });
s.addText('✅ Excel Online connector reads/writes rows perfectly\n✅ Office Script runs and returns 181 rows\n❌ HTTP triggers suspended by DLP (Personal Developer env)\nLesson: Excel connector = great; HTTP trigger needs team env', {
  x: 0.5, y: 3.25, w: 4.2, h: 1.05, fontSize: 10.5, color: C.dark, fontFace: FONT, lineSpacingMultiple: 1.3 });

// Copilot Studio
s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 5.15, y: 2.85, w: 4.55, h: 1.55, fill: { color: 'FFF3E0' }, rectRadius: 0.06 });
s.addText('🤖  Copilot Studio', { x: 5.35, y: 2.9, w: 4.1, h: 0.3, fontSize: 13, bold: true, color: C.orange, fontFace: FONT });
s.addText('✅ Agent created, Excel knowledge "Ready"\n⚠️ Knowledge source can\'t do structured row lookups\n⚠️ "Get a row" tool configured but routing needs work\nLesson: Use Topics + connector tools for tabular data', {
  x: 5.35, y: 3.25, w: 4.2, h: 1.05, fontSize: 10.5, color: C.dark, fontFace: FONT, lineSpacingMultiple: 1.3 });

s.addText('Detailed notes and code available in the GitHub repo: Saneesha-Rao/QQIA-Agent', {
  x: 0.4, y: 4.6, w: 9.2, h: 0.35, fontSize: 11, italic: true, color: C.muted, fontFace: FONT, align: 'center',
});

// ============================================================
// SLIDE 8 — SCORECARD (from v2)
// ============================================================
s = pptx.addSlide();
hdr(s, 'The Scorecard', 'Where each platform landed');

const scores = [
  { tool: 'TypeScript Bot (custom)', status: '✅ Built & working', note: '15+ commands, dependency DAG, dashboards', c: C.green },
  { tool: 'Azure Deployment', status: '❌ Blocked', note: 'Conditional Access, no subscription', c: C.red },
  { tool: 'Graph API', status: '❌ Blocked', note: 'No app secrets, admin consent required', c: C.red },
  { tool: 'Power Automate', status: '⚠️ Partial', note: 'Excel connector works; HTTP triggers DLP-blocked', c: C.orange },
  { tool: 'Excel COM Sync', status: '✅ Works locally', note: 'PowerShell 5.1 + running Excel instance', c: C.green },
  { tool: 'Copilot Studio', status: '⚠️ In progress', note: 'Excel tool configured, needs topic routing', c: C.orange },
  { tool: 'GitHub Codespaces', status: '✅ Works', note: 'Free cloud hosting, public URL', c: C.green },
  { tool: 'Teams Tab (Web UI)', status: '✅ Live in Teams', note: 'Chat interface accessible to the team', c: C.green },
];

s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 9.2, h: 0.38, fill: { color: C.darkPurple } });
s.addText('Platform / Tool', { x: 0.5, y: 1.15, w: 3.0, h: 0.38, fontSize: 12, bold: true, color: C.white, fontFace: FONT, valign: 'middle' });
s.addText('Status', { x: 3.5, y: 1.15, w: 1.8, h: 0.38, fontSize: 12, bold: true, color: C.white, fontFace: FONT, valign: 'middle' });
s.addText('Detail', { x: 5.3, y: 1.15, w: 4.3, h: 0.38, fontSize: 12, bold: true, color: C.white, fontFace: FONT, valign: 'middle' });

scores.forEach((r, i) => {
  const y = 1.53 + i * 0.42;
  const bg = i % 2 === 0 ? C.gray : C.white;
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.42, fill: { color: bg } });
  s.addText(r.tool, { x: 0.5, y, w: 3.0, h: 0.42, fontSize: 12, bold: true, color: C.dark, fontFace: FONT, valign: 'middle' });
  s.addText(r.status, { x: 3.5, y, w: 1.8, h: 0.42, fontSize: 11, bold: true, color: r.c, fontFace: FONT, valign: 'middle' });
  s.addText(r.note, { x: 5.3, y, w: 4.3, h: 0.42, fontSize: 11, color: C.muted, fontFace: FONT, valign: 'middle' });
});

// ============================================================
// SLIDE 9 — KEY LEARNINGS (team-framed)
// ============================================================
s = pptx.addSlide();
hdr(s, 'Key Learnings for the Team', 'What this project taught us about building agents in enterprise');

const learnings = [
  {
    num: '01', title: 'Check compliance before writing code',
    text: 'Azure access, DLP policies, app registration restrictions, admin consent — 10 minutes with IT upfront saves weeks of rework.',
    color: C.blue,
  },
  {
    num: '02', title: 'Design for fallbacks from day one',
    text: 'This bot has 4 sync paths: COM → Power Automate → Graph → local. In enterprise, assume the primary path will be blocked.',
    color: C.green,
  },
  {
    num: '03', title: 'Copilot Studio ≠ structured data queries',
    text: 'Knowledge sources use semantic search — great for documents, poor for row lookups. Pair with Topics + connector tools.',
    color: C.orange,
  },
  {
    num: '04', title: 'Personal vs Team environments matter',
    text: 'Power Platform DLP varies by environment. The "Personal Developer" env is the most restrictive. A team env may unblock everything.',
    color: C.purple,
  },
  {
    num: '05', title: 'Prototype fast, pivot faster',
    text: 'When "proper" paths are blocked, Codespaces + Teams Tab delivered a working experience to the team in hours — not weeks.',
    color: C.darkPurple,
  },
];

learnings.forEach((l, i) => {
  const y = 1.2 + i * 0.72;
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.6, fill: { color: C.gray }, rectRadius: 0.06 });
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.4, y, w: 0.07, h: 0.6, fill: { color: l.color }, rectRadius: 0.03 });
  s.addText(l.num, { x: 0.6, y, w: 0.5, h: 0.6, fontSize: 18, bold: true, color: l.color, fontFace: FONT, valign: 'middle' });
  s.addText(l.title, { x: 1.15, y, w: 3.0, h: 0.6, fontSize: 13, bold: true, color: C.dark, fontFace: FONT, valign: 'middle' });
  s.addText(l.text, { x: 4.2, y, w: 5.2, h: 0.6, fontSize: 11, color: C.muted, fontFace: FONT, valign: 'middle', lineSpacingMultiple: 1.2 });
});

// ============================================================
// SLIDE 10 — WHAT'S NEXT (actionable asks)
// ============================================================
s = pptx.addSlide();
hdr(s, 'Recommended Next Steps', 'What would unblock production deployment');

const nexts = [
  { icon: '🏢', text: 'Get a team Power Platform environment — unblocks Power Automate flows & Copilot Studio', priority: 'High' },
  { icon: '🔑', text: 'Request one app secret from IT — unlocks Graph API for full SharePoint Excel sync', priority: 'High' },
  { icon: '🤖', text: 'Finish Copilot Studio agent — Topics + Excel tool for queries + write-back', priority: 'Medium' },
  { icon: '☁️', text: 'Azure subscription — enables App Service + Bot Service for production hosting', priority: 'Medium' },
];

nexts.forEach((n, i) => {
  const y = 1.25 + i * 0.85;
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.7, fill: { color: C.gray }, rectRadius: 0.06 });
  s.addText(n.icon, { x: 0.55, y, w: 0.5, h: 0.7, fontSize: 22, align: 'center', valign: 'middle' });
  s.addText(n.text, { x: 1.15, y, w: 7.0, h: 0.7, fontSize: 14, color: C.dark, fontFace: FONT, valign: 'middle' });
  const pColor = n.priority === 'High' ? C.red : C.orange;
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 8.4, y: y + 0.2, w: 1.0, h: 0.3, fill: { color: pColor }, rectRadius: 0.15 });
  s.addText(n.priority, { x: 8.4, y: y + 0.2, w: 1.0, h: 0.3, fontSize: 10, bold: true, color: C.white, fontFace: FONT, align: 'center', valign: 'middle' });
});

s.addText('The agent code is built and tested. The infrastructure is the remaining bottleneck.', {
  x: 0.5, y: 4.8, w: 9.0, h: 0.4, fontSize: 14, bold: true, italic: true, color: C.purple, fontFace: FONT, align: 'center',
});

// ============================================================
// SLIDE 11 — CLOSING
// ============================================================
s = pptx.addSlide();
s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.purple } });
s.addText('The agent is built.\nThe constraints are mapped.\nThe path to production is clear.', {
  x: 0.8, y: 1.3, w: 8.5, h: 2.0, fontSize: 26, color: C.white, fontFace: FONT, lineSpacingMultiple: 1.6, align: 'center',
});
s.addShape(pptx.shapes.RECTANGLE, { x: 4.2, y: 3.4, w: 1.6, h: 0.04, fill: { color: C.white } });
s.addText('Questions & Discussion', { x: 0.8, y: 3.7, w: 8.5, h: 0.6, fontSize: 20, color: 'D0D0FF', fontFace: FONT, align: 'center' });
s.addText('GitHub: Saneesha-Rao/QQIA-Agent', { x: 0.8, y: 4.5, w: 8.5, h: 0.35, fontSize: 13, color: 'A0A0D0', fontFace: FONT, align: 'center' });

// ---- Generate ----
const outputPath = path.join(__dirname, '..', 'docs', 'QQIA-Agent-Learnings.pptx');
pptx.writeFile({ fileName: outputPath })
  .then(() => console.log(`✅ Presentation saved to: ${outputPath}`))
  .catch(err => console.error('Error:', err));
