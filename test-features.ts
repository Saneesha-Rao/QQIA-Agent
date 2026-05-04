/**
 * Quick integration test for new NLP features.
 * Tests the bot's message handling logic without starting the HTTP server.
 */
import { QQIABot } from './src/bot/qqiaBot';
import { InMemoryDataService } from './src/services/inMemoryDataService';
import { DependencyEngine } from './src/services/dependencyEngine';
import { ExcelSyncService } from './src/services/excelSyncService';
import { NotificationService } from './src/services/notificationService';
import { RolloverStep } from './src/models/types';

// ---- Test helpers ----
const responses: string[] = [];
const cards: any[] = [];

function makeMockContext(text: string, userName = 'Test User'): any {
  responses.length = 0;
  cards.length = 0;
  return {
    activity: {
      text,
      from: { name: userName },
      conversation: { id: 'test-conv-1' },
      value: undefined,
    },
    sendActivity: async (msg: any) => {
      if (typeof msg === 'string') {
        responses.push(msg);
      } else if (msg?.attachments) {
        cards.push(msg.attachments[0]?.content);
        responses.push('[CARD]');
      } else {
        responses.push(JSON.stringify(msg));
      }
    },
  };
}

function assert(condition: boolean, testName: string) {
  if (condition) {
    console.log(`  ✅ ${testName}`);
  } else {
    console.log(`  ❌ ${testName}`);
    console.log(`     Responses: ${responses.join(' | ').substring(0, 200)}`);
  }
}

// ---- Setup ----
async function runTests() {
  const dataService = new InMemoryDataService();
  const depEngine = new DependencyEngine();
  const excelSync = new ExcelSyncService(dataService, 'none');
  const notifService = new NotificationService(dataService);

  // Seed test data
  const testSteps: Partial<RolloverStep>[] = [
    { id: '1.A', description: 'Plan Design Rollover', workstream: 'System Rollover', corpStatus: 'In Progress', corpEndDate: '2026-04-25', wwicPoc: 'Jim R', engineeringDri: '', fedPoc: '', fedStatus: 'Not Started', dependencies: [], corpStartDate: '2026-04-01', corpCompletedDate: null, fedStartDate: null, fedEndDate: null, setupValidation: '', engineeringDependent: 'SPM', engineeringLead: '', adoLink: '', referenceNotes: '', fy26CorpStart: null, fy26CorpEnd: null, fy26FedStart: null, fy26FedEnd: null, lastModified: '2026-04-23T10:00:00Z', lastModifiedBy: 'Jim R', lastModifiedSource: 'bot' },
    { id: '1.B', description: 'Role Excellence Rollover', workstream: 'System Rollover', corpStatus: 'Not Started', corpEndDate: '2026-05-01', wwicPoc: '', engineeringDri: '', fedPoc: '', fedStatus: 'Not Started', dependencies: ['1.A'], corpStartDate: '2026-04-15', corpCompletedDate: null, fedStartDate: null, fedEndDate: null, setupValidation: '', engineeringDependent: '', engineeringLead: '', adoLink: '', referenceNotes: '', fy26CorpStart: null, fy26CorpEnd: null, fy26FedStart: null, fy26FedEnd: null, lastModified: '2026-04-22T08:00:00Z', lastModifiedBy: 'system', lastModifiedSource: 'excel' },
    { id: '1.C', description: 'Data Management Rollover', workstream: 'MSC FY27', corpStatus: 'Completed', corpEndDate: '2026-04-10', wwicPoc: 'Sarah K', engineeringDri: 'Dev Team', fedPoc: 'Fed Lead', fedStatus: 'In Progress', dependencies: [], corpStartDate: '2026-03-15', corpCompletedDate: '2026-04-10', fedStartDate: '2026-04-01', fedEndDate: '2026-05-15', setupValidation: '', engineeringDependent: '', engineeringLead: '', adoLink: '', referenceNotes: '', fy26CorpStart: null, fy26CorpEnd: null, fy26FedStart: null, fy26FedEnd: null, lastModified: '2026-04-10T14:00:00Z', lastModifiedBy: 'Sarah K', lastModifiedSource: 'bot' },
    { id: '2.A', description: 'Territory Management Setup', workstream: 'Territory Mgmt', corpStatus: 'Blocked', corpEndDate: '2026-04-20', wwicPoc: 'Amit T', engineeringDri: 'Eng Lead', fedPoc: '', fedStatus: 'Not Started', dependencies: ['1.A'], corpStartDate: '2026-04-10', corpCompletedDate: null, fedStartDate: null, fedEndDate: null, setupValidation: '', engineeringDependent: 'SPM', engineeringLead: '', adoLink: '', referenceNotes: 'Waiting on SPM team', fy26CorpStart: null, fy26CorpEnd: null, fy26FedStart: null, fy26FedEnd: null, lastModified: '2026-04-23T09:00:00Z', lastModifiedBy: 'Amit T', lastModifiedSource: 'bot' },
  ];

  for (const s of testSteps) {
    await dataService.upsertStep(s as RolloverStep);
  }

  const bot = new QQIABot(dataService, depEngine, excelSync, notifService);

  // Access private handleMessage via any cast
  const handle = async (text: string, user = 'Test User') => {
    const ctx = makeMockContext(text, user);
    await (bot as any).handleMessage(ctx);
  };

  console.log('\n=== 1. GREETINGS & CHITCHAT ===');
  await handle('hi');
  assert(responses[0]?.includes('Hi') || responses[0]?.includes('hi'), 'hi → greeting');

  await handle('thanks');
  assert(responses[0]?.includes('welcome') || responses[0]?.includes('Welcome'), 'thanks → you\'re welcome');

  await handle('bye');
  assert(responses[0]?.includes('later') || responses[0]?.includes('See you'), 'bye → goodbye');

  console.log('\n=== 2. SYNONYM EXPANSION ===');
  await handle('show finished steps');
  // "finished" → "completed" via synonym, should match in NL handler or step query
  assert(!responses.join(' ').includes("didn't understand"), 'finished → completed synonym works');

  await handle('blockers');
  // Direct command should work with test data containing Blocked status
  assert(responses.some(r => r === '[CARD]' || r.includes('Blocked') || r.includes('blocked') || r.includes('No blocked')), 'blockers command works');

  console.log('\n=== 3. FUZZY COMMAND MATCHING ===');
  await handle('dashbord');
  // "dashbord" → "dashboard" via Levenshtein
  assert(responses.some(r => r === '[CARD]'), 'dashbord → dashboard (fuzzy match)');

  await handle('sumary');
  // "sumary" → "summary"
  assert(responses.some(r => r.includes('Executive Summary') || r.includes('Progress')), 'sumary → summary (fuzzy match)');

  console.log('\n=== 4. BATCH STATUS UPDATES ===');
  await handle('mark 1.A and 1.B as completed');
  assert(responses.some(r => r.includes('Batch update') || r.includes('1.A') || r === '[CARD]'), 'batch update 1.A, 1.B');

  // Reset statuses
  await dataService.updateStepStatus('1.A', 'corpStatus', 'In Progress', 'system', 'bot');
  await dataService.updateStepStatus('1.B', 'corpStatus', 'Not Started', 'system', 'bot');

  console.log('\n=== 5. NEGATIVE / INVERSE QUERIES ===');
  await handle("what hasn't started");
  assert(responses.some(r => r === '[CARD]' || r.includes('Not Started')), "hasn't started → shows not started steps");

  await handle('unassigned steps');
  assert(responses.some(r => r === '[CARD]' || r.includes('Unassigned') || r.includes('No Owner')), 'unassigned → shows steps without owner');

  await handle('show incomplete steps');
  assert(responses.some(r => r === '[CARD]' || r.includes('Incomplete')), 'incomplete → shows non-completed steps');

  console.log('\n=== 6. DATE-RANGE QUERIES ===');
  await handle('what changed today');
  assert(!responses.join(' ').includes("didn't understand"), 'what changed today → handled');

  await handle('due before may 1');
  assert(responses.some(r => r === '[CARD]') || !responses.join(' ').includes("didn't understand"), 'due before May 1 → date range query');

  await handle('due this month');
  assert(!responses.join(' ').includes("didn't understand"), 'due this month → handled');

  console.log('\n=== 7. WHAT CHANGED QUERIES ===');
  await handle('changes');
  assert(!responses.join(' ').includes("didn't understand"), 'changes → handled');

  await handle('who updated steps');
  assert(!responses.join(' ').includes("didn't understand"), 'who updated steps → handled');

  console.log('\n=== 8. SUGGESTED ACTIONS ===');
  await handle('dashboard');
  // Should have a card (dashboard) + another card (suggested actions)
  assert(cards.length >= 2, 'dashboard → includes suggested actions card');
  const lastCard = cards[cards.length - 1];
  assert(lastCard?.actions?.some((a: any) => a.data?.action === 'quick_action'), 'suggested actions card has quick_action buttons');

  await handle('blockers');
  const blockerCards = [...cards];
  assert(blockerCards.some(c => c?.actions?.some((a: any) => a.data?.text === 'dashboard')), 'after blockers → suggested action includes dashboard');

  console.log('\n=== 9. CLICKABLE STEP ROWS ===');
  await handle('upcoming');
  // Check that the list card rows have selectAction
  const listCard = cards.find(c => c?.body?.some((b: any) => b.type === 'ColumnSet' && b.selectAction));
  assert(!!listCard, 'step list card rows have selectAction (clickable)');

  console.log('\n=== 10. QUICK ACTION BUTTON CLICK ===');
  // Simulate clicking a suggested action button
  const quickCtx = makeMockContext('');
  quickCtx.activity.value = { action: 'quick_action', text: 'blockers' };
  quickCtx.activity.text = '';
  await (bot as any).handleMessage(quickCtx);
  assert(responses.some(r => r === '[CARD]' || r.includes('Blocked') || r.includes('No blocked')), 'quick_action button click routes correctly');

  console.log('\n=== 11. VIEW STEP CLICK FROM LIST ===');
  const viewCtx = makeMockContext('');
  viewCtx.activity.value = { action: 'view_step', stepId: '1.A' };
  viewCtx.activity.text = '';
  await (bot as any).handleMessage(viewCtx);
  assert(responses.some(r => r === '[CARD]'), 'view_step action shows step detail card');
  const detailCard = cards.find(c => c?.actions?.some((a: any) => a.data?.action === 'update_status'));
  assert(!!detailCard, 'step detail card has Mark Complete/In Progress/Blocked buttons');

  console.log('\n=== DONE ===\n');
}

runTests().catch(err => {
  console.error('Test error:', err);
  process.exit(1);
});
