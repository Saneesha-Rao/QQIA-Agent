import * as restify from 'restify';
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
} from 'botbuilder';
import * as dotenv from 'dotenv';
import { config } from './config/appConfig';
import { QQIABot } from './bot/qqiaBot';
import { DataService } from './services/dataService';
import { InMemoryDataService } from './services/inMemoryDataService';
import { DependencyEngine } from './services/dependencyEngine';
import { ExcelSyncService } from './services/excelSyncService';
import { NotificationService } from './services/notificationService';
import { GraphService } from './services/graphService';
import { SyncEngine } from './services/syncEngine';
import { ProactiveMessenger } from './services/proactiveMessenger';
import {
  parseRolloverRow,
  parseMilestoneRow,
} from './utils/excelParser';
import * as XLSX from 'xlsx';

dotenv.config();

// ---- Initialize Services ----
// Use in-memory store when Cosmos DB is not configured (local dev)
const useCosmosDB = !!(config.cosmos.endpoint && config.cosmos.key);
const dataService: DataService | InMemoryDataService = useCosmosDB
  ? new DataService()
  : new InMemoryDataService();
const dependencyEngine = new DependencyEngine();
const graphService = new GraphService();
const excelSync = new ExcelSyncService(dataService as any);
const notificationService = new NotificationService(dataService as any, dependencyEngine);

// ---- Bot Framework Setup ----
const isLocalDev = !config.bot.appId;
const botFrameworkAuth = new ConfigurationBotFrameworkAuthentication(
  isLocalDev
    ? {} // Empty config allows anonymous access for local Bot Emulator testing
    : {
        MicrosoftAppId: config.bot.appId,
        MicrosoftAppPassword: config.bot.appPassword,
        MicrosoftAppTenantId: config.bot.tenantId,
        MicrosoftAppType: 'SingleTenant',
      }
);

const adapter = new CloudAdapter(botFrameworkAuth);

// Error handler
adapter.onTurnError = async (context: TurnContext, error: Error) => {
  console.error(`[onTurnError] ${error.message}`, error.stack);
  await context.sendActivity('❌ The bot encountered an error. Please try again.');
};

// Initialize proactive messenger and sync engine (need adapter reference)
const proactiveMessenger = new ProactiveMessenger(adapter, config.bot.appId, dataService as any, notificationService);
const syncEngine = new SyncEngine(dataService as any, graphService, dependencyEngine, notificationService);

// Create the bot with all services
const bot = new QQIABot(dataService as any, dependencyEngine, excelSync, notificationService);

// ---- HTTP Server ----
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, async (context) => {
    // Save conversation reference for proactive messaging on every interaction
    if (context.activity.from) {
      proactiveMessenger.saveConversationReference(context.activity);
    }
    await bot.run(context);
  });
});

// Health check endpoint with status details
server.get('/api/health', async (req, res) => {
  let stepCount = 0;
  try {
    const steps = await dataService.getAllSteps();
    stepCount = steps.length;
  } catch { /* Cosmos unavailable */ }

  res.send(200, {
    status: 'healthy',
    timestamp: new Date().toISOString(),
    version: '1.0.0',
    registeredUsers: proactiveMessenger.getRegisteredUserCount(),
    trackedSteps: stepCount,
  });
});

// Manual sync trigger endpoint (webhook for Power Automate / Logic Apps)
server.post('/api/sync', async (req, res) => {
  try {
    const result = await syncEngine.runSync();
    res.send(200, result);
  } catch (err: any) {
    res.send(500, { error: err.message });
  }
});

// Webhook for external automation status updates (ADO, pipelines, etc.)
server.post('/api/automation/status', async (req, res) => {
  try {
    const { stepId, status, track, source, notes } = req.body || {};
    if (!stepId || !status) {
      res.send(400, { error: 'stepId and status are required' });
      return;
    }

    const field = track === 'Fed' ? 'fedStatus' : 'corpStatus';
    const updated = await dataService.updateStepStatus(
      stepId, field as any, status, source || 'automation', 'automation'
    );

    if (!updated) {
      res.send(404, { error: `Step ${stepId} not found` });
      return;
    }

    // Add notes if provided
    if (notes && updated.referenceNotes !== undefined) {
      updated.referenceNotes = updated.referenceNotes
        ? `${updated.referenceNotes}\n[Auto ${new Date().toISOString().split('T')[0]}]: ${notes}`
        : `[Auto ${new Date().toISOString().split('T')[0]}]: ${notes}`;
      await dataService.upsertStep(updated);
    }

    // Trigger dependency notifications if completed
    if (status === 'Completed') {
      await proactiveMessenger.deliverPredecessorNotifications(stepId, track || 'Corp');
    }

    res.send(200, { success: true, step: updated });
  } catch (err: any) {
    res.send(500, { error: err.message });
  }
});

// ---- Scheduled Tasks ----

/** Excel sync via Graph API (every 15 minutes) */
function startSyncScheduler() {
  setInterval(async () => {
    try {
      console.log(`[Sync] Starting bi-directional Excel sync...`);
      const result = await syncEngine.runSync();
      console.log(`[Sync] Complete: ${result.fromExcel} from Excel, ${result.toExcel} to Excel, ${result.conflicts.length} conflicts`);
    } catch (err: any) {
      console.error(`[Sync] Error: ${err.message}`);
      // Fallback to local sync
      try {
        const localResult = await excelSync.fullSync();
        console.log(`[Sync] Local fallback: ${localResult.fromExcel} from Excel, ${localResult.toExcel} to Excel`);
      } catch (localErr: any) {
        console.error(`[Sync] Local fallback also failed: ${localErr.message}`);
      }
    }
  }, config.notifications.syncIntervalMinutes * 60 * 1000);
}

/** Proactive notification delivery (every hour) */
function startNotificationScheduler() {
  setInterval(async () => {
    try {
      console.log(`[Notifications] Running notification check...`);
      const result = await proactiveMessenger.deliverNotifications();
      console.log(`[Notifications] Sent: ${result.sent}, Failed: ${result.failed}`);
    } catch (err: any) {
      console.error(`[Notifications] Error: ${err.message}`);
    }
  }, 60 * 60 * 1000);
}

/** Weekly digest (check every hour, send on Monday mornings) */
function startWeeklyDigestScheduler() {
  setInterval(async () => {
    const now = new Date();
    // Monday = 1, send between 8:00-8:59 AM
    if (now.getDay() === 1 && now.getHours() === 8) {
      try {
        console.log(`[WeeklyDigest] Generating Monday morning digest...`);
        const result = await proactiveMessenger.deliverWeeklyDigest();
        console.log(`[WeeklyDigest] Sent: ${result.sent}, Failed: ${result.failed}`);
      } catch (err: any) {
        console.error(`[WeeklyDigest] Error: ${err.message}`);
      }
    }
  }, 60 * 60 * 1000);
}

// ---- Startup ----
async function main() {
  console.log('🚀 QQIA Agent starting...');
  console.log(`   Environment: ${process.env.NODE_ENV || 'development'}`);
  console.log(`   Data store:  ${useCosmosDB ? 'Azure Cosmos DB' : 'In-Memory (local dev)'}`);

  // Initialize data store
  try {
    await dataService.initialize();
    console.log(useCosmosDB ? '✅ Cosmos DB initialized' : '✅ In-memory store initialized');
  } catch (err: any) {
    console.warn(`⚠️ Data store init failed (${err.message}).`);
  }

  // Initialize Graph API
  try {
    if (config.graph.clientId) {
      await graphService.initialize();
      console.log('✅ Graph API initialized');
    } else {
      console.warn('⚠️ Graph API credentials not configured. Using local Excel file only.');
    }
  } catch (err: any) {
    console.warn(`⚠️ Graph API init failed (${err.message}). Using local Excel fallback.`);
  }

  // Seed data from Excel into the data store on startup
  try {
    const excelPath = 'C:\\Users\\salingal\\OneDrive - Microsoft\\Seller Incentives\\QQIA\\FY27_Mint_RolloverTimeline.xlsx';
    const wb = XLSX.readFile(excelPath);

    // Import FY27_Rollover steps
    const rolloverSheet = wb.Sheets['FY27_Rollover'];
    if (rolloverSheet) {
      const rows: any[][] = XLSX.utils.sheet_to_json(rolloverSheet, { header: 1, defval: '' });
      let stepCount = 0;
      for (let i = 3; i < rows.length; i++) {
        const step = parseRolloverRow(rows[i], i);
        if (step) {
          await dataService.upsertStep(step);
          stepCount++;
        }
      }
      console.log(`📊 Loaded ${stepCount} rollover steps from Excel`);
    }

    // Import milestones
    const msSheet = wb.Sheets['HighLevelMilestones'];
    if (msSheet) {
      const msRows: any[][] = XLSX.utils.sheet_to_json(msSheet, { header: 1, defval: '' });
      let msCount = 0;
      for (let i = 2; i < msRows.length; i++) {
        const ms = parseMilestoneRow(msRows[i], i);
        if (ms) {
          await dataService.upsertMilestone(ms);
          msCount++;
        }
      }
      console.log(`🏁 Loaded ${msCount} milestones from Excel`);
    }

    // Build dependency graph
    const allSteps = await dataService.getAllSteps();
    dependencyEngine.buildGraph(allSteps);
    const validation = dependencyEngine.validateDAG();
    console.log(`🔗 Dependency graph: ${allSteps.length} nodes, DAG valid: ${validation.valid}`);
    if (!validation.valid) {
      console.warn(`⚠️ Cycles detected: ${validation.cycles.map(c => c.join('→')).join('; ')}`);
    }
  } catch (err: any) {
    console.warn(`⚠️ Excel seed failed: ${err.message}`);
  }

  // Start all schedulers
  startSyncScheduler();
  startNotificationScheduler();
  startWeeklyDigestScheduler();
  console.log('⏰ Schedulers started (sync: 15min, notifications: 1hr, digest: Mon 8AM)');

  // Start HTTP server
  server.listen(config.bot.port, () => {
    console.log(`\n✅ QQIA Agent listening on http://localhost:${config.bot.port}`);
    console.log(`   Bot endpoint:   POST /api/messages`);
    console.log(`   Health check:   GET  /api/health`);
    console.log(`   Manual sync:    POST /api/sync`);
    console.log(`   Auto-update:    POST /api/automation/status\n`);
  });
}

main().catch((err) => {
  console.error('Fatal error:', err);
  process.exit(1);
});
