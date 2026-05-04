import express from 'express';
import helmet from 'helmet';
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
import { PowerAutomateSyncService } from './services/powerAutomateSyncService';
import { ExcelComSyncService } from './services/excelComSyncService';
import { NotificationService } from './services/notificationService';
import { GraphService } from './services/graphService';
import { DelegatedGraphService } from './services/delegatedGraphService';
import { SyncEngine } from './services/syncEngine';
import { ProactiveMessenger } from './services/proactiveMessenger';
import { WebhookHandler } from './webhookHandler';
import { AnalyticsService } from './services/analyticsService';
import * as XLSX from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';

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
const paSyncService = new PowerAutomateSyncService(dataService as any);
const comSync = new ExcelComSyncService(dataService as any);
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
const analyticsService = new AnalyticsService();

// Create the bot with all services (pass PA sync and COM sync for SharePoint write-back)
const bot = new QQIABot(dataService as any, dependencyEngine, excelSync, notificationService, paSyncService, comSync);

// Create webhook handler for Teams Outgoing Webhook (no AAD/Azure required)
const webhookHandler = new WebhookHandler(
  dataService, dependencyEngine, excelSync, notificationService,
  paSyncService, comSync, process.env.WEBHOOK_HMAC_SECRET, analyticsService
);

// ---- HTTP Server ----
const server = express();
server.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "'unsafe-inline'", "https://cdn.jsdelivr.net", "https://adaptivecards.io"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://cdn.jsdelivr.net"],
      imgSrc: ["'self'", "data:", "https:"],
      connectSrc: ["'self'"],
    },
  },
}));
server.use(express.json({ limit: '100kb' }));
server.use(express.urlencoded({ extended: true, limit: '100kb' }));

// ---- Access Control ----
const ACCESS_CODE = process.env.ACCESS_CODE || '';
if (ACCESS_CODE) {
  console.log('🔒 Access code protection enabled');
}

/** Check access code from header, query param, or cookie */
function isAuthorized(req: any): boolean {
  if (!ACCESS_CODE) return true; // No code configured = open access
  const code = req.get('x-access-code')
    || (req.query && req.query.code)
    || parseCookie(req, 'qqia_access');
  return code === ACCESS_CODE;
}

function parseCookie(req: any, name: string): string | null {
  const cookies = req.get('cookie') || '';
  const match = cookies.match(new RegExp('(?:^|;\\s*)' + name + '=([^;]*)'));
  return match ? decodeURIComponent(match[1]) : null;
}

/** Middleware: protect API endpoints */
function requireAccess(req: any, res: any, next: any) {
  // Allow health check and login endpoint without auth
  const openPaths = ['/api/health', '/api/auth/login', '/api/auth/check'];
  if (!ACCESS_CODE || openPaths.indexOf(req.path) >= 0) return next();
  if (isAuthorized(req)) return next();
  res.status(401).json({ error: 'Access code required' });
}

// Apply access control to all /api/ routes
server.use((req: any, res: any, next: any) => {
  if (req.path.startsWith('/api/')) {
    return requireAccess(req, res, next);
  }
  return next();
});

// Auth endpoints
server.post('/api/auth/login', (req: any, res: any) => {
  const { code } = req.body || {};
  if (code === ACCESS_CODE) {
    const securePart = process.env.NODE_ENV === 'production' ? ' Secure;' : '';
    res.setHeader('Set-Cookie', 'qqia_access=' + encodeURIComponent(code) + '; Path=/;' + securePart + ' HttpOnly; SameSite=Strict; Max-Age=604800');
    res.json({ success: true });
  } else {
    res.status(401).json({ error: 'Invalid access code' });
  }
});

server.get('/api/auth/check', (req: any, res: any) => {
  res.json({ authenticated: isAuthorized(req), required: !!ACCESS_CODE });
});

// CORS preflight for Office Script support
server.options('/api/steps/json', (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-access-code');
  res.sendStatus(204);
});

// Serve the web UI at root
const publicDir = path.join(__dirname, '..', 'public');
server.get('/', (req, res) => {
  const indexPath = path.join(publicDir, 'index.html');
  res.sendFile(indexPath);
});

// Serve static files from public/
server.use(express.static(publicDir));

// Serve the Office Script file
server.get('/api/sync-script', (req, res) => {
  const filePath = path.join(publicDir, 'SyncFromBot.osts');
  fs.readFile(filePath, 'utf8', (err, data) => {
    if (err) { res.status(404).send('Script not found'); return; }
    // Dynamically replace the bot URL with the actual host
    const host = req.get('host') || 'localhost:3978';
    const protocol = req.get('x-forwarded-proto') || 'https';
    const actualUrl = protocol + '://' + host + '/api/steps/json';
    let updated = data.replace(
      /var botUrl = ".*?";/,
      'var botUrl = "' + actualUrl + '";'
    );
    // Inject access code if configured
    if (ACCESS_CODE) {
      updated = updated.replace(
        /var accessCode = ".*?";/,
        'var accessCode = "' + ACCESS_CODE + '";'
      );
    }
    res.type('text/plain').send(updated);
  });
});

server.post('/api/messages', async (req, res) => {
  await adapter.process(req as any, res as any, async (context) => {
    // Save conversation reference for proactive messaging on every interaction
    if (context.activity.from) {
      proactiveMessenger.saveConversationReference(context.activity);
    }
    await bot.run(context);
  });
});

// Download the updated Excel file
server.get('/api/download/excel', (req, res) => {
  const excelPath = path.join(__dirname, '..', 'data', 'FY27_Mint_RolloverTimeline.xlsx');
  if (!fs.existsSync(excelPath)) {
    res.status(404).json({ error: 'Excel file not found' });
    return;
  }
  res.download(excelPath, `FY27_Mint_RolloverTimeline_${new Date().toISOString().split('T')[0]}.xlsx`);
});

// JSON API for Office Script sync - returns all steps with status data
server.get('/api/steps/json', (req, res) => {
  dataService.getAllSteps().then(steps => {
    // Add CORS headers so Office Scripts can call this endpoint
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.json({
      timestamp: new Date().toISOString(),
      stepCount: steps.length,
      steps: steps.map(s => ({
        stepId: s.id,
        status: s.corpStatus,
        fedStatus: s.fedStatus,
        completedDate: s.corpCompletedDate || null,
        lastModifiedBy: s.lastModifiedBy || null,
        lastModifiedDate: s.lastModified || null,
        lastModifiedSource: s.lastModifiedSource || null,
        description: s.description || null,
      })),
    });
  }).catch(() => {
    res.status(500).json({ error: 'Failed to retrieve steps' });
  });
});

// Health check endpoint with status details
server.get('/api/health', async (req, res) => {
  let stepCount = 0;
  try {
    const steps = await dataService.getAllSteps();
    stepCount = steps.length;
  } catch { /* Cosmos unavailable */ }

  res.json({
    status: 'healthy',
    timestamp: new Date().toISOString(),
    version: '1.0.0',
    registeredUsers: proactiveMessenger.getRegisteredUserCount(),
    trackedSteps: stepCount,
    excelUrl: config.graph.excelSharingUrl || '',
  });
});

// Audit log endpoint — admin view of all changes
server.get('/api/audit', async (req, res) => {
  try {
    const stepId = req.query.stepId as string | undefined;
    const source = req.query.source as string | undefined;
    const changedBy = req.query.changedBy as string | undefined;
    const limit = Math.min(parseInt(req.query.limit as string) || 200, 500);

    let entries = await (dataService as any).getAllAudit(limit);

    if (stepId) {
      entries = entries.filter((e: any) => e.stepId === stepId);
    }
    if (source) {
      entries = entries.filter((e: any) => e.source === source);
    }
    if (changedBy) {
      const q = changedBy.toLowerCase();
      entries = entries.filter((e: any) => e.changedBy.toLowerCase().includes(q));
    }

    res.json({
      count: entries.length,
      entries,
      timestamp: new Date().toISOString(),
    });
  } catch (err: any) {
    console.error('Audit endpoint error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ---- Analytics API Endpoints ----

// Burndown chart data
server.get('/api/analytics/burndown', async (req, res) => {
  try {
    const track = (req.query.track as string || 'Corp') as 'Corp' | 'Fed';
    const steps = await dataService.getAllSteps();
    // Ensure today's snapshot exists
    analyticsService.recordSnapshot(steps, track);
    const auditEntries = (dataService as any).getAllAudit ? await (dataService as any).getAllAudit(1000) : [];
    const burndown = analyticsService.getBurndown(track, auditEntries);
    res.json({ track, data: burndown });
  } catch (err: any) {
    console.error('Burndown error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Workstream health
server.get('/api/analytics/workstream-health', async (req, res) => {
  try {
    const track = (req.query.track as string || 'Corp') as 'Corp' | 'Fed';
    const steps = await dataService.getAllSteps();
    dependencyEngine.buildGraph(steps);
    const health = analyticsService.getWorkstreamHealth(steps, dependencyEngine, track);
    res.json({ track, workstreams: health });
  } catch (err: any) {
    console.error('Workstream health error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Timeline / Gantt data
server.get('/api/analytics/timeline', async (req, res) => {
  try {
    const track = (req.query.track as string || 'Corp') as 'Corp' | 'Fed';
    const steps = await dataService.getAllSteps();
    const timeline = analyticsService.getTimeline(steps, track);
    const milestones = await (dataService as any).getAllMilestones?.() || [];
    res.json({ track, steps: timeline, milestones });
  } catch (err: any) {
    console.error('Timeline error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Recent changes feed
server.get('/api/analytics/changes', async (req, res) => {
  try {
    const hours = Math.min(parseInt(req.query.hours as string) || 24, 168); // max 7 days
    const auditEntries = (dataService as any).getAllAudit ? await (dataService as any).getAllAudit(1000) : [];
    const changes = analyticsService.getRecentChanges(auditEntries, hours);
    res.json({ hours, count: changes.length, changes });
  } catch (err: any) {
    console.error('Changes error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Manual sync trigger endpoint (webhook for Power Automate / Logic Apps)
server.post('/api/sync', async (req, res) => {
  try {
    const result = await syncEngine.runSync();
    // Auto-export staging file for Excel Online sync
    try { await syncEngine.exportStagingFile(); } catch (e: any) {
      console.warn('[Sync] Staging file export failed:', e.message);
    }
    res.json(result);
  } catch (err: any) {
    console.error('Sync endpoint error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Export staging file for Excel Online Office Script sync
server.post('/api/export-staging', async (req, res) => {
  try {
    const filePath = await syncEngine.exportStagingFile();
    res.json({ success: true, filePath, timestamp: new Date().toISOString() });
  } catch (err: any) {
    console.error('Export staging error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Teams Outgoing Webhook endpoint (no Azure/AAD required)
server.post('/api/webhook', async (req, res) => {
  try {
    // Validate HMAC signature if configured
    const authHeader = req.headers['authorization'] as string || '';
    const rawBody = JSON.stringify(req.body);
    if (!webhookHandler.validateSignature(rawBody, authHeader)) {
      res.status(401).json({ error: 'Invalid HMAC signature' });
      return;
    }

    const response = await webhookHandler.processRequest(req.body);
    res.json(response);
  } catch (err: any) {
    console.error('Webhook error:', err);
    res.json({ type: 'message', text: '❌ Something went wrong. Please try again.' });
  }
});

// Webhook for external automation status updates (ADO, pipelines, etc.)
server.post('/api/automation/status', async (req, res) => {
  try {
    // Require API key for automation writes
    const AUTOMATION_KEY = process.env.AUTOMATION_API_KEY || '';
    if (AUTOMATION_KEY) {
      const provided = req.headers['x-api-key'] as string || '';
      if (provided !== AUTOMATION_KEY) {
        res.status(401).json({ error: 'Invalid or missing API key' });
        return;
      }
    }

    const { stepId, status, track, source, notes } = req.body || {};
    if (!stepId || !status) {
      res.status(400).json({ error: 'stepId and status are required' });
      return;
    }

    const field = track === 'Fed' ? 'fedStatus' : 'corpStatus';
    const updated = await dataService.updateStepStatus(
      stepId, field as any, status, source || 'automation', 'automation'
    );

    if (!updated) {
      res.status(404).json({ error: `Step ${stepId} not found` });
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

    // Sync to Excel immediately (COM → Power Automate → Graph API → local file fallback)
    try {
      if (comSync.isAvailable) {
        const syncCount = await comSync.syncToExcel();
        console.log(`Automation update: synced ${syncCount} step(s) via Excel COM`);
      } else if (paSyncService.isConfigured) {
        const syncCount = await paSyncService.syncToExcel();
        console.log(`Automation update: synced ${syncCount} step(s) via Power Automate`);
      } else {
        const syncCount = await excelSync.syncToExcel();
        console.log(`Automation update: synced ${syncCount} step(s) to Excel`);
      }
    } catch (syncErr: any) {
      console.warn(`Automation update: Excel sync failed: ${syncErr.message}`);
    }

    // Export staging file for Excel Online sync
    try { await syncEngine.exportStagingFile(); } catch (e: any) {
      console.warn('Staging file export failed:', e.message);
    }

    res.json({ success: true, step: updated });
  } catch (err: any) {
    console.error('Automation endpoint error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ---- Scheduled Tasks ----

/** Excel sync: Power Automate → Graph API → local file (every 15 minutes) */
function startSyncScheduler() {
  setInterval(async () => {
    try {
      console.log(`[Sync] Starting bi-directional Excel sync...`);
      if (paSyncService.isConfigured) {
        // Power Automate sync (no app registration needed)
        const result = await paSyncService.fullSync();
        console.log(`[Sync] PA sync: ${result.fromExcel} from Excel, ${result.toExcel} to Excel`);
      } else {
        // Graph API or local file sync
        const result = await syncEngine.runSync();
        console.log(`[Sync] Complete: ${result.fromExcel} from Excel, ${result.toExcel} to Excel, ${result.conflicts.length} conflicts`);
      }
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

  // Initialize Graph API (app-only — client credentials)
  try {
    if (config.graph.clientId) {
      await graphService.initialize();
      console.log('✅ Graph API initialized');
      excelSync.enableGraphApi(graphService);
    } else {
      console.warn('⚠️ Graph API credentials not configured. Using local Excel file only.');
    }
  } catch (err: any) {
    console.warn(`⚠️ Graph API init failed (${err.message}). Using local Excel fallback.`);
  }

  // Delegated Graph API disabled — Conditional Access blocks device code flow from Codespace.
  // Excel sync is done via Office Script instead: the Excel file pulls from /api/steps/json.
  console.log('📡 Excel sync mode: Office Script (Excel pulls from /api/steps/json)');

  // Seed data from Excel into the data store on startup
  try {
    const importResult = await excelSync.importFromExcel();
    console.log(`📊 Loaded ${importResult.steps} rollover steps from Excel`);
    console.log(`🏁 Loaded ${importResult.milestones} milestones from Excel`);

    // Build dependency graph
    const allSteps = await dataService.getAllSteps();
    dependencyEngine.buildGraph(allSteps);
    const validation = dependencyEngine.validateDAG();
    console.log(`🔗 Dependency graph: ${allSteps.length} nodes, DAG valid: ${validation.valid}`);
    if (!validation.valid) {
      console.warn(`⚠️ Cycles detected: ${validation.cycles.map(c => c.join('→')).join('; ')}`);
    }

    // Seed today's analytics snapshot
    analyticsService.recordSnapshot(allSteps, 'Corp');
    analyticsService.recordSnapshot(allSteps, 'Fed');
    console.log('📈 Analytics snapshot recorded');
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
    console.log(`   Webhook:        POST /api/webhook  (Teams Outgoing Webhook)`);
    console.log(`   Health check:   GET  /api/health`);
    console.log(`   Manual sync:    POST /api/sync`);
    console.log(`   Auto-update:    POST /api/automation/status`);
    console.log(`   Analytics:      GET  /api/analytics/*\n`);
  });
}

main().catch((err) => {
  console.error('Fatal error:', err);
  process.exit(1);
});
