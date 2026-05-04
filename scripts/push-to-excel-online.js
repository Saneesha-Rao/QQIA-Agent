/**
 * Push bot updates directly to Excel Online via Graph API.
 * 
 * Uses MSAL interactive browser login — no app registration needed (uses Azure CLI client ID).
 * Writes cells directly to the server-side Excel file, bypassing OneDrive sync.
 * 
 * Usage: node scripts/push-to-excel-online.js
 */

const { PublicClientApplication, InteractionRequiredAuthError } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
const http = require('http');
const { URL } = require('url');
const { execSync } = require('child_process');

// Microsoft Office first-party client ID (always registered in MS corporate tenants)
const MS_OFFICE_CLIENT_ID = 'd3590ed6-52b3-4102-aeff-aad2292ab01c';
const REDIRECT_URI = 'http://localhost:8400';
const SCOPES = ['https://graph.microsoft.com/Files.ReadWrite'];

// Sharing URL from config
const SHARING_URL = 'https://microsoftapc-my.sharepoint.com/:x:/g/personal/salingal_microsoft_com/IQC6ExDXsSeQSJFRRucpbQ3UAc6CvGveaYkO5eEt_uJ1sxc?e=nPH6Yx';

function encodeSharingUrl(url) {
  const base64 = Buffer.from(url).toString('base64');
  return 'u!' + base64.replace(/\//g, '_').replace(/\+/g, '-').replace(/=+$/, '');
}

/** Get auth code via local HTTP server + browser redirect */
function getAuthCodeFromBrowser(authUrl) {
  return new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      const url = new URL(req.url, REDIRECT_URI);
      const code = url.searchParams.get('code');
      const error = url.searchParams.get('error');
      
      if (code) {
        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.end('<h2>Authentication successful!</h2><p>You can close this tab.</p><script>window.close()</script>');
        server.close();
        resolve(code);
      } else if (error) {
        res.writeHead(400, { 'Content-Type': 'text/html' });
        res.end(`<h2>Error: ${error}</h2><p>${url.searchParams.get('error_description')}</p>`);
        server.close();
        reject(new Error(`Auth error: ${error} - ${url.searchParams.get('error_description')}`));
      }
    });
    
    server.listen(8400, () => {
      // Open browser
      try { execSync(`start "" "${authUrl}"`); } catch { /* user opens manually */ }
    });
    
    // Timeout after 2 minutes
    setTimeout(() => { server.close(); reject(new Error('Auth timeout')); }, 120000);
  });
}

async function getToken() {
  const pca = new PublicClientApplication({
    auth: {
      clientId: MS_OFFICE_CLIENT_ID,
      authority: 'https://login.microsoftonline.com/organizations',
    },
  });

  // Try cached token first
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await pca.acquireTokenSilent({ account: accounts[0], scopes: SCOPES });
      console.log(`Using cached token for: ${accounts[0].username}`);
      return result.accessToken;
    } catch { /* need interactive */ }
  }

  // Interactive: auth code flow with local redirect
  const authUrl = await pca.getAuthCodeUrl({
    scopes: SCOPES,
    redirectUri: REDIRECT_URI,
    prompt: 'select_account',
  });

  console.log('Opening browser for sign-in...');
  const code = await getAuthCodeFromBrowser(authUrl);

  const result = await pca.acquireTokenByCode({
    code,
    scopes: SCOPES,
    redirectUri: REDIRECT_URI,
  });

  console.log(`✅ Signed in as: ${result.account.username}`);
  return result.accessToken;
}

async function main() {
  console.log('=== Push Updates to Excel Online ===\n');

  // Step 1: Get token
  const token = await getToken();
  console.log('');

  // Step 2: Create Graph client
  const client = Client.init({
    authProvider: (done) => done(null, token),
  });

  // Step 3: Resolve sharing URL to get drive/item IDs
  const shareToken = encodeSharingUrl(SHARING_URL);
  console.log('Resolving file location...');
  
  let driveId, itemId;
  try {
    const driveItem = await client.api(`/shares/${shareToken}/driveItem`).get();
    driveId = driveItem.parentReference?.driveId;
    itemId = driveItem.id;
    console.log(`  File: ${driveItem.name}`);
    console.log(`  Drive: ${driveId}`);
    console.log(`  Item:  ${itemId}\n`);
  } catch (err) {
    console.error(`Failed to resolve file: ${err.message}`);
    process.exit(1);
  }

  // Step 4: Get current steps from local bot API
  let steps;
  try {
    const resp = await fetch('http://localhost:3978/api/steps/json');
    const data = await resp.json();
    steps = data.steps;
    console.log(`Got ${steps.length} steps from bot API\n`);
  } catch (err) {
    console.error(`Bot API not running: ${err.message}`);
    process.exit(1);
  }

  // Step 5: Read current Excel Online data to find row mappings
  console.log('Reading Excel Online data...');
  const rangeUrl = `/drives/${driveId}/items/${itemId}/workbook/worksheets/FY27_Rollover/usedRange`;
  let excelRows;
  try {
    const range = await client.api(rangeUrl).get();
    excelRows = range.values;
    console.log(`  ${excelRows.length} rows in Excel Online\n`);
  } catch (err) {
    console.error(`Failed to read Excel: ${err.message}`);
    process.exit(1);
  }

  // Build row map (step ID → row number, 1-based)
  const idToRow = new Map();
  for (let i = 3; i < excelRows.length; i++) {
    const id = excelRows[i][0]?.toString().trim();
    if (id) idToRow.set(id, i + 1); // +1 because Excel rows are 1-based
  }

  // Step 6: Find steps that differ from Excel Online and update them
  const updates = [];
  for (const step of steps) {
    const row = idToRow.get(step.stepId);
    if (!row) continue;

    const excelIdx = row - 1; // back to 0-based for array access
    const excelStatus = (excelRows[excelIdx][5] || '').toString().trim();
    
    if (step.status !== excelStatus) {
      updates.push({ stepId: step.stepId, row, field: 'F', value: step.status, oldValue: excelStatus });
    }
  }

  if (updates.length === 0) {
    console.log('✅ Excel Online is already up to date!');
    return;
  }

  console.log(`Found ${updates.length} step(s) to update:\n`);
  for (const u of updates) {
    console.log(`  ${u.stepId}: "${u.oldValue}" → "${u.value}" (row ${u.row})`);
  }
  console.log('');

  // Step 7: Write updates to Excel Online
  let successCount = 0;
  for (const u of updates) {
    const cellUrl = `/drives/${driveId}/items/${itemId}/workbook/worksheets/FY27_Rollover/range(address='${u.field}${u.row}')`;
    try {
      await client.api(cellUrl).patch({ values: [[u.value]] });
      successCount++;
      console.log(`  ✅ ${u.stepId} → ${u.value}`);
    } catch (err) {
      console.error(`  ❌ ${u.stepId}: ${err.message}`);
    }
  }

  console.log(`\n${successCount}/${updates.length} cells updated in Excel Online.`);

  // Also update completed dates for completed steps
  for (const u of updates) {
    if (u.value === 'Completed') {
      const step = steps.find(s => s.stepId === u.stepId);
      if (step?.completedDate) {
        const cellUrl = `/drives/${driveId}/items/${itemId}/workbook/worksheets/FY27_Rollover/range(address='G${u.row}')`;
        try {
          await client.api(cellUrl).patch({ values: [[step.completedDate]] });
          console.log(`  ✅ ${u.stepId} completed date → ${step.completedDate}`);
        } catch { /* non-critical */ }
      }
    }
  }

  console.log('\nDone! Refresh Excel Online to see changes.');
}

main().catch(err => {
  console.error('Fatal error:', err.message);
  process.exit(1);
});
