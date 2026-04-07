/**
 * Helper script to find the Drive ID and Item ID for your Excel file.
 * 
 * Usage:
 *   1. Set GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_TENANT_ID in .env
 *   2. Run: node scripts/find-excel-ids.js
 * 
 * This uses app-only auth (client credentials flow) to look up the file.
 * Requires Files.ReadWrite.All or Sites.ReadWrite.All permission.
 */

require('dotenv').config();
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

async function main() {
  const clientId = process.env.GRAPH_CLIENT_ID;
  const clientSecret = process.env.GRAPH_CLIENT_SECRET;
  const tenantId = process.env.GRAPH_TENANT_ID;

  if (!clientId || !clientSecret || !tenantId) {
    console.error('❌ Missing Graph API credentials in .env file.');
    console.error('   Set GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, and GRAPH_TENANT_ID');
    process.exit(1);
  }

  console.log('🔑 Authenticating with Graph API...');

  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });
  const client = Client.initWithMiddleware({ authProvider });

  // Search for the file across SharePoint sites
  const filePath = process.env.EXCEL_FILE_PATH || '/Seller Incentives/QQIA/FY27_Mint_RolloverTimeline.xlsx';
  const fileName = filePath.split('/').pop();

  console.log(`\n🔍 Searching for "${fileName}" across SharePoint/OneDrive...\n`);

  try {
    // Method 1: Search across all drives
    const searchResults = await client
      .api('/search/query')
      .post({
        requests: [{
          entityTypes: ['driveItem'],
          query: { queryString: `filename:${fileName}` },
        }],
      });

    const hits = searchResults.value?.[0]?.hitsContainers?.[0]?.hits || [];

    if (hits.length === 0) {
      console.log('No files found. Try Method 2 below.\n');
    } else {
      console.log(`Found ${hits.length} result(s):\n`);
      for (const hit of hits) {
        const resource = hit.resource;
        console.log(`📄 ${resource.name}`);
        console.log(`   Path: ${resource.webUrl}`);
        console.log(`   Drive ID:  ${resource.parentReference?.driveId || 'N/A'}`);
        console.log(`   Item ID:   ${resource.id}`);
        console.log(`   Site ID:   ${resource.parentReference?.siteId || 'N/A'}`);
        console.log('');
      }

      if (hits.length > 0) {
        const first = hits[0].resource;
        console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
        console.log('Add these to your .env file:');
        console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
        console.log(`EXCEL_DRIVE_ID=${first.parentReference?.driveId || ''}`);
        console.log(`EXCEL_ITEM_ID=${first.id}`);
        console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      }
    }
  } catch (err) {
    console.error(`Search failed: ${err.message}`);
    console.log('\nTrying direct path lookup...\n');
  }

  // Method 2: Try specific user's OneDrive path
  try {
    console.log('\n📂 Trying OneDrive path lookup...');
    // Note: App-only auth can't use /me/drive — need a specific user ID or site
    // Try listing sites to help the user find the right one
    const sites = await client.api('/sites?search=*').top(10).get();
    
    if (sites.value?.length > 0) {
      console.log('\nAvailable SharePoint sites:');
      for (const site of sites.value) {
        console.log(`  • ${site.displayName} — ${site.webUrl} (ID: ${site.id})`);
      }
      console.log('\nTo find your file on a specific site, use Graph Explorer:');
      console.log(`  GET https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:${filePath}`);
    }
  } catch (err) {
    console.log(`Site listing: ${err.message}`);
  }
}

main().catch(err => {
  console.error('Fatal error:', err.message);
  process.exit(1);
});
