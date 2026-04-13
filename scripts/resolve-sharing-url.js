/**
 * Resolve a SharePoint/OneDrive sharing URL to get Drive ID and Item ID.
 * 
 * Usage:
 *   1. Set GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_TENANT_ID in .env
 *   2. Run: node scripts/resolve-sharing-url.js "<sharing-url>"
 */

require('dotenv').config();
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

function encodeSharingUrl(url) {
  const base64 = Buffer.from(url).toString('base64');
  // Convert to base64url: replace / with _, + with -, remove trailing =
  const encoded = base64.replace(/\//g, '_').replace(/\+/g, '-').replace(/=+$/, '');
  return 'u!' + encoded;
}

async function main() {
  const sharingUrl = process.argv[2] || 'https://microsoftapc-my.sharepoint.com/:x:/g/personal/salingal_microsoft_com/IQC6ExDXsSeQSJFRRucpbQ3UAc6CvGveaYkO5eEt_uJ1sxc?e=QcBZAg';

  const clientId = process.env.GRAPH_CLIENT_ID;
  const clientSecret = process.env.GRAPH_CLIENT_SECRET;
  const tenantId = process.env.GRAPH_TENANT_ID;

  if (!clientId || !clientSecret || !tenantId) {
    console.error('Missing Graph API credentials. Set GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_TENANT_ID');
    process.exit(1);
  }

  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });
  const client = Client.initWithMiddleware({ authProvider });

  const shareToken = encodeSharingUrl(sharingUrl);
  console.log(`Resolving sharing URL...`);
  console.log(`Share token: ${shareToken}\n`);

  try {
    const driveItem = await client.api(`/shares/${shareToken}/driveItem`).get();
    
    console.log('File found:');
    console.log(`  Name:     ${driveItem.name}`);
    console.log(`  Web URL:  ${driveItem.webUrl}`);
    console.log(`  Drive ID: ${driveItem.parentReference?.driveId}`);
    console.log(`  Item ID:  ${driveItem.id}`);
    console.log(`  Path:     ${driveItem.parentReference?.path}\n`);

    console.log('Add these to your .env file:');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log(`EXCEL_DRIVE_ID=${driveItem.parentReference?.driveId}`);
    console.log(`EXCEL_ITEM_ID=${driveItem.id}`);
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  } catch (err) {
    console.error(`Failed to resolve sharing URL: ${err.message}`);
    if (err.statusCode === 403) {
      console.error('The app registration needs Sites.Read.All or Files.Read.All permission.');
    }
  }
}

main().catch(console.error);
