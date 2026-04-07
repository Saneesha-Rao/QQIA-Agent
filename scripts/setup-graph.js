/**
 * QQIA Agent — SharePoint Graph API Setup Script
 * 
 * This script:
 *   1. Opens your browser for Microsoft login
 *   2. Creates an App Registration (QQIA-Agent-Graph)
 *   3. Adds Files.ReadWrite.All permission
 *   4. Creates a client secret
 *   5. Finds your Excel file's Drive ID & Item ID
 *   6. Updates your .env file
 * 
 * Usage: node scripts/setup-graph.js
 */

const { PublicClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
const fs = require('fs');
const path = require('path');
const open = (() => {
  // Use start command on Windows to open browser
  return (url) => require('child_process').exec(`start "${url}"`);
})();

// Microsoft tenant for interactive login
// We use a well-known first-party client ID for device code / interactive login
// that has permission to manage app registrations
const WELL_KNOWN_CLIENT_ID = '04b07795-a71b-4346-935f-02f9a1efa439'; // Azure CLI client ID

const GRAPH_SCOPES = [
  'Application.ReadWrite.All',   // Create app registrations
  'Files.Read.All',              // Find Excel file
  'User.Read',                   // Get tenant info
];

const envPath = path.join(__dirname, '..', '.env');

async function main() {
  console.log('╔══════════════════════════════════════════════════╗');
  console.log('║   QQIA Agent — SharePoint Graph API Setup       ║');
  console.log('╚══════════════════════════════════════════════════╝\n');

  // Step 1: Interactive login via device code flow
  console.log('Step 1: Signing in to Microsoft Graph...\n');
  
  const msalConfig = {
    auth: {
      clientId: WELL_KNOWN_CLIENT_ID,
      authority: 'https://login.microsoftonline.com/organizations',
    },
  };

  const pca = new PublicClientApplication(msalConfig);

  let tokenResponse;
  try {
    // Use device code flow (works with corp Conditional Access policies)
    tokenResponse = await pca.acquireTokenByDeviceCode({
      scopes: GRAPH_SCOPES.map(s => `https://graph.microsoft.com/${s}`),
      deviceCodeCallback: (response) => {
        console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
        console.log(response.message);
        console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
        // Try to open browser automatically
        try {
          require('child_process').exec(`start ${response.verificationUri}`);
        } catch { /* user can open manually */ }
      },
    });
    console.log(`✅ Signed in as: ${tokenResponse.account.username}\n`);
  } catch (err) {
    console.error(`❌ Login failed: ${err.message}`);
    process.exit(1);
  }

  // Create Graph client with the user's token
  const client = Client.init({
    authProvider: (done) => done(null, tokenResponse.accessToken),
  });

  // Get tenant ID
  const tenantId = tokenResponse.tenantId;
  console.log(`📋 Tenant ID: ${tenantId}\n`);

  // Step 2: Create App Registration
  console.log('Step 2: Creating App Registration...\n');
  
  let appRegistration;
  try {
    // Check if it already exists
    const existing = await client
      .api('/applications')
      .filter("displayName eq 'QQIA-Agent-Graph'")
      .get();

    if (existing.value?.length > 0) {
      appRegistration = existing.value[0];
      console.log(`📌 App Registration already exists: ${appRegistration.appId}\n`);
    } else {
      // Create new app registration
      appRegistration = await client.api('/applications').post({
        displayName: 'QQIA-Agent-Graph',
        signInAudience: 'AzureADMyOrg',
        requiredResourceAccess: [{
          resourceAppId: '00000003-0000-0000-c000-000000000000', // Microsoft Graph
          resourceAccess: [
            { id: '01d4f6ba-44d6-490e-ab99-304e510cfc68', type: 'Role' }, // Files.ReadWrite.All
            { id: '332a536c-c7ef-4017-ab91-336970924f0d', type: 'Role' }, // Sites.ReadWrite.All
          ],
        }],
      });
      console.log(`✅ Created App Registration: ${appRegistration.appId}\n`);
    }
  } catch (err) {
    console.error(`❌ Failed to create App Registration: ${err.message}`);
    console.error('   You may need Application.ReadWrite.All permission.');
    console.error('   Try creating the app registration manually in the Azure Portal.');
    process.exit(1);
  }

  const clientId = appRegistration.appId;
  const objectId = appRegistration.id;

  // Step 3: Create client secret
  console.log('Step 3: Creating client secret...\n');

  let clientSecret;
  try {
    const secretResult = await client
      .api(`/applications/${objectId}/addPassword`)
      .post({
        passwordCredential: {
          displayName: 'QQIA Agent Excel Access',
          endDateTime: new Date(Date.now() + 365 * 24 * 60 * 60 * 1000).toISOString(),
        },
      });
    clientSecret = secretResult.secretText;
    console.log(`✅ Client secret created (expires in 12 months)\n`);
  } catch (err) {
    console.error(`❌ Failed to create secret: ${err.message}`);
    process.exit(1);
  }

  // Step 4: Request admin consent
  console.log('Step 4: Admin consent...\n');

  try {
    // Try to grant admin consent via service principal
    // First, create the service principal if it doesn't exist
    let servicePrincipal;
    const spSearch = await client
      .api('/servicePrincipals')
      .filter(`appId eq '${clientId}'`)
      .get();

    if (spSearch.value?.length > 0) {
      servicePrincipal = spSearch.value[0];
    } else {
      servicePrincipal = await client.api('/servicePrincipals').post({
        appId: clientId,
      });
    }

    // Get Microsoft Graph service principal
    const graphSP = await client
      .api('/servicePrincipals')
      .filter("appId eq '00000003-0000-0000-c000-000000000000'")
      .get();

    if (graphSP.value?.length > 0) {
      const graphSpId = graphSP.value[0].id;
      const appRoles = [
        '01d4f6ba-44d6-490e-ab99-304e510cfc68', // Files.ReadWrite.All
        '332a536c-c7ef-4017-ab91-336970924f0d', // Sites.ReadWrite.All
      ];

      for (const roleId of appRoles) {
        try {
          await client.api('/servicePrincipals/' + servicePrincipal.id + '/appRoleAssignments').post({
            principalId: servicePrincipal.id,
            resourceId: graphSpId,
            appRoleId: roleId,
          });
        } catch (e) {
          // May fail if already granted or insufficient permissions
          if (!e.message?.includes('already exists')) {
            console.warn(`   ⚠️ Could not auto-grant role ${roleId}: ${e.message?.substring(0, 80)}`);
          }
        }
      }
      console.log(`✅ Admin consent granted (or already existed)\n`);
    }
  } catch (err) {
    console.warn(`⚠️ Auto-consent failed: ${err.message?.substring(0, 100)}`);
    console.log(`   Please grant admin consent manually in Azure Portal:`);
    console.log(`   https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/${clientId}\n`);
  }

  // Step 5: Find Excel file
  console.log('Step 5: Finding your Excel file on OneDrive/SharePoint...\n');

  let driveId = '';
  let itemId = '';

  try {
    // Try user's OneDrive first
    const fileInfo = await client
      .api('/me/drive/root:/Seller Incentives/QQIA/FY27_Mint_RolloverTimeline.xlsx')
      .select('id,name,parentReference,webUrl')
      .get();

    driveId = fileInfo.parentReference?.driveId || '';
    itemId = fileInfo.id;
    console.log(`✅ Found Excel file!`);
    console.log(`   Name: ${fileInfo.name}`);
    console.log(`   URL:  ${fileInfo.webUrl}`);
    console.log(`   Drive ID: ${driveId}`);
    console.log(`   Item ID:  ${itemId}\n`);
  } catch (err) {
    console.warn(`⚠️ Could not find file on OneDrive: ${err.message?.substring(0, 80)}`);
    console.log('   Searching SharePoint sites...\n');

    try {
      // Search SharePoint
      const searchResults = await client
        .api('/search/query')
        .post({
          requests: [{
            entityTypes: ['driveItem'],
            query: { queryString: 'filename:FY27_Mint_RolloverTimeline.xlsx' },
          }],
        });

      const hits = searchResults.value?.[0]?.hitsContainers?.[0]?.hits || [];
      if (hits.length > 0) {
        const resource = hits[0].resource;
        driveId = resource.parentReference?.driveId || '';
        itemId = resource.id;
        console.log(`✅ Found via search!`);
        console.log(`   Drive ID: ${driveId}`);
        console.log(`   Item ID:  ${itemId}\n`);
      } else {
        console.warn('⚠️ File not found. You can set EXCEL_DRIVE_ID and EXCEL_ITEM_ID manually later.');
        console.log('   Use Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer\n');
      }
    } catch (searchErr) {
      console.warn(`⚠️ Search failed: ${searchErr.message?.substring(0, 80)}`);
    }
  }

  // Step 6: Update .env file
  console.log('Step 6: Updating .env file...\n');

  try {
    let envContent = fs.readFileSync(envPath, 'utf-8');

    const updates = {
      'GRAPH_CLIENT_ID': clientId,
      'GRAPH_CLIENT_SECRET': clientSecret,
      'GRAPH_TENANT_ID': tenantId,
      'EXCEL_DRIVE_ID': driveId,
      'EXCEL_ITEM_ID': itemId,
    };

    for (const [key, value] of Object.entries(updates)) {
      if (value) {
        const regex = new RegExp(`^${key}=.*$`, 'm');
        if (regex.test(envContent)) {
          envContent = envContent.replace(regex, `${key}=${value}`);
        } else {
          envContent += `\n${key}=${value}`;
        }
      }
    }

    fs.writeFileSync(envPath, envContent);
    console.log(`✅ Updated .env with Graph API configuration\n`);
  } catch (err) {
    console.error(`❌ Failed to update .env: ${err.message}`);
    console.log('\nManually add these to your .env file:');
    console.log(`GRAPH_CLIENT_ID=${clientId}`);
    console.log(`GRAPH_CLIENT_SECRET=${clientSecret}`);
    console.log(`GRAPH_TENANT_ID=${tenantId}`);
    console.log(`EXCEL_DRIVE_ID=${driveId}`);
    console.log(`EXCEL_ITEM_ID=${itemId}`);
  }

  // Summary
  console.log('╔══════════════════════════════════════════════════╗');
  console.log('║   ✅ Setup Complete!                             ║');
  console.log('╚══════════════════════════════════════════════════╝\n');
  console.log('App Registration:');
  console.log(`  Client ID:     ${clientId}`);
  console.log(`  Tenant ID:     ${tenantId}`);
  console.log(`  Secret:        ${clientSecret.substring(0, 4)}...${clientSecret.substring(clientSecret.length - 4)}`);
  console.log(`\nExcel File:`);
  console.log(`  Drive ID:      ${driveId || '(not found — set manually)'}`);
  console.log(`  Item ID:       ${itemId || '(not found — set manually)'}`);
  console.log('\nNext steps:');
  console.log('  1. Restart the bot:  npm run build && npm start');
  console.log('  2. You should see: "📡 Excel sync: Graph API mode enabled"');
  console.log('  3. Try: "update 1.C completed" — SharePoint Excel updates instantly!\n');
  
  if (!driveId || !itemId) {
    console.log('⚠️ Excel file IDs not found. Set them manually:');
    console.log('   1. Open Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer');
    console.log('   2. GET /me/drive/root:/Seller Incentives/QQIA/FY27_Mint_RolloverTimeline.xlsx');
    console.log('   3. Copy "id" → EXCEL_ITEM_ID, "parentReference.driveId" → EXCEL_DRIVE_ID\n');
  }
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
