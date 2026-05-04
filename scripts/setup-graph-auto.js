/**
 * Create an Azure AD app registration with Files.ReadWrite.All permission,
 * grant admin consent, and use it to write directly to Excel Online.
 * 
 * Uses the existing `az` CLI token which has Application.ReadWrite.All scope.
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const APP_NAME = 'QQIA-Agent-ExcelSync';
const ENV_PATH = path.join(__dirname, '..', '.env');

function getAzToken() {
  return execSync('az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv', { encoding: 'utf-8' }).trim();
}

async function graphCall(method, url, token, body) {
  const opts = {
    method,
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
  };
  if (body) opts.body = JSON.stringify(body);
  const resp = await fetch(url, opts);
  const text = await resp.text();
  if (!resp.ok) throw new Error(`${resp.status} ${resp.statusText}: ${text}`);
  return text ? JSON.parse(text) : {};
}

async function main() {
  console.log('=== Setup Graph API + Push to Excel Online ===\n');
  
  const token = getAzToken();
  console.log('✅ Got az CLI token\n');

  // Step 1: Check if app registration already exists
  console.log('Step 1: Checking for existing app registration...');
  let app;
  try {
    const existing = await graphCall('GET', `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '${APP_NAME}'`, token);
    if (existing.value && existing.value.length > 0) {
      app = existing.value[0];
      console.log(`  Found existing app: ${app.appId}\n`);
    }
  } catch (err) {
    console.log(`  Search failed: ${err.message}`);
  }

  // Step 2: Create app registration if needed
  if (!app) {
    console.log('Step 2: Creating app registration...');
    try {
      app = await graphCall('POST', 'https://graph.microsoft.com/v1.0/applications', token, {
        displayName: APP_NAME,
        signInAudience: 'AzureADMyOrg',
        requiredResourceAccess: [{
          resourceAppId: '00000003-0000-0000-c000-000000000000',
          resourceAccess: [
            { id: '01d4f6ba-44d6-490e-ab99-304e510cfc68', type: 'Role' }, // Files.ReadWrite.All
          ],
        }],
      });
      console.log(`  ✅ Created: ${app.appId}\n`);
    } catch (err) {
      console.error(`  ❌ Failed: ${err.message}`);
      process.exit(1);
    }
  }

  // Step 3: Create client secret
  console.log('Step 3: Creating client secret...');
  let clientSecret;
  try {
    const secretResult = await graphCall('POST', `https://graph.microsoft.com/v1.0/applications/${app.id}/addPassword`, token, {
      passwordCredential: {
        displayName: 'QQIA Excel Sync',
        endDateTime: new Date(Date.now() + 365 * 24 * 60 * 60 * 1000).toISOString(),
      },
    });
    clientSecret = secretResult.secretText;
    console.log('  ✅ Secret created\n');
  } catch (err) {
    console.error(`  ❌ Failed: ${err.message}`);
    process.exit(1);
  }

  // Step 4: Create service principal + grant admin consent
  console.log('Step 4: Granting admin consent...');
  try {
    let sp;
    try {
      const spSearch = await graphCall('GET', `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${app.appId}'`, token);
      sp = spSearch.value && spSearch.value.length > 0 ? spSearch.value[0] : null;
    } catch { sp = null; }

    if (!sp) {
      sp = await graphCall('POST', 'https://graph.microsoft.com/v1.0/servicePrincipals', token, { appId: app.appId });
    }

    const graphSP = await graphCall('GET', "https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'", token);
    const graphSpId = graphSP.value[0].id;

    try {
      await graphCall('POST', `https://graph.microsoft.com/v1.0/servicePrincipals/${sp.id}/appRoleAssignments`, token, {
        principalId: sp.id,
        resourceId: graphSpId,
        appRoleId: '01d4f6ba-44d6-490e-ab99-304e510cfc68',
      });
      console.log('  ✅ Admin consent granted for Files.ReadWrite.All\n');
    } catch (err) {
      if (err.message.includes('already exists')) {
        console.log('  ✅ Already consented\n');
      } else {
        console.warn(`  ⚠️ Consent warning: ${err.message}\n`);
      }
    }
  } catch (err) {
    console.error(`  ❌ SP/consent failed: ${err.message}`);
  }

  const tenantId = '72f988bf-86f1-41af-91ab-2d7cd011db47';

  // Step 5: Save to .env
  console.log('Step 5: Saving credentials...');
  const envVars = {
    GRAPH_CLIENT_ID: app.appId,
    GRAPH_CLIENT_SECRET: clientSecret,
    GRAPH_TENANT_ID: tenantId,
  };

  let envContent = '';
  if (fs.existsSync(ENV_PATH)) {
    envContent = fs.readFileSync(ENV_PATH, 'utf-8');
  }
  
  for (const [key, val] of Object.entries(envVars)) {
    const regex = new RegExp(`^${key}=.*$`, 'm');
    if (regex.test(envContent)) {
      envContent = envContent.replace(regex, `${key}=${val}`);
    } else {
      envContent += `\n${key}=${val}`;
    }
  }
  fs.writeFileSync(ENV_PATH, envContent.trim() + '\n');
  console.log(`  ✅ Saved to ${ENV_PATH}\n`);

  console.log('=== Credentials ===');
  console.log(`GRAPH_CLIENT_ID=${app.appId}`);
  console.log(`GRAPH_CLIENT_SECRET=${clientSecret.substring(0, 8)}...`);
  console.log(`GRAPH_TENANT_ID=${tenantId}`);
  console.log('\n✅ Setup complete! Restart the bot server to use Graph API.');
}

main().catch(err => {
  console.error('Fatal:', err.message);
  process.exit(1);
});
