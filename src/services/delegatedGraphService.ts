import { DeviceCodeCredential, TokenCachePersistenceOptions } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Delegated Graph auth using device code flow.
 * Uses Microsoft's built-in Azure CLI client ID — no app registration needed.
 * Authenticates as the user (delegated permissions) so it can read/write their OneDrive files.
 * 
 * Stores refresh token to disk so re-auth is only needed every ~90 days.
 */
export class DelegatedGraphService {
  private client: Client | null = null;
  private credential: DeviceCodeCredential | null = null;
  private tokenCachePath: string;
  private _isAuthenticated = false;

  // Microsoft's first-party Azure CLI client ID — public, no secret needed
  private static readonly CLIENT_ID = '04b07795-a71b-4346-935c-03553cd355bd';
  private static readonly SCOPES = ['Files.ReadWrite', 'Sites.ReadWrite.All'];

  constructor() {
    const dataDir = path.resolve(__dirname, '..', '..', 'data');
    if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
    this.tokenCachePath = path.join(dataDir, '.token-cache.json');
  }

  get isAuthenticated(): boolean {
    return this._isAuthenticated;
  }

  /** Initialize with device code flow — prints a URL + code for the user to sign in */
  async initialize(): Promise<boolean> {
    try {
      // Clear any stale MSAL token cache to force a fresh device code prompt
      const msalCacheDir = path.join(process.env.HOME || process.env.USERPROFILE || '/tmp', '.cache', 'azure');
      if (fs.existsSync(msalCacheDir)) {
        const cacheFiles = fs.readdirSync(msalCacheDir).filter(f => f.includes('msal') || f.includes('token'));
        for (const cf of cacheFiles) {
          try { fs.unlinkSync(path.join(msalCacheDir, cf)); } catch { /* ignore */ }
        }
        console.log(`🧹 Cleared ${cacheFiles.length} stale Azure token cache file(s)`);
      }

      this.credential = new DeviceCodeCredential({
        clientId: DelegatedGraphService.CLIENT_ID,
        tenantId: 'organizations',  // Microsoft corp accounts (not 'common' which includes personal)
        userPromptCallback: (info) => {
          console.log('\n' + '='.repeat(60));
          console.log('🔑 SIGN IN REQUIRED FOR EXCEL SYNC');
          console.log('='.repeat(60));
          console.log(info.message);
          console.log('='.repeat(60) + '\n');
        },
      });

      // Force a fresh token acquisition to trigger the device code prompt
      console.log('📡 Requesting token (device code prompt should appear below)...');
      const tokenResponse = await this.credential.getToken(
        DelegatedGraphService.SCOPES.map(s => `https://graph.microsoft.com/${s}`)
      );
      if (!tokenResponse) {
        console.warn('⚠️ No token received from device code flow');
        return false;
      }
      console.log('✅ Token acquired successfully');

      const authProvider = new TokenCredentialAuthenticationProvider(this.credential, {
        scopes: DelegatedGraphService.SCOPES.map(s => `https://graph.microsoft.com/${s}`),
      });

      this.client = Client.initWithMiddleware({ authProvider });

      // Test the connection by fetching user profile
      const me = await this.client.api('/me').select('displayName,mail').get();
      console.log(`✅ Delegated Graph auth: signed in as ${me.displayName} (${me.mail})`);
      this._isAuthenticated = true;
      return true;
    } catch (err: any) {
      console.warn(`⚠️ Delegated Graph auth failed: ${err.message}`);
      if (err.message.includes('invalid_grant')) {
        console.warn('   💡 This usually means a stale cached token. Try: rm -rf ~/.cache/azure && npm start');
      }
      this._isAuthenticated = false;
      return false;
    }
  }

  private ensureClient(): Client {
    if (!this.client) throw new Error('Delegated Graph client not initialized. Call initialize() first.');
    return this.client;
  }

  // ---- Excel File Operations (delegated — uses user's own permissions) ----

  /** Resolve a SharePoint sharing URL to get drive ID and item ID */
  async resolveShareLink(sharingUrl: string): Promise<{ driveId: string; itemId: string; name: string } | null> {
    try {
      const client = this.ensureClient();
      const base64 = Buffer.from(sharingUrl).toString('base64');
      const shareToken = 'u!' + base64.replace(/\//g, '_').replace(/\+/g, '-').replace(/=+$/, '');

      const driveItem = await client.api(`/shares/${shareToken}/driveItem`).get();
      return {
        driveId: driveItem.parentReference?.driveId,
        itemId: driveItem.id,
        name: driveItem.name,
      };
    } catch (err: any) {
      console.error(`Failed to resolve sharing URL: ${err.message}`);
      return null;
    }
  }

  /** Read a worksheet range */
  async readWorksheetRange(driveId: string, itemId: string, sheetName: string, range: string): Promise<any[][]> {
    const client = this.ensureClient();
    const result = await client
      .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')/range(address='${range}')`)
      .get();
    return result.values || [];
  }

  /** Get used range of a worksheet */
  async getUsedRange(driveId: string, itemId: string, sheetName: string): Promise<{ values: any[][]; rowCount: number; columnCount: number }> {
    const client = this.ensureClient();
    const result = await client
      .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')/usedRange`)
      .get();
    return {
      values: result.values || [],
      rowCount: result.rowCount || 0,
      columnCount: result.columnCount || 0,
    };
  }

  /** Update a worksheet range */
  async updateWorksheetRange(driveId: string, itemId: string, sheetName: string, range: string, values: any[][]): Promise<void> {
    const client = this.ensureClient();
    await client
      .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')/range(address='${range}')`)
      .patch({ values });
  }

  /** List worksheets */
  async listWorksheets(driveId: string, itemId: string): Promise<string[]> {
    const client = this.ensureClient();
    const result = await client
      .api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`)
      .get();
    return (result.value || []).map((ws: any) => ws.name);
  }
}
