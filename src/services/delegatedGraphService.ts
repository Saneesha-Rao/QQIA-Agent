import { Client } from '@microsoft/microsoft-graph-client';
import * as msal from '@azure/msal-node';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Delegated Graph auth using device code flow via MSAL Node directly.
 * Uses Microsoft's built-in Azure CLI client ID — no app registration needed.
 * Authenticates as the user (delegated permissions) so it can read/write their OneDrive files.
 */
export class DelegatedGraphService {
  private client: Client | null = null;
  private msalClient: msal.PublicClientApplication | null = null;
  private accountInfo: msal.AccountInfo | null = null;
  private _isAuthenticated = false;

  // Microsoft Graph PowerShell well-known client ID — public, supports device code flow
  private static readonly CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
  private static readonly AUTHORITY = 'https://login.microsoftonline.com/organizations';
  private static readonly SCOPES = [
    'https://graph.microsoft.com/Files.ReadWrite',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
  ];

  constructor() {}

  get isAuthenticated(): boolean {
    return this._isAuthenticated;
  }

  /** Initialize with device code flow — prints a URL + code for the user to sign in */
  async initialize(): Promise<boolean> {
    try {
      const msalConfig: msal.Configuration = {
        auth: {
          clientId: DelegatedGraphService.CLIENT_ID,
          authority: DelegatedGraphService.AUTHORITY,
        },
      };

      this.msalClient = new msal.PublicClientApplication(msalConfig);

      console.log('📡 Starting device code flow...');
      const deviceCodeRequest: msal.DeviceCodeRequest = {
        scopes: DelegatedGraphService.SCOPES,
        deviceCodeCallback: (response) => {
          // Log all properties to debug what the SDK actually provides
          console.log('[DeviceCode] Full response keys:', Object.keys(response));
          console.log('[DeviceCode] Full response:', JSON.stringify(response, null, 2));
          const uri = (response as any).verificationUri || (response as any).verification_uri || (response as any).verificationUrl;
          const code = (response as any).userCode || (response as any).user_code;
          const msg = (response as any).message;
          console.log('\n' + '='.repeat(60));
          console.log('🔑 SIGN IN REQUIRED FOR EXCEL SYNC');
          console.log('='.repeat(60));
          if (msg) {
            console.log(msg);
          } else {
            console.log(`👉 Open: ${uri || 'https://microsoft.com/devicelogin'}`);
            console.log(`👉 Enter code: ${code || 'see above'}`);
          }
          console.log('='.repeat(60) + '\n');
        },
      };

      const authResult = await this.msalClient.acquireTokenByDeviceCode(deviceCodeRequest);
      if (!authResult || !authResult.accessToken) {
        console.warn('⚠️ No token received from device code flow');
        return false;
      }

      this.accountInfo = authResult.account;
      console.log(`✅ Token acquired for: ${authResult.account?.name || authResult.account?.username}`);

      // Set up Graph client with custom auth provider that refreshes tokens
      this.client = Client.init({
        authProvider: async (done) => {
          try {
            const token = await this.getAccessToken();
            done(null, token);
          } catch (err: any) {
            done(err, null);
          }
        },
      });

      // Test the connection
      const me = await this.client.api('/me').select('displayName,mail').get();
      console.log(`✅ Delegated Graph auth: signed in as ${me.displayName} (${me.mail})`);
      this._isAuthenticated = true;
      return true;
    } catch (err: any) {
      console.warn(`⚠️ Delegated Graph auth failed: ${err.message}`);
      this._isAuthenticated = false;
      return false;
    }
  }

  /** Get a valid access token (refreshes silently if cached) */
  private async getAccessToken(): Promise<string> {
    if (!this.msalClient || !this.accountInfo) {
      throw new Error('Not authenticated');
    }
    try {
      const silentResult = await this.msalClient.acquireTokenSilent({
        scopes: DelegatedGraphService.SCOPES,
        account: this.accountInfo,
      });
      return silentResult.accessToken;
    } catch {
      // Silent refresh failed — need interactive re-auth
      throw new Error('Token expired. Restart the server to re-authenticate.');
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
