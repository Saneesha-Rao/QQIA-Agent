import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { config } from '../config/appConfig';

/**
 * Microsoft Graph API client for SharePoint/OneDrive Excel operations.
 * Handles authentication and provides methods for reading/writing Excel files.
 */
export class GraphService {
  private client: Client | null = null;

  /** Initialize the Graph client with app-only (client credentials) auth */
  async initialize(): Promise<void> {
    const credential = new ClientSecretCredential(
      config.graph.tenantId,
      config.graph.clientId,
      config.graph.clientSecret
    );

    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default'],
    });

    this.client = Client.initWithMiddleware({ authProvider });
    console.log('✅ Graph API client initialized');
  }

  private ensureClient(): Client {
    if (!this.client) throw new Error('Graph client not initialized. Call initialize() first.');
    return this.client;
  }

  // ---- Excel File Operations via Graph API ----

  /** Download the Excel file as a buffer from OneDrive/SharePoint */
  async downloadExcelFile(): Promise<Buffer> {
    const client = this.ensureClient();

    if (config.graph.excelItemId) {
      // Use item ID directly (most reliable)
      const stream = await client
        .api(`/drives/${config.graph.excelDriveId}/items/${config.graph.excelItemId}/content`)
        .getStream();
      return await this.streamToBuffer(stream);
    }

    // Fallback: use the user's OneDrive path
    const stream = await client
      .api(`/me/drive/root:${config.graph.excelFilePath}:/content`)
      .getStream();
    return await this.streamToBuffer(stream);
  }

  /** Upload the modified Excel file back to OneDrive/SharePoint */
  async uploadExcelFile(buffer: Buffer): Promise<void> {
    const client = this.ensureClient();

    if (config.graph.excelItemId) {
      await client
        .api(`/drives/${config.graph.excelDriveId}/items/${config.graph.excelItemId}/content`)
        .putStream(buffer);
    } else {
      await client
        .api(`/me/drive/root:${config.graph.excelFilePath}:/content`)
        .put(buffer);
    }
    console.log('✅ Excel file uploaded to OneDrive/SharePoint');
  }

  /** Read a specific worksheet range using Excel REST API (no download needed) */
  async readWorksheetRange(sheetName: string, range: string): Promise<any[][]> {
    const client = this.ensureClient();
    const itemRef = config.graph.excelItemId
      ? `/drives/${config.graph.excelDriveId}/items/${config.graph.excelItemId}`
      : `/me/drive/root:${config.graph.excelFilePath}:`;

    const result = await client
      .api(`${itemRef}/workbook/worksheets('${sheetName}')/range(address='${range}')`)
      .get();

    return result.values || [];
  }

  /** Update a specific worksheet range using Excel REST API */
  async updateWorksheetRange(sheetName: string, range: string, values: any[][]): Promise<void> {
    const client = this.ensureClient();
    const itemRef = config.graph.excelItemId
      ? `/drives/${config.graph.excelDriveId}/items/${config.graph.excelItemId}`
      : `/me/drive/root:${config.graph.excelFilePath}:`;

    await client
      .api(`${itemRef}/workbook/worksheets('${sheetName}')/range(address='${range}')`)
      .patch({ values });
  }

  /** Get the used range of a worksheet */
  async getUsedRange(sheetName: string): Promise<{ values: any[][]; rowCount: number; columnCount: number }> {
    const client = this.ensureClient();
    const itemRef = config.graph.excelItemId
      ? `/drives/${config.graph.excelDriveId}/items/${config.graph.excelItemId}`
      : `/me/drive/root:${config.graph.excelFilePath}:`;

    const result = await client
      .api(`${itemRef}/workbook/worksheets('${sheetName}')/usedRange`)
      .get();

    return {
      values: result.values || [],
      rowCount: result.rowCount || 0,
      columnCount: result.columnCount || 0,
    };
  }

  /** List all worksheets in the workbook */
  async listWorksheets(): Promise<string[]> {
    const client = this.ensureClient();
    const itemRef = config.graph.excelItemId
      ? `/drives/${config.graph.excelDriveId}/items/${config.graph.excelItemId}`
      : `/me/drive/root:${config.graph.excelFilePath}:`;

    const result = await client
      .api(`${itemRef}/workbook/worksheets`)
      .get();

    return (result.value || []).map((ws: any) => ws.name);
  }

  // ---- Proactive Messaging Helpers ----

  /** Look up a user's Teams ID by email for proactive messaging */
  async getUserTeamsId(email: string): Promise<string | null> {
    try {
      const client = this.ensureClient();
      const user = await client.api(`/users/${email}`).select('id').get();
      return user.id || null;
    } catch {
      return null;
    }
  }

  /** Send a proactive 1:1 chat message to a user (requires conversation reference) */
  async getInstallationForUser(userId: string, teamsAppId: string): Promise<any> {
    try {
      const client = this.ensureClient();
      const result = await client
        .api(`/users/${userId}/teamwork/installedApps`)
        .filter(`teamsApp/id eq '${teamsAppId}'`)
        .expand('teamsApp')
        .get();
      return result.value?.[0] || null;
    } catch {
      return null;
    }
  }

  /** Post a message to a Teams channel */
  async postChannelMessage(teamId: string, channelId: string, content: string): Promise<void> {
    const client = this.ensureClient();
    await client
      .api(`/teams/${teamId}/channels/${channelId}/messages`)
      .post({
        body: { contentType: 'html', content },
      });
  }

  // ---- Utility ----

  private async streamToBuffer(stream: NodeJS.ReadableStream): Promise<Buffer> {
    const chunks: Buffer[] = [];
    return new Promise((resolve, reject) => {
      stream.on('data', (chunk: Buffer) => chunks.push(chunk));
      stream.on('end', () => resolve(Buffer.concat(chunks)));
      stream.on('error', reject);
    });
  }
}
