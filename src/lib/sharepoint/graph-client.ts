import { ConfidentialClientApplication } from '@azure/msal-node';

export class GraphClient {
  private static instance: GraphClient;
  private pca: ConfidentialClientApplication;
  private accessToken: string | null = null;
  private tokenExpiresAt: number = 0;

  private tenantId = process.env.AZURE_AD_TENANT_ID;
  private clientId = process.env.AZURE_AD_CLIENT_ID;
  private clientSecret = process.env.AZURE_AD_CLIENT_SECRET;

  private constructor() {
    if (!this.tenantId || !this.clientId || !this.clientSecret) {
      throw new Error('Azure AD credentials missing in environment variables.');
    }

    this.pca = new ConfidentialClientApplication({
      auth: {
        clientId: this.clientId,
        authority: `https://login.microsoftonline.com/${this.tenantId}`,
        clientSecret: this.clientSecret,
      }
    });
  }

  public static getInstance(): GraphClient {
    if (!GraphClient.instance) {
      GraphClient.instance = new GraphClient();
    }
    return GraphClient.instance;
  }

  private async getToken(): Promise<string> {
    if (this.accessToken && Date.now() < this.tokenExpiresAt) {
      return this.accessToken;
    }

    const authResult = await this.pca.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default'],
    });

    if (!authResult || !authResult.accessToken) {
      throw new Error('Failed to acquire token from Azure AD.');
    }

    this.accessToken = authResult.accessToken;
    // Buffer de 5 minutos antes de que expire
    this.tokenExpiresAt = authResult.expiresOn ? authResult.expiresOn.getTime() - 5 * 60 * 1000 : Date.now() + 55 * 60 * 1000;
    return this.accessToken;
  }

  private async request(url: string, options: RequestInit = {}): Promise<Response> {
    const token = await this.getToken();
    const headers = new Headers(options.headers);
    headers.set('Authorization', `Bearer ${token}`);
    
    return fetch(url, { ...options, headers });
  }

  public async getSiteId(siteSearch: string = 'FPK Tesoreria'): Promise<string> {
    const res = await this.request(`https://graph.microsoft.com/v1.0/sites?search=${siteSearch}`);
    if (!res.ok) throw new Error(`Failed to find site: ${await res.text()}`);
    const data = await res.json();
    const site = data.value.find((s: any) => s.displayName.toLowerCase().includes('tesoreria'));
    if (!site) throw new Error('Site not found');
    return site.id;
  }

  public async getDriveId(siteId: string, driveName: string = 'Documents'): Promise<string> {
    const res = await this.request(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`);
    if (!res.ok) throw new Error(`Failed to get drives: ${await res.text()}`);
    const data = await res.json();
    const drive = data.value.find((d: any) => d.name === driveName || d.name === 'Documentos') || data.value[0];
    if (!drive) throw new Error('Drive not found');
    return drive.id;
  }

  public async getFolderFiles(driveId: string, folderPath: string): Promise<any[]> {
    const encodedPath = encodeURIComponent(folderPath);
    const res = await this.request(`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedPath}:/children`);
    if (!res.ok) throw new Error(`Failed to get folder files: ${await res.text()}`);
    const data = await res.json();
    return data.value;
  }

  public async downloadFile(driveId: string, fileId: string): Promise<ArrayBuffer> {
    const res = await this.request(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`);
    if (!res.ok) throw new Error(`Failed to download file: ${await res.text()}`);
    return await res.arrayBuffer();
  }

  public async uploadFile(driveId: string, fileId: string, buffer: ArrayBuffer): Promise<any> {
    const res = await this.request(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      },
      body: buffer as any
    });
    
    if (!res.ok) throw new Error(`Failed to upload file: ${await res.text()}`);
    return await res.json();
  }

  public async createOrUpdateFile(driveId: string, folderId: string, fileName: string, buffer: ArrayBuffer): Promise<any> {
    const res = await this.request(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}:/${encodeURIComponent(fileName)}:/content`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      },
      body: buffer as any
    });
    
    if (!res.ok) throw new Error(`Failed to create file: ${await res.text()}`);
    return await res.json();
  }
}

export const graphClient = GraphClient.getInstance();
