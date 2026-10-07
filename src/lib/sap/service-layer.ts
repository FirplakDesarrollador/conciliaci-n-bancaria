import 'server-only';
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';


const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

// El Service Layer de SAP falla de forma intermitente (502 del proxy,
// conexiones reiniciadas). Un fallo pasajero no debe tumbar una sincronizacion
// completa: se reintenta con espera creciente antes de rendirse.
async function fetchWithRetry(url: string, init: RequestInit, attempts = 4): Promise<Response> {
  let lastError: unknown;
  for (let i = 1; i <= attempts; i++) {
    try {
      const res = await fetch(url, init);
      if (res.status !== 502 && res.status !== 503 && res.status !== 504) return res;
      lastError = new Error(`SAP respondio ${res.status}`);
      if (i === attempts) return res;
    } catch (e) {
      lastError = e;
      if (i === attempts) throw e;
    }
    await sleep(2000 * i);
  }
  throw lastError;
}

export class SAPServiceLayer {
  private static instance: SAPServiceLayer;
  private sessionId: string | null = null;
  private routeId: string | null = null;

  private constructor() {}

  public static getInstance(): SAPServiceLayer {
    if (!SAPServiceLayer.instance) {
      SAPServiceLayer.instance = new SAPServiceLayer();
    }
    return SAPServiceLayer.instance;
  }

  private async login(): Promise<void> {
    const baseUrl = process.env.SAP_BASE_URL;
    const companyDB = process.env.SAP_COMPANY_DB;
    const userName = process.env.SAP_USERNAME;
    const password = process.env.SAP_PASSWORD;

    if (!baseUrl || !companyDB || !userName || !password) {
      throw new Error('Las credenciales de SAP no están configuradas correctamente en las variables de entorno.');
    }

    try {
      const response = await fetchWithRetry(`${baseUrl}/Login?_t=${Date.now()}`, {
        method: 'POST',
        cache: 'no-store',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          CompanyDB: companyDB,
          UserName: userName,
          Password: password,
        }),
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Error en el login a SAP: ${response.status} ${response.statusText} - ${errorText}`);
      }

      // Obtener cookies de la respuesta
      const setCookieHeaders = response.headers.getSetCookie 
        ? response.headers.getSetCookie() 
        : response.headers.get('set-cookie')?.split(',') || [];

      for (const cookieStr of setCookieHeaders) {
        const cookie = cookieStr.trim();
        if (cookie.startsWith('B1SESSION=')) {
          this.sessionId = cookie.split(';')[0].split('=')[1];
        }
        if (cookie.startsWith('ROUTEID=')) {
          this.routeId = cookie.split(';')[0].split('=')[1];
        }
      }
    } catch (error) {
      console.error('Excepción al conectar con SAP Service Layer:', error);
      throw error;
    }
  }

  /**
   * Realiza una petición HTTP autenticada hacia SAP Service Layer.
   * Maneja automáticamente la reconexión si la sesión ha expirado.
   */
  public async request(endpoint: string, options: RequestInit = {}): Promise<Response> {
    if (!this.sessionId) {
      await this.login();
    }

    const baseUrl = process.env.SAP_BASE_URL;
    const url = `${baseUrl}${endpoint.startsWith('/') ? endpoint : `/${endpoint}`}`;

    const headers = new Headers(options.headers);
    headers.set('Content-Type', 'application/json');
    headers.set('Cookie', `B1SESSION=${this.sessionId}; ROUTEID=${this.routeId}`);

    const fetchOptions: RequestInit = {
      cache: 'no-store',
      ...options,
      headers,
    };

    let response = await fetchWithRetry(url, fetchOptions);

    let shouldRetry = false;
    if (response.status === 401) {
      shouldRetry = true;
    } else if (!response.ok) {
      try {
        const cloned = response.clone();
        const errData = await cloned.json();
        if (errData?.error?.code === 301 || errData?.error?.message?.value?.includes('Invalid session') || errData?.error?.code === -1000) {
          shouldRetry = true;
        }
      } catch (e) {
        // ignore
      }
    }

    // Si la sesión ha expirado, reconectar y reintentar
    if (shouldRetry) {
      console.log('Sesión de SAP expirada o inválida. Intentando reconectar...');
      await this.login();
      
      const retryHeaders = new Headers(options.headers);
      retryHeaders.set('Content-Type', 'application/json');
      retryHeaders.set('Cookie', `B1SESSION=${this.sessionId}; ROUTEID=${this.routeId}`);
      
      response = await fetchWithRetry(url, {
        ...fetchOptions,
        headers: retryHeaders,
      });
    }

    return response;
  }
}

export const sapClient = SAPServiceLayer.getInstance();
