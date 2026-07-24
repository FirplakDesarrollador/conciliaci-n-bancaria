# Conciliación diaria bancos

Esqueleto base de Next.js (App Router + TypeScript + Tailwind), conectado a Supabase, listo para desplegar en Vercel.

## Estructura

- `src/app/page.tsx` — home, protegida: solo se ve si hay sesión activa.
- `src/app/login/page.tsx` + `src/app/login/actions.ts` — formulario de login y Server Actions (`login`, `logout`) contra Supabase Auth.
- `src/proxy.ts` — redirige a `/login` si no hay sesión, y de `/login` a `/` si ya la hay.
- `src/lib/supabase/` — clientes de Supabase (`client.ts` navegador, `server.ts` Server Components/Actions, `middleware.ts` para el proxy).
- `src/app/api/health/route.ts` — endpoint que verifica la conexión con Supabase (excluido del login).
- `src/lib/graph/` — cliente de Microsoft Graph (`token.ts` obtiene el access token vía client credentials con MSAL, `client.ts` hace las llamadas a Graph).
- `src/app/api/sharepoint/health/route.ts` — verifica la conexión con Microsoft Graph/SharePoint (requiere sesión, como el resto de la app).
- `src/app/api/sharepoint/sites/route.ts` — lista/busca sitios de SharePoint visibles para la app (`?q=` para buscar), útil para encontrar el `id`/`webUrl` del sitio que se vaya a usar.
- `src/lib/conciliacion/` — motor de conciliación SAP B1 ↔ bancos (puerto a TypeScript del script `conciliar.py`): `sap.ts` lee el informe SAP, `readers.ts` lee cada formato de banco, `matching.ts` hace el cruce (fecha/valor exacto, tolerancia, sumas de recibos, categorías agrupadas), `reconcile.ts` orquesta todo y arma el resumen.
- `src/lib/graph/sharepoint.ts` — descarga/sube archivos de la carpeta de SharePoint `FIRPLAK 2026` (`CONCILIACION_DRIVE_ID`/`CONCILIACION_FOLDER_PATH`).
- `src/app/conciliacion/` — página protegida donde se sube el informe SAP (`Informe_de_recaudos_y_pagos.xlsx`) y se ejecuta la conciliación: descarga los 5 archivos de banco de SharePoint, corre el cruce, y sube de vuelta cada `_CONCILIADO.xlsx` + `RESUMEN_CONCILIACION.xlsx` a la misma carpeta.

### Conciliación SAP ↔ bancos

- El mapeo cuenta SAP → archivo de SharePoint está en `ACCOUNT_MAP` (`src/lib/conciliacion/config.ts`). El archivo `FIDUBOGOTÁ MAYO 2026.xlsx` de esa carpeta queda **fuera** del cruce automático (no tiene cuenta SAP asociada en el script original).
- El informe SAP **se sube manualmente** en la página por ahora; la conexión directa a SAP para traerlo automáticamente queda pendiente.
- Corre las migraciones/pruebas contra un archivo real primero: revisa el resumen (`RESUMEN_CONCILIACION.xlsx`) y los colores en cada archivo `_CONCILIADO` (verde = conciliado, amarillo = ambiguo/revisar, rojo = sin documento) antes de confiar en un resultado a ciegas.

## 1. Configuración local

```bash
npm install
npm run dev
```

Abre [http://localhost:3000](http://localhost:3000) — redirige a `/login` si no hay sesión. Verifica la conexión a Supabase en [http://localhost:3000/api/health](http://localhost:3000/api/health).

### Crear un usuario para probar el login

Este proyecto no incluye registro (signup); los usuarios se crean desde Supabase:

**Dashboard → Authentication → Users → Add user** (marca el usuario como confirmado para poder iniciar sesión de inmediato).

## 2. Variables de entorno

Ya están cargadas en `.env.local` (ignorado por git) y como referencia en `.env.example`:

```
NEXT_PUBLIC_SUPABASE_URL=https://qaqfinnckpcbkwfwbuld.supabase.co
NEXT_PUBLIC_SUPABASE_ANON_KEY=<anon key>
```

El `anon key` es una llave pública pensada para usarse desde el cliente (protegida por Row Level Security en Supabase, no por secreto). Úsala tal cual para consultas del lado del cliente o del servidor con `supabase-js`.

Si más adelante necesitas operaciones que se salten Row Level Security (tareas de servidor/admin), usa la `service_role key` de Supabase **solo en el servidor** (nunca en variables `NEXT_PUBLIC_*` ni en el navegador).

### Microsoft Graph / SharePoint

```
AZURE_AD_TENANT_ID=<directory (tenant) id>
AZURE_AD_CLIENT_ID=<application (client) id>
AZURE_AD_CLIENT_SECRET=<client secret value>
```

Credenciales del registro de app `conciliacion_app` en Azure AD (FIRPLAK SA). A diferencia del anon key de Supabase, **el client secret sí es sensible**: nunca debe llevar prefijo `NEXT_PUBLIC_`, ni comitearse, ni exponerse al navegador — solo se usa en `src/lib/graph/token.ts` (marcado `server-only`).

La app usa el flujo *client credentials* (permisos de aplicación, sin usuario interactivo), autenticada con los permisos de Graph que ya tengan consentimiento de administrador en el registro de la app (ej. `Sites.Read.All` / `Sites.ReadWrite.All`).

Si el secreto llegó a compartirse por un canal no seguro (chat, correo, etc.), regénéralo en Azure: **Registro de la app → Certificados y secretos → nuevo secreto → eliminar el anterior**.

## 3. Desplegar en Vercel

1. Sube el proyecto a un repositorio Git (GitHub/GitLab/Bitbucket) o usa `vercel` CLI desde esta carpeta.
2. Importa el proyecto en Vercel.
3. Configura en el proyecto de Vercel las variables de entorno:
   - `NEXT_PUBLIC_SUPABASE_URL`
   - `NEXT_PUBLIC_SUPABASE_ANON_KEY`
   - `AZURE_AD_TENANT_ID`
   - `AZURE_AD_CLIENT_ID`
   - `AZURE_AD_CLIENT_SECRET`
4. Despliega. Next.js en Vercel no requiere configuración adicional de build.

## Próximos pasos

Este es solo el esqueleto (Next.js + Supabase + Vercel funcionando). Falta definir la lógica real de conciliación: qué bancos, qué tablas/esquema en Supabase, qué formato de archivos se cargan, y qué reglas de cruce contra los registros contables se deben aplicar.
