/**
 * Google Drive client for web (non-Tauri) builds.
 *
 * Uses Google Identity Services (GIS) for browser-based OAuth 2.0 access tokens.
 * Tokens are held in memory (sessionStorage) — no refresh tokens, so re-consent
 * is needed after ~1 hour or when the token expires.
 *
 * Requires the GIS script loaded in index.html:
 *   <script src="https://accounts.google.com/gsi/client" async defer></script>
 * And a Google Cloud OAuth client ID configured for a Web application with
 * the origin of the deployed site added as an authorised JavaScript origin.
 */

export interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
  size?: string;
  modifiedTime?: string;
  webViewLink?: string;
  webContentLink?: string;
  thumbnailLink?: string;
  trashed?: boolean;
  version?: string;
  headRevisionId?: string;
  canEdit?: boolean;
  appProperties?: Record<string, string>;
}

export interface DriveStatus {
  connected: boolean;
  email?: string;
  error?: string;
}

const CLIENT_ID_KEY = 'church-presenter-gdrive-web-client-id';
const TOKEN_KEY = 'church-presenter-gdrive-web-access-token';
const TOKEN_EXPIRY_KEY = 'church-presenter-gdrive-web-token-expiry';
const SCOPE = 'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/userinfo.email';

let clientId: string | null = null;
let accessToken: string | null = null;
let tokenExpiry = 0;

function isGsiAvailable(): boolean {
  return typeof (window as any).google !== 'undefined'
    && typeof (window as any).google.accounts !== 'undefined'
    && typeof (window as any).google.accounts.oauth2 !== 'undefined';
}

function loadClientId(): string | null {
  try {
    const v = localStorage.getItem(CLIENT_ID_KEY);
    return v || null;
  } catch {
    return null;
  }
}

export function getConfiguredClientId(): string {
  if (!clientId) clientId = loadClientId();
  return clientId || '';
}

function saveClientId(id: string): void {
  try { localStorage.setItem(CLIENT_ID_KEY, id); } catch {}
}

function loadToken(): { token: string; expiry: number } | null {
  try {
    const raw = localStorage.getItem(TOKEN_KEY);
    const exp = localStorage.getItem(TOKEN_EXPIRY_KEY);
    if (!raw || !exp) return null;
    return { token: raw, expiry: parseInt(exp, 10) || 0 };
  } catch {
    return null;
  }
}

function saveToken(token: string, expiresInSeconds: number): void {
  accessToken = token;
  tokenExpiry = Date.now() + expiresInSeconds * 1000;
  try {
    localStorage.setItem(TOKEN_KEY, token);
    localStorage.setItem(TOKEN_EXPIRY_KEY, String(tokenExpiry));
  } catch {}
}

function clearToken(): void {
  accessToken = null;
  tokenExpiry = 0;
  try {
    localStorage.removeItem(TOKEN_KEY);
    localStorage.removeItem(TOKEN_EXPIRY_KEY);
  } catch {}
}

function tokenValid(): boolean {
  return !!accessToken && tokenExpiry > Date.now() + 60000; // 1 minute skew
}

async function getUserEmail(token: string): Promise<string | null> {
  try {
    const res = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) return null;
    const data = await res.json();
    return data.email || null;
  } catch {
    return null;
  }
}

export async function driveConfigure(clientIdValue: string): Promise<void> {
  const nextClientId = clientIdValue.trim() || null;
  const previousClientId = clientId || loadClientId();
  if (previousClientId && previousClientId !== nextClientId) clearToken();
  clientId = nextClientId;
  if (clientId) saveClientId(clientId);
  else {
    try { localStorage.removeItem(CLIENT_ID_KEY); } catch {}
  }
  // Try to restore an existing token
  const stored = loadToken();
  if (stored && stored.expiry > Date.now()) {
    accessToken = stored.token;
    tokenExpiry = stored.expiry;
  }
}

export async function driveConnect(): Promise<DriveStatus> {
  if (!clientId) {
    clientId = loadClientId();
  }
  if (!clientId) {
    return { connected: false, error: 'No Google Client ID configured. Set gdrive client id in Settings.' };
  }

  // If we already have a valid token, just return status
  if (tokenValid()) {
    const email = await getUserEmail(accessToken!);
    return { connected: true, email };
  }

  if (!isGsiAvailable()) {
    return {
      connected: false,
      error: 'Google Identity Services not loaded. Add <script src="https://accounts.google.com/gsi/client"> to index.html.',
    };
  }

  return new Promise((resolve) => {
    try {
      const client = (window as any).google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: SCOPE,
        prompt: 'consent',
        callback: (response: any) => {
          if (response.error) {
            resolve({ connected: false, error: response.error || 'Google sign-in failed' });
            return;
          }
          if (response.access_token) {
            const expiresIn = response.expires_in || 3600;
            saveToken(response.access_token, expiresIn);
            getUserEmail(response.access_token).then((email) => {
              resolve({ connected: true, email });
            }).catch(() => resolve({ connected: true }));
          } else {
            resolve({ connected: false, error: 'No access token returned' });
          }
        },
      });
      client.requestAccessToken();
    } catch (err) {
      resolve({ connected: false, error: err instanceof Error ? err.message : 'Sign-in failed' });
    }
  });
}

export async function driveDisconnect(): Promise<void> {
  if (accessToken && isGsiAvailable()) {
    try {
      (window as any).google.accounts.oauth2.revoke(accessToken);
    } catch {}
  }
  clearToken();
  clientId = null;
}

export async function driveStatus(): Promise<DriveStatus> {
  if (!clientId) clientId = loadClientId();
  if (!clientId) return { connected: false };
  if (!accessToken) {
    const stored = loadToken();
    if (stored && stored.expiry > Date.now()) {
      accessToken = stored.token;
      tokenExpiry = stored.expiry;
    }
  }
  if (!tokenValid()) return { connected: false };
  const email = await getUserEmail(accessToken!);
  return { connected: true, email };
}

async function driveFetch(path: string, init?: RequestInit): Promise<Response> {
  if (!tokenValid()) {
    throw new Error('Not connected to Google Drive. Click Connect Google Drive.');
  }
  const res = await fetch(`https://www.googleapis.com${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(init?.headers || {}),
    },
  });
  if (res.status === 401) {
    clearToken();
    throw new Error('Drive access expired. Reconnect to Google Drive.');
  }
  return res;
}

export async function driveListFiles(folderId?: string): Promise<DriveFile[]> {
  const params = new URLSearchParams({
    fields: 'files(id,name,mimeType,size,modifiedTime,webViewLink,webContentLink,thumbnailLink,trashed,version,headRevisionId,capabilities(canEdit),appProperties)',
    orderBy: 'modifiedTime desc',
    pageSize: '1000',
    supportsAllDrives: 'true',
    includeItemsFromAllDrives: 'true',
  });
  if (folderId) params.set('q', `'${folderId.replace(/\\/g, '\\\\').replace(/'/g, "\\'")}' in parents and trashed = false`);
  const res = await driveFetch(
    `/drive/v3/files?${params.toString()}`
  );
  if (!res.ok) throw new Error(`Drive list failed: ${res.status}`);
  const data = await res.json();
  return (data.files || []).map((f: any) => ({
    id: f.id,
    name: f.name,
    mimeType: f.mimeType,
    size: f.size,
    modifiedTime: f.modifiedTime,
    webViewLink: f.webViewLink,
    webContentLink: f.webContentLink,
    thumbnailLink: f.thumbnailLink,
    trashed: f.trashed,
    version: f.version,
    headRevisionId: f.headRevisionId,
    canEdit: f.capabilities?.canEdit,
    appProperties: f.appProperties,
  }));
}

export async function driveUploadFile(
  file: Blob,
  name: string,
  folderId?: string,
  onProgress?: (bytesSent: number, totalBytes: number) => void,
  options: { existingFileId?: string; appProperties?: Record<string, string> } = {}
): Promise<{ id: string; name: string; modifiedTime: string; webViewLink?: string; version?: string; headRevisionId?: string }> {
  const size = file.size;
  const metadata = {
    name,
    ...(folderId && !options.existingFileId ? { parents: [folderId] } : {}),
    ...(options.appProperties ? { appProperties: options.appProperties } : {}),
  };
  if (!tokenValid()) throw new Error('Not connected to Google Drive. Click Connect Google Drive.');

  // Initiate resumable upload
  const sessionRes = await fetch(
    `https://www.googleapis.com/upload/drive/v3/files${options.existingFileId ? `/${encodeURIComponent(options.existingFileId)}` : ''}?uploadType=resumable&supportsAllDrives=true&fields=id,name,modifiedTime,webViewLink,version,headRevisionId`,
    {
      method: options.existingFileId ? 'PATCH' : 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json; charset=UTF-8',
        'X-Upload-Content-Type': file.type || 'application/octet-stream',
        'X-Upload-Content-Length': String(size),
      },
      body: JSON.stringify(metadata),
    }
  );
  if (!sessionRes.ok) {
    const body = await sessionRes.text();
    throw new Error(`Drive upload session failed: ${sessionRes.status} ${body}`);
  }
  const sessionUrl = sessionRes.headers.get('Location');
  if (!sessionUrl) throw new Error('Drive did not return an upload URL.');

  const CHUNK_SIZE = 8 * 1024 * 1024;
  let offset = 0;
  let attempt = 0;

  while (offset < size || size === 0) {
    const end = Math.min(offset + CHUNK_SIZE, size);
    const chunk = file.slice(offset, end);

    const res = await fetch(sessionUrl, {
      method: 'PUT',
      headers: {
        'Content-Range': size ? `bytes ${offset}-${end - 1}/${size}` : 'bytes */0',
      },
      body: chunk,
    });

    if (res.status === 308) {
      const range = res.headers.get('Range');
      const received = range ? parseInt(range.replace('bytes=0-', ''), 10) + 1 : 0;
      if (received <= offset || received > end || size === 0) {
        throw new Error('Drive did not acknowledge upload progress.');
      }
      offset = received;
      onProgress?.(offset, size);
      attempt = 0;
      continue;
    }

    if (res.ok) {
      const data = await res.json();
      onProgress?.(size, size);
      return {
        id: data.id,
        name: data.name,
        modifiedTime: data.modifiedTime || '',
        webViewLink: data.webViewLink,
        version: data.version,
        headRevisionId: data.headRevisionId,
      };
    }

    if (attempt < 3 && (res.status >= 500 || res.status === 429)) {
      attempt++;
      await new Promise((r) => setTimeout(r, 1000 * attempt));
      continue;
    }

    const body = await res.text();
    throw new Error(`Drive upload failed: ${res.status} ${body}`);
  }

  throw new Error('Drive upload did not complete.');
}

export async function driveDownloadFile(fileId: string, mimeType = ''): Promise<Blob> {
  const id = encodeURIComponent(fileId);
  const url = mimeType === 'application/vnd.google-apps.presentation'
    ? `/drive/v3/files/${id}/export?mimeType=${encodeURIComponent('application/vnd.openxmlformats-officedocument.presentationml.presentation')}`
    : `/drive/v3/files/${id}?alt=media&supportsAllDrives=true`;
  const res = await driveFetch(url);
  if (res.ok) return res.blob();
  throw new Error(`Drive download failed: ${res.status} ${await res.text()}`);
}

export async function driveGetFile(fileId: string): Promise<DriveFile> {
  const fields = 'id,name,mimeType,size,modifiedTime,webViewLink,webContentLink,thumbnailLink,trashed,version,headRevisionId,capabilities(canEdit,canDelete,canTrash),appProperties,owners(emailAddress)';
  const res = await driveFetch(`/drive/v3/files/${encodeURIComponent(fileId)}?fields=${encodeURIComponent(fields)}&supportsAllDrives=true`);
  if (!res.ok) throw new Error(`Drive file lookup failed: ${res.status} ${await res.text()}`);
  const f = await res.json();
  return {
    ...f,
    canEdit: f.capabilities?.canEdit,
  };
}

export async function driveEnsureFolder(parentId: string, name: string): Promise<{ id: string; name: string }> {
  const q = `'${parentId.replace(/\\/g, '\\\\').replace(/'/g, "\\'")}' in parents and name = '${name.replace(/\\/g, '\\\\').replace(/'/g, "\\'")}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`;
  const params = new URLSearchParams({ q, fields: 'files(id,name)', supportsAllDrives: 'true', includeItemsFromAllDrives: 'true' });
  const list = await driveFetch(`/drive/v3/files?${params.toString()}`);
  if (!list.ok) throw new Error(`Drive folder lookup failed: ${list.status} ${await list.text()}`);
  const existing = (await list.json()).files?.[0];
  if (existing) return existing;
  const create = await driveFetch('/drive/v3/files?supportsAllDrives=true&fields=id,name', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] }),
  });
  if (!create.ok) throw new Error(`Drive folder creation failed: ${create.status} ${await create.text()}`);
  return create.json();
}

export async function driveDeleteFile(fileId: string): Promise<void> {
  const id = encodeURIComponent(fileId);
  const res = await driveFetch(`/drive/v3/files/${id}?supportsAllDrives=true`, { method: 'DELETE' });
  if (res.ok || res.status === 204) return;
  const fallback = await driveFetch(`/drive/v3/files/${id}?supportsAllDrives=true`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ trashed: true }),
  });
  if (!fallback.ok) throw new Error(`Drive delete failed: ${fallback.status} ${await fallback.text()}`);
}

export async function driveDownloadThumbnail(fileId: string): Promise<Blob> {
  const file = await driveGetFile(fileId);
  if (!file.thumbnailLink) throw new Error('Drive did not provide a thumbnail for this file.');
  const res = await fetch(file.thumbnailLink, { headers: { Authorization: `Bearer ${accessToken}` } });
  if (!res.ok) throw new Error(`Drive thumbnail download failed: ${res.status}`);
  return res.blob();
}

export const driveWeb = {
  configure: driveConfigure,
  connect: driveConnect,
  disconnect: driveDisconnect,
  status: driveStatus,
  listFiles: driveListFiles,
  getFile: driveGetFile,
  ensureFolder: driveEnsureFolder,
  deleteFile: driveDeleteFile,
  downloadThumbnail: driveDownloadThumbnail,
  uploadFile: driveUploadFile,
  downloadFile: driveDownloadFile,
};
