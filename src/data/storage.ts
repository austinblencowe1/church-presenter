/**
 * Storage adapter for church-presenter state.
 *
 * Web builds use localStorage; desktop Tauri builds can later swap in the
 * SQLite-backed implementation behind the same interface.
 */

export interface MediaAsset {
  id: string;
  name: string;
  kind: 'image' | 'video';
  source: string;
  size: number;
  addedAt: string;
}

export interface ServiceSummary {
  id: string;
  title: string;
  date: string;
  folder: string;
}

export interface StorageAdapter {
  // ---- service state (single JSON blob) ----
  readServiceState(): Promise<{
    services: any[];
    activeService: number;
    theme: string;
    display: string;
    mediaUrls?: Record<string, string>;
  } | null>;
  writeServiceState(payload: {
    services: any[];
    activeService: number;
    theme: string;
    display: string;
  }): Promise<void>;

  // ---- media assets (metadata) ----
  listMedia(): Promise<MediaAsset[]>;
  saveMedia(asset: MediaAsset): Promise<void>;
  deleteMedia(id: string): Promise<void>;

  // ---- media blobs (IndexedDB) ----
  saveMediaBlob(id: string, blob: Blob): Promise<void>;
  getMediaBlob(id: string): Promise<Blob | null>;
  deleteMediaBlob(id: string): Promise<void>;
  getMediaBlobUrl(id: string): Promise<string | null>;
  revokeMediaBlobUrl(id: string): Promise<void>;

  // ---- key/value settings ----
  getItem(key: string): Promise<string | null>;
  setItem(key: string, value: string): Promise<void>;
  removeItem(key: string): Promise<void>;
}

const STORAGE_KEY = 'service-slides-state';

export class LocalStorageAdapter implements StorageAdapter {
  async readServiceState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      const saved = JSON.parse(raw);
      return {
        services: saved.services || [],
        activeService: saved.activeService ?? 0,
        theme: saved.theme || 'system',
        display: saved.display || '',
        mediaUrls: saved.mediaUrls || {},
      };
    } catch {
      return null;
    }
  }

  async writeServiceState({ services, activeService, theme, display }) {
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({ services, activeService, theme, display }),
      );
    } catch {
      /* ignore quota errors */
    }
  }

  async listMedia() {
    try {
      const raw = localStorage.getItem('service-slides-media-meta');
      if (!raw) return [];
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    } catch {
      return [];
    }
  }

  async saveMedia(asset: MediaAsset) {
    try {
      const existing = await this.listMedia();
      const next = existing.filter((a) => a.id !== asset.id);
      next.push(asset);
      localStorage.setItem(
        'service-slides-media-meta',
        JSON.stringify(next),
      );
    } catch {
      /* ignore quota errors */
    }
  }

  async deleteMedia(id: string) {
    try {
      const existing = await this.listMedia();
      const next = existing.filter((a) => a.id !== id);
      localStorage.setItem('service-slides-media-meta', JSON.stringify(next));
    } catch {
      /* ignore */
    }
  }

  async getItem(key: string) {
    try {
      return localStorage.getItem(key);
    } catch {
      return null;
    }
  }

  async setItem(key: string, value: string) {
    try {
      localStorage.setItem(key, value);
    } catch {
      /* ignore quota errors */
    }
  }

  async removeItem(key: string) {
    try {
      localStorage.removeItem(key);
    } catch {
      /* ignore */
    }
  }

  // ---- IndexedDB media blobs ----

  private mediaDbPromise: Promise<IDBDatabase | null> | null = null;
  private blobUrlCache: Map<string, string> = new Map();

  private async mediaDb(): Promise<IDBDatabase | null> {
    if (this.mediaDbPromise) return this.mediaDbPromise;
    this.mediaDbPromise = new Promise((resolve) => {
      try {
        const req = indexedDB.open('church-presenter-media-v1', 1);
        req.onupgradeneeded = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains('blobs')) {
            db.createObjectStore('blobs', { keyPath: 'id' });
          }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(null);
      } catch {
        resolve(null);
      }
    });
    return this.mediaDbPromise;
  }

  async saveMediaBlob(id: string, blob: Blob): Promise<void> {
    const db = await this.mediaDb();
    if (!db) return;
    return new Promise((resolve, reject) => {
      const tx = db.transaction('blobs', 'readwrite');
      tx.objectStore('blobs').put({ id, blob });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async getMediaBlob(id: string): Promise<Blob | null> {
    const db = await this.mediaDb();
    if (!db) return null;
    return new Promise((resolve) => {
      const tx = db.transaction('blobs', 'readonly');
      const q = tx.objectStore('blobs').get(id);
      q.onsuccess = () => resolve((q.result as { id: string; blob: Blob } | undefined)?.blob ?? null);
      q.onerror = () => resolve(null);
    });
  }

  async deleteMediaBlob(id: string): Promise<void> {
    await this.revokeMediaBlobUrl(id);
    const db = await this.mediaDb();
    if (!db) return;
    return new Promise((resolve) => {
      const tx = db.transaction('blobs', 'readwrite');
      tx.objectStore('blobs').delete(id);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    });
  }

  async getMediaBlobUrl(id: string): Promise<string | null> {
    const cached = this.blobUrlCache.get(id);
    if (cached) return cached;
    const blob = await this.getMediaBlob(id);
    if (!blob) return null;
    const url = URL.createObjectURL(blob);
    this.blobUrlCache.set(id, url);
    return url;
  }

  async revokeMediaBlobUrl(id: string): Promise<void> {
    const url = this.blobUrlCache.get(id);
    if (url) {
      try { URL.revokeObjectURL(url); } catch {}
      this.blobUrlCache.delete(id);
    }
  }
}

export const storage = new LocalStorageAdapter();
