import { invoke, isTauri } from "@tauri-apps/api/core";
import type { MediaAsset, Service, ServiceSummary, Song, SlideTheme } from "../types/presentation";

const storageKey = "church-presenter-services";
const songsStorageKey = "church-presenter-songs";
const mediaDatabaseName = "church-presenter-media";
const mediaObjectStore = "assets";
const themesStorageKey = "church-presenter-slide-themes";

function openMediaDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(mediaDatabaseName, 1);
    request.addEventListener("upgradeneeded", () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(mediaObjectStore)) {
        database.createObjectStore(mediaObjectStore, { keyPath: "id" });
      }
    });
    request.addEventListener("success", () => resolve(request.result));
    request.addEventListener("error", () => reject(request.error ?? new Error("Could not open the media library")));
  });
}

async function browserMediaRequest<T>(mode: IDBTransactionMode, run: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const database = await openMediaDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(mediaObjectStore, mode);
    const request = run(transaction.objectStore(mediaObjectStore));
    let result: T;
    let hasResult = false;
    request.addEventListener("success", () => {
      result = request.result;
      hasResult = true;
    });
    request.addEventListener("error", () => reject(request.error ?? new Error("Media library operation failed")));
    transaction.addEventListener("complete", () => {
      database.close();
      if (hasResult) resolve(result);
      else reject(new Error("Media library transaction completed without a result"));
    });
    transaction.addEventListener("abort", () => {
      database.close();
      reject(transaction.error ?? new Error("Media library transaction was interrupted"));
    });
  });
}

function readBrowserServices(): Service[] {
  const value = localStorage.getItem(storageKey);
  if (!value) return [];
  try {
    const parsed: unknown = JSON.parse(value);
    return Array.isArray(parsed) ? parsed as Service[] : [];
  } catch {
    return [];
  }
}

function writeBrowserServices(services: Service[]) {
  localStorage.setItem(storageKey, JSON.stringify(services));
}

function readBrowserSongs(): Song[] {
  const value = localStorage.getItem(songsStorageKey);
  if (!value) return [];
  try {
    const parsed: unknown = JSON.parse(value);
    return Array.isArray(parsed) ? parsed as Song[] : [];
  } catch {
    return [];
  }
}

function writeBrowserSongs(songs: Song[]) {
  localStorage.setItem(songsStorageKey, JSON.stringify(songs));
}

function readBrowserThemes(): SlideTheme[] {
  const value = localStorage.getItem(themesStorageKey);
  if (!value) return [];
  try {
    const parsed: unknown = JSON.parse(value);
    return Array.isArray(parsed) ? parsed as SlideTheme[] : [];
  } catch {
    return [];
  }
}

function applyThemeToBrowserSlides(theme: SlideTheme) {
  const services = readBrowserServices().map((service) => ({
    ...service,
    items: service.items.map((item) => ({
      ...item,
      slides: item.slides.map((slide) => slide.themeId === theme.id ? {
        ...slide,
        background: theme.background,
        fontSize: theme.fontSize,
        textAlign: theme.textAlign,
        transition: theme.transition,
        transitionDuration: theme.transitionDuration,
      } : slide),
    })),
  }));
  writeBrowserServices(services);
}

export async function listServices(): Promise<ServiceSummary[]> {
  if (isTauri()) return invoke<ServiceSummary[]>("list_services");
  return readBrowserServices().map(({ id, title, date, folder }) => ({ id, title, date, folder: folder ?? "" }));
}

export async function loadService(serviceId: string): Promise<Service> {
  if (isTauri()) return invoke<Service>("load_service", { serviceId });
  const service = readBrowserServices().find((item) => item.id === serviceId);
  if (!service) throw new Error("Service not found");
  return service;
}

export async function saveService(service: Service): Promise<void> {
  if (isTauri()) {
    await invoke("save_service", { service });
    return;
  }
  const services = readBrowserServices();
  const index = services.findIndex((item) => item.id === service.id);
  if (index < 0) services.push(service);
  else services[index] = service;
  writeBrowserServices(services);
}

export async function deleteService(serviceId: string): Promise<void> {
  if (isTauri()) {
    await invoke("delete_service", { serviceId });
    return;
  }
  writeBrowserServices(readBrowserServices().filter((item) => item.id !== serviceId));
}

export async function listSongs(): Promise<Song[]> {
  if (isTauri()) return invoke<Song[]>("list_songs");
  return readBrowserSongs();
}

export async function saveSong(song: Song): Promise<void> {
  if (isTauri()) {
    await invoke("save_song", { song });
    return;
  }
  const songs = readBrowserSongs();
  const index = songs.findIndex((item) => item.id === song.id);
  if (index < 0) songs.push(song);
  else songs[index] = song;
  writeBrowserSongs(songs);
}

export async function deleteSong(songId: string): Promise<void> {
  if (isTauri()) {
    await invoke("delete_song", { songId });
    return;
  }
  writeBrowserSongs(readBrowserSongs().filter((song) => song.id !== songId));
}

export async function listMedia(): Promise<MediaAsset[]> {
  if (isTauri()) return invoke<MediaAsset[]>("list_media");
  return browserMediaRequest("readonly", (store) => store.getAll());
}

export async function saveMedia(asset: MediaAsset): Promise<void> {
  if (isTauri()) {
    await invoke("save_media", { asset });
    return;
  }
  await browserMediaRequest("readwrite", (store) => store.put(asset));
}

export async function deleteMedia(mediaId: string): Promise<void> {
  if (isTauri()) {
    await invoke("delete_media", { mediaId });
    return;
  }
  await browserMediaRequest("readwrite", (store) => store.delete(mediaId));
}


export async function listSlideThemes(): Promise<SlideTheme[]> {
  if (isTauri()) return invoke<SlideTheme[]>("list_slide_themes");
  return readBrowserThemes();
}

export async function saveSlideTheme(theme: SlideTheme): Promise<void> {
  if (isTauri()) {
    await invoke("save_slide_theme", { theme });
    return;
  }
  const themes = readBrowserThemes();
  const index = themes.findIndex((entry) => entry.id === theme.id);
  if (index < 0) themes.push(theme);
  else themes[index] = theme;
  localStorage.setItem(themesStorageKey, JSON.stringify(themes));
  applyThemeToBrowserSlides(theme);
}

export async function deleteSlideTheme(themeId: string): Promise<void> {
  if (isTauri()) {
    await invoke("delete_slide_theme", { themeId });
    return;
  }
  localStorage.setItem(themesStorageKey, JSON.stringify(readBrowserThemes().filter((theme) => theme.id !== themeId)));
  writeBrowserServices(readBrowserServices().map((service) => ({
    ...service,
    items: service.items.map((item) => ({
      ...item,
      slides: item.slides.map((slide) => slide.themeId === themeId ? { ...slide, themeId: null } : slide),
    })),
  })));
}
