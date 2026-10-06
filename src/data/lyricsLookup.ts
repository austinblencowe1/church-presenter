/**
 * Online lyric lookup using LRCLIB (https://lrclib.net) - free, no account or API key,
 * nothing to install. Search results already contain the full lyrics, so choosing a
 * result needs no second request.
 *
 * Speed/reliability notes:
 *  - the title search and the general search run in parallel and results appear as soon as the first returns
 *  - every request has a timeout and can be cancelled when the user keeps typing
 *  - results are cached in memory and in localStorage, so repeat searches are instant
 */

export interface LyricsResult {
  readonly id: string;
  readonly title: string;
  readonly artist: string;
  readonly album?: string;
  readonly lyrics: string;
  readonly snippet: string;
}

interface LrclibItem {
  id?: number;
  name?: string;
  trackName?: string;
  artistName?: string;
  albumName?: string;
  instrumental?: boolean;
  plainLyrics?: string | null;
  syncedLyrics?: string | null;
}

export interface SearchOptions {
  signal?: AbortSignal;
  /** Called each time more results arrive, before the whole search has finished. */
  onResults?: (results: LyricsResult[]) => void;
}

const ENDPOINT = "https://lrclib.net/api/search";
const REQUEST_TIMEOUT_MS = 10_000;
const MAX_RESULTS = 12;
const STORAGE_KEY = "church-presenter.lyrics-search-cache.v1";
const STORAGE_LIMIT = 20;

const memoryCache = new Map<string, LyricsResult[]>();

function cacheKey(query: string): string {
  return query.trim().toLowerCase().replace(/\s+/g, " ");
}

function normalize(text: string): string {
  return text
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\([^)]*\)|\[[^\]]*\]/g, " ")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function stripTimestamps(synced: string): string {
  return synced
    .split("\n")
    .map((line) => line.replace(/^\[\d+:\d+(?:\.\d+)?\]\s*/, "").trim())
    .join("\n");
}

function tidyLyrics(text: string): string {
  return text.replace(/\r/g, "").replace(/[ \t]+$/gm, "").replace(/\n{3,}/g, "\n\n").trim();
}

function toResult(item: LrclibItem): LyricsResult | null {
  if (!item || item.instrumental) return null;
  const raw = item.plainLyrics?.trim() || (item.syncedLyrics ? stripTimestamps(item.syncedLyrics) : "");
  const lyrics = tidyLyrics(raw);
  if (lyrics.length < 30) return null;
  const title = (item.trackName || item.name || "").trim();
  if (!title) return null;
  const snippet = lyrics.split("\n").map((line) => line.trim()).filter(Boolean).slice(0, 3).join(" / ");
  return {
    id: `lrclib_${item.id ?? `${title}_${item.artistName ?? ""}`}`,
    title,
    artist: (item.artistName || "").trim(),
    album: item.albumName?.trim() || undefined,
    lyrics,
    snippet,
  };
}

function addResult(into: Map<string, LyricsResult>, result: LyricsResult) {
  // The same song is often uploaded many times; keep the most complete copy.
  const key = `${normalize(result.title)}|${normalize(result.artist)}`;
  const existing = into.get(key);
  if (!existing || result.lyrics.length > existing.lyrics.length) into.set(key, result);
}

function score(result: LyricsResult, query: string): number {
  const q = normalize(query);
  const title = normalize(result.title);
  const artist = normalize(result.artist);
  let value = 0;
  if (title === q) value += 100;
  else if (title.startsWith(q)) value += 60;
  else if (title.includes(q)) value += 40;
  if (artist && q.includes(artist)) value += 25;
  const words = q.split(" ").filter(Boolean);
  if (words.length > 0) {
    const haystack = `${title} ${artist}`;
    value += (words.filter((word) => haystack.includes(word)).length / words.length) * 30;
  }
  return value;
}

function rank(results: Iterable<LyricsResult>, query: string): LyricsResult[] {
  return [...results]
    .map((result, index) => ({ result, index, value: score(result, query) }))
    .sort((a, b) => b.value - a.value || a.index - b.index)
    .slice(0, MAX_RESULTS)
    .map((entry) => entry.result);
}

interface StoredEntry { key: string; results: LyricsResult[] }

function readStore(): StoredEntry[] {
  try {
    const parsed: unknown = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "[]");
    return Array.isArray(parsed) ? (parsed as StoredEntry[]) : [];
  } catch {
    return [];
  }
}

function writeStore(key: string, results: LyricsResult[]) {
  try {
    const entries = [{ key, results }, ...readStore().filter((entry) => entry.key !== key)].slice(0, STORAGE_LIMIT);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
  } catch {
    // Storage full or unavailable - caching is only an optimisation.
  }
}

/** Instant results for a query that was searched before (this session or an earlier one). */
export function peekCachedLyrics(query: string): LyricsResult[] | null {
  const key = cacheKey(query);
  const hit = memoryCache.get(key);
  if (hit) return hit;
  const stored = readStore().find((entry) => entry.key === key);
  if (stored && Array.isArray(stored.results) && stored.results.length > 0) {
    memoryCache.set(key, stored.results);
    return stored.results;
  }
  return null;
}

async function fetchItems(params: Record<string, string>, external?: AbortSignal): Promise<LrclibItem[]> {
  const controller = new AbortController();
  let timedOut = false;
  const onAbort = () => controller.abort();
  external?.addEventListener("abort", onAbort);
  const timer = window.setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, REQUEST_TIMEOUT_MS);
  try {
    const response = await fetch(`${ENDPOINT}?${new URLSearchParams(params).toString()}`, { signal: controller.signal });
    if (!response.ok) throw new LyricsSearchError(`The lyric database returned an error (${response.status}).`);
    const data: unknown = await response.json();
    return Array.isArray(data) ? (data as LrclibItem[]) : [];
  } catch (error) {
    if (external?.aborted) throw error;
    if (error instanceof LyricsSearchError) throw error;
    if (timedOut) throw new LyricsSearchError("The lyric database took too long to respond. Try again in a moment.");
    console.warn("Lyric search request failed:", error);
    throw new LyricsSearchError("Couldn't reach the lyric database. Check your internet connection.");
  } finally {
    window.clearTimeout(timer);
    external?.removeEventListener("abort", onAbort);
  }
}

export class LyricsSearchError extends Error {}

export function clearLyricsCache() {
  memoryCache.clear();
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // ignore
  }
}

export async function searchLyrics(query: string, { signal, onResults }: SearchOptions = {}): Promise<LyricsResult[]> {
  const q = query.trim().replace(/\s+/g, " ");
  if (q.length < 2) return [];

  const cached = peekCachedLyrics(q);
  if (cached) {
    onResults?.(cached);
    return cached;
  }

  const requests: Array<Record<string, string>> = [{ track_name: q }, { q }];
  const split = q.match(/^(.+?)\s+(?:-|–|—|by)\s+(.+)$/i);
  if (split) requests.push({ track_name: split[1], artist_name: split[2] });

  const merged = new Map<string, LyricsResult>();
  const failures: Error[] = [];
  let successes = 0;

  await Promise.all(requests.map(async (params) => {
    try {
      const items = await fetchItems(params, signal);
      successes += 1;
      for (const item of items) {
        const result = toResult(item);
        if (result) addResult(merged, result);
      }
      if (merged.size > 0) onResults?.(rank(merged.values(), q));
    } catch (error) {
      if (signal?.aborted) throw error;
      failures.push(error instanceof Error ? error : new Error(String(error)));
    }
  }));

  if (successes === 0) throw failures[0] ?? new LyricsSearchError("The lyric search failed.");

  const ranked = rank(merged.values(), q);
  if (ranked.length > 0) {
    memoryCache.set(cacheKey(q), ranked);
    writeStore(cacheKey(q), ranked);
  }
  return ranked;
}
