// Compatibility wrapper. index.html imports { configureSongSearch, searchSongs, getLyrics } from this file.
// The real implementation (debounce-friendly, parallel queries, caching, timeouts) lives in lyricsLookup.ts.
import { searchLyrics, clearLyricsCache } from "./lyricsLookup.ts";

const lyricsById = new Map();

export function configureSongSearch() {
  // Old Ollama / OpenAI / Tavily settings are no longer used - the free LRCLIB lookup is the only provider.
  lyricsById.clear();
  clearLyricsCache();
}

export function clearSongSearchCache() {
  lyricsById.clear();
  clearLyricsCache();
}

export async function searchSongs(query, { signal, onResults } = {}) {
  const toLegacy = (results) => results.map((result) => {
    lyricsById.set(result.id, result.lyrics);
    return {
      id: result.id,
      title: result.title,
      authors: result.artist || "Various Artists",
      snippet: result.snippet.replace(/ \/ /g, "\n"),
      sourceUrl: "https://lrclib.net",
      confidence: 0.95,
    };
  });
  const results = await searchLyrics(query, {
    signal,
    onResults: onResults ? (partial) => onResults(toLegacy(partial)) : undefined,
  });
  return toLegacy(results);
}

export async function getLyrics(id) {
  const lyrics = lyricsById.get(id);
  if (lyrics) return lyrics;
  throw new Error("Lyrics not found for this song. Please try searching again.");
}
