const SYSTEM_PROMPT = `You help a church find song lyrics. Search the web for the requested song and only report what you find on real pages. Return JSON matching the schema. Include title, authors and year only if a source states them. For the snippet, return only the first 4 lines of the song. Never write lyrics from memory and never guess. If you are unsure that a result matches, lower its confidence or leave it out.`;

const searchResultSchema = {
  type: 'object',
  properties: { results: {
    type: 'array',
    maxItems: 8,
    items: {
    type: 'object',
    properties: {
      id: { type: 'string' },
      title: { type: 'string' },
      authors: { type: 'string' },
      year: { type: 'integer' },
      snippet: { type: 'string' },
      sourceUrl: { type: 'string' },
      confidence: { type: 'number' },
    },
    required: ['id', 'title', 'snippet', 'sourceUrl', 'confidence'],
    additionalProperties: false,
    },
  } },
  required: ['results'],
  additionalProperties: false,
};

const lyricsSchema = {
  type: 'object',
  properties: { lyrics: { type: 'string' } },
  required: ['lyrics'],
  additionalProperties: false,
};

const searchCache = new Map();
const sourceById = new Map();
let settings = {};

export function configureSongSearch(nextSettings) {
  settings = { ...nextSettings };
  searchCache.clear();
  sourceById.clear();
}

export function clearSongSearchCache() {
  searchCache.clear();
  sourceById.clear();
}

function fail(message, code = 'error') {
  const error = new Error(message);
  error.code = code;
  throw error;
}

function requireOnline() {
  if (typeof navigator !== 'undefined' && navigator.onLine === false) {
    fail('You appear to be offline. Reconnect and try again.', 'offline');
  }
}

async function postJson(url, body, signal, headers = {}) {
  let response;
  try {
    response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...headers },
      body: JSON.stringify(body),
      signal,
    });
  } catch (error) {
    if (error.name === 'AbortError') throw error;
    if (typeof navigator !== 'undefined' && navigator.onLine === false) {
      fail('You appear to be offline. Reconnect and try again.', 'offline');
    }
    fail('Could not connect to the search or AI endpoint. Check the endpoint and try again.', 'error');
  }
  if (!response.ok) {
    const code = response.status === 401 || response.status === 403 ? 'config' : 'error';
    fail(`Request failed (${response.status}). Check your provider settings and try again.`, code);
  }
  try {
    return await response.json();
  } catch {
    fail('The endpoint returned invalid JSON. Check that the configured endpoint is correct.');
  }
}

function providerReady() {
  if (!settings.searchApiKey?.trim()) fail('Add your Tavily API key in Settings before searching.', 'config');
  if (settings.provider === 'ollama') {
    if (!settings.localModel?.trim()) fail('Add the name of an installed Ollama model in Settings.', 'config');
    if (!settings.localEndpoint?.trim()) fail('Add your Ollama endpoint in Settings.', 'config');
  } else if (!settings.hostedEndpoint?.trim() || !settings.hostedModel?.trim() || !settings.hostedApiKey?.trim()) {
    fail('Add the hosted endpoint, model name, and API key in Settings.', 'config');
  }
}

async function searchWeb(query, signal) {
  const result = await postJson(settings.searchEndpoint, {
    query: `${query} song lyrics authors release year`,
    topic: 'general',
    max_results: 10,
    include_answer: false,
    include_raw_content: false,
  }, signal, { Authorization: `Bearer ${settings.searchApiKey}` });
  const list = Array.isArray(result.results) ? result.results : [];
  return list.filter(item => {
    try { return ['http:', 'https:'].includes(new URL(item.url).protocol); }
    catch { return false; }
  }).slice(0, 10).map(item => ({
    id: crypto.randomUUID(),
    url: item.url,
    title: String(item.title || '').slice(0, 240),
    content: String(item.content || '').slice(0, 2400),
  }));
}

async function modelJson(messages, schema, signal) {
  requireOnline();
  if (settings.provider === 'ollama') {
    const root = settings.localEndpoint.trim().replace(/\/+$/, '').replace(/\/api\/chat$/, '');
    const response = await postJson(`${root}/api/chat`, {
      model: settings.localModel.trim(),
      messages,
      stream: false,
      format: schema,
      options: { temperature: 0.1 },
    }, signal);
    const text = response.message?.content;
    if (typeof text !== 'string') fail('The local model returned no response. Check its model name and try again.');
    try { return JSON.parse(text); } catch { fail('The model response was not valid JSON. Retry the search.'); }
  }
  const response = await postJson(settings.hostedEndpoint.trim(), {
    model: settings.hostedModel.trim(),
    messages,
    temperature: 0.1,
    response_format: { type: 'json_object' },
  }, signal, { Authorization: `Bearer ${settings.hostedApiKey}` });
  const text = response.choices?.[0]?.message?.content;
  if (typeof text !== 'string') fail('The hosted model returned no response. Check the endpoint and try again.');
  try { return JSON.parse(text); } catch { fail('The model response was not valid JSON. Retry the search.'); }
}

function validateSearchResults(value, sources) {
  if (!Array.isArray(value)) return [];
  const sourceMap = new Map(sources.map(source => [source.id, source]));
  const seen = new Set();
  const valid = [];
  for (const item of value) {
    if (!item || typeof item !== 'object' || typeof item.id !== 'string' ||
        typeof item.title !== 'string' || !item.title.trim() ||
        typeof item.snippet !== 'string' || typeof item.sourceUrl !== 'string' ||
        typeof item.confidence !== 'number' || !Number.isFinite(item.confidence) ||
        item.confidence < 0 || item.confidence > 1) continue;
    const source = sourceMap.get(item.id);
    if (!source || item.sourceUrl !== source.url) continue;
    const key = item.title.trim().toLocaleLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const snippetLines = item.snippet.split(/\r?\n/).map(line => line.trim()).filter(Boolean).slice(0, 4);
    const snippet = snippetLines.join('\n').slice(0, 420);
    const excerpt = normalizeForEvidence(source.content);
    const snippetIsGrounded = snippetLines.length > 0 && snippetLines.every(line => {
      const normalized = normalizeForEvidence(line);
      return normalized.length > 3 && excerpt.includes(normalized);
    });
    if (!snippetIsGrounded || item.confidence < 0.2) continue;
    const authors = Array.isArray(item.authors) ? item.authors.filter(x => typeof x === 'string').join(', ') : item.authors;
    const result = {
      id: item.id,
      title: item.title.trim().slice(0, 200),
      snippet,
      sourceUrl: source.url,
      confidence: item.confidence,
    };
    if (typeof authors === 'string' && authors.trim() && !/^unknown$/i.test(authors.trim())) result.authors = authors.trim().slice(0, 240);
    const year = typeof item.year === 'number' ? item.year : Number(item.year);
    if (Number.isInteger(year) && year >= 1000 && year <= new Date().getFullYear()) result.year = year;
    valid.push(result);
    if (valid.length === 8) break;
  }
  return valid;
}

export async function searchSongs(query, { signal } = {}) {
  requireOnline();
  providerReady();
  const key = query.trim().toLocaleLowerCase();
  if (searchCache.has(key)) return searchCache.get(key).map(item => ({ ...item }));
  const sources = await searchWeb(query, signal);
  if (!sources.length) {
    searchCache.set(key, []);
    return [];
  }
  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    { role: 'system', content: 'Return JSON only as an object with a results array of up to 8 objects matching the supplied schema. Use only source IDs and URLs below. Each snippet must contain only the first 4 lyric lines found in that source. Do not include full lyrics. If the page does not provide a reliable match or lyric snippet, omit it.' },
    { role: 'user', content: JSON.stringify({ query, sources }) },
  ];
  const raw = await modelJson(messages, searchResultSchema, signal);
  const results = validateSearchResults(raw?.results, sources);
  sources.forEach(source => sourceById.set(source.id, source));
  searchCache.set(key, results);
  return results.map(item => ({ ...item }));
}

function normalizeForEvidence(text) {
  return text.normalize('NFKD').toLocaleLowerCase().replace(/[^\p{L}\p{N}]+/gu, '');
}

function verifyLyrics(lyrics, pageText) {
  const evidence = normalizeForEvidence(pageText);
  const lines = lyrics.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  if (!lines.length) fail('No lyrics were found on the source page.', 'source');
  const section = /^(verse|chorus|bridge|pre-chorus|tag|intro|outro)\b/i;
  const lyricLines = lines.filter(line => !section.test(line));
  if (!lyricLines.length || lyricLines.some(line => {
    const normalized = normalizeForEvidence(line);
    return normalized.length > 3 && !evidence.includes(normalized);
  })) fail('The source did not contain all returned lyric lines, so nothing was added. Try another result or source.', 'source');
  return lines.join('\n');
}

export async function getLyrics(id, { signal } = {}) {
  requireOnline();
  providerReady();
  const source = sourceById.get(id);
  if (!source) fail('That search result has expired. Search again and retry.', 'source');
  const extracted = await postJson(settings.extractEndpoint, {
    urls: [source.url],
    extract_depth: 'advanced',
    format: 'text',
  }, signal, { Authorization: `Bearer ${settings.searchApiKey}` });
  const pageText = (extracted.results || []).map(item => String(item.raw_content || item.content || '')).join('\n\n').slice(0, 50000);
  if (!pageText.trim()) fail('Could not read lyrics from that source page. Try another result.', 'source');
  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    { role: 'system', content: 'For this request return JSON only as an object with one string field named lyrics. Extract the full lyrics and section labels exactly as they appear on this source page, with a blank line between sections. Keep labels such as Verse 1, Chorus, Bridge, Pre-chorus, Tag, Intro and Outro when present. Never complete missing lines or add lyrics that are not on the page. If the page does not contain the lyrics, return an empty string.' },
    { role: 'user', content: JSON.stringify({ title: source.title, sourceUrl: source.url, pageText }) },
  ];
  const raw = await modelJson(messages, lyricsSchema, signal);
  if (!raw || typeof raw.lyrics !== 'string') fail('The model returned malformed lyrics. Nothing was added.', 'source');
  return verifyLyrics(raw.lyrics, pageText);
}
