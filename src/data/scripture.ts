export type ScriptureTranslation = "niv" | "nirv" | "esv" | "nlt" | "csb" | (string & {});

export interface AvailableScriptureTranslation {
  id: ScriptureTranslation;
  name: string;
  bundled: boolean;
  verseCount?: number;
}

export interface ScriptureVerse {
  reference: string;
  book: string;
  chapter: number;
  verse: number;
  text: string;
  translation: ScriptureTranslation;
}

export interface ScriptureResult extends ScriptureVerse {
  primaryText: string;
  comparisonText: string | null;
}

interface BibleBook {
  englishName: string;
  aliases?: string[];
  chapters: Array<{ number: number; verses: Array<{ number: number; text: string }> }>;
}

interface BibleDocument {
  name: string;
  books: BibleBook[];
}

interface ReferenceRange {
  firstChapter: number;
  firstVerse: number | null;
  lastChapter: number;
  lastVerse: number | null;
}

interface ParsedReference {
  book: string;
  ranges: ReferenceRange[];
}

const bibleCache = new Map<ScriptureTranslation, Promise<BibleDocument>>();
const verseCache = new Map<ScriptureTranslation, ScriptureVerse[]>();
const translationNames: Record<ScriptureTranslation, string> = {
  niv: "New International Version (NIV)",
  nirv: "New International Reader's Version (NIRV)",
  esv: "English Standard Version (ESV)",
  nlt: "New Living Translation (NLT)",
  csb: "Christian Standard Bible (CSB)",
};

const SCRIPTURE_DB = "church-presenter-scripture";
const SCRIPTURE_STORE = "translations";
const bundledTranslations: AvailableScriptureTranslation[] = [
  { id: "niv", name: "NIV", bundled: true },
  { id: "nirv", name: "NIRV", bundled: true },
  { id: "esv", name: "ESV", bundled: true },
  { id: "nlt", name: "NLT", bundled: true },
  { id: "csb", name: "CSB", bundled: true },
];

function scriptureDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(SCRIPTURE_DB, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(SCRIPTURE_STORE, { keyPath: "id" });
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("Could not open the scripture library."));
  });
}

async function customTranslation(id: string): Promise<BibleDocument | undefined> {
  const db = await scriptureDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(SCRIPTURE_STORE, "readonly").objectStore(SCRIPTURE_STORE).get(id);
    request.onsuccess = () => resolve(request.result?.bible as BibleDocument | undefined);
    request.onerror = () => reject(request.error ?? new Error("Could not read the scripture library."));
  });
}

export async function listScriptureTranslations(): Promise<AvailableScriptureTranslation[]> {
  const db = await scriptureDb();
  const custom = await new Promise<Array<{ id: string; name: string; verseCount?: number; removed?: boolean }>>((resolve, reject) => {
    const request = db.transaction(SCRIPTURE_STORE, "readonly").objectStore(SCRIPTURE_STORE).getAll();
    request.onsuccess = () => resolve(request.result as Array<{ id: string; name: string; verseCount?: number; removed?: boolean }>);
    request.onerror = () => reject(request.error ?? new Error("Could not list scripture translations."));
  });
  return [
    ...bundledTranslations,
    ...custom.filter((translation) => !translation.removed).map(({ id, name, verseCount }) => ({ id, name, verseCount, bundled: false })).sort((a, b) => a.name.localeCompare(b.name)),
  ];
}

export async function removeScriptureTranslation(id: ScriptureTranslation): Promise<void> {
  if (bundledTranslations.some((translation) => translation.id === id)) {
    throw new Error("Translations bundled with the app cannot be removed.");
  }
  const db = await scriptureDb();
  await new Promise<void>((resolve, reject) => {
    const store = db.transaction(SCRIPTURE_STORE, "readwrite").objectStore(SCRIPTURE_STORE);
    const request = store.delete(id);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error ?? new Error("Could not remove the translation."));
  });
  bibleCache.delete(id);
  verseCache.delete(id);
}

interface ImportedBible { name: string; books: BibleBook[]; }
const bookCodes = "Gen|Exod|Lev|Num|Deut|Josh|Judg|Ruth|1Sam|2Sam|1Kgs|2Kgs|1Chr|2Chr|Ezra|Neh|Esth|Job|Ps|Prov|Eccl|Song|Isa|Jer|Lam|Ezek|Dan|Hos|Joel|Amos|Obad|Jonah|Mic|Nah|Hab|Zeph|Hag|Zech|Mal|Matt|Mark|Luke|John|Acts|Rom|1Cor|2Cor|Gal|Eph|Phil|Col|1Thess|2Thess|1Tim|2Tim|Titus|Phlm|Heb|Jas|1Pet|2Pet|1John|2John|3John|Jude|Rev".split("|");
const bookNames = "Genesis|Exodus|Leviticus|Numbers|Deuteronomy|Joshua|Judges|Ruth|1 Samuel|2 Samuel|1 Kings|2 Kings|1 Chronicles|2 Chronicles|Ezra|Nehemiah|Esther|Job|Psalms|Proverbs|Ecclesiastes|Song of Solomon|Isaiah|Jeremiah|Lamentations|Ezekiel|Daniel|Hosea|Joel|Amos|Obadiah|Jonah|Micah|Nahum|Habakkuk|Zephaniah|Haggai|Zechariah|Malachi|Matthew|Mark|Luke|John|Acts|Romans|1 Corinthians|2 Corinthians|Galatians|Ephesians|Philippians|Colossians|1 Thessalonians|2 Thessalonians|1 Timothy|2 Timothy|Titus|Philemon|Hebrews|James|1 Peter|2 Peter|1 John|2 John|3 John|Jude|Revelation".split("|");
const bookByCode = new Map(bookCodes.map((code, index) => [code.toLowerCase(), bookNames[index]]));

function parseImportedBibleJson(value: unknown): ImportedBible {
  if (!value || typeof value !== "object" || !Array.isArray((value as BibleDocument).books)) throw new Error("The JSON file must contain a books array in the app Bible format.");
  const bible = value as BibleDocument;
  const books = bible.books.map((book) => ({
    englishName: String(book.englishName || "").trim(),
    chapters: Array.isArray(book.chapters) ? book.chapters.map((chapter) => ({
      number: Number(chapter.number),
      verses: Array.isArray(chapter.verses) ? chapter.verses.map((verse) => ({ number: Number(verse.number), text: String(verse.text || "").trim() })).filter((verse) => verse.number > 0 && verse.text) : [],
    })).filter((chapter) => chapter.number > 0) : [],
  })).filter((book) => book.englishName && book.chapters.length);
  if (!books.length) throw new Error("No Bible books and verses were found in the JSON file.");
  return { name: String(bible.name || "Imported Bible"), books };
}

function parseImportedBibleXml(xmlText: string): ImportedBible {
  const document = new DOMParser().parseFromString(xmlText, "application/xml");
  if (document.querySelector("parsererror")) throw new Error("The XML file could not be parsed.");
  const books = new Map<string, Map<number, Map<number, string>>>();
  const addVerse = (bookCode: string, chapter: number, verse: number, text: string) => {
    const bookNumber = Number(bookCode);
    const name = bookByCode.get(bookCode.toLowerCase()) || bookByCode.get(bookCode.replace(/^0/, "").toLowerCase()) || (Number.isInteger(bookNumber) ? bookNames[bookNumber - 1] : undefined);
    const content = text.replace(/\s+/g, " ").trim();
    if (!name || chapter < 1 || verse < 1 || !content) return;
    if (!books.has(name)) books.set(name, new Map());
    const chapters = books.get(name)!;
    if (!chapters.has(chapter)) chapters.set(chapter, new Map());
    chapters.get(chapter)!.set(verse, content);
  };
  const numberedBooks = [...document.getElementsByTagName("book")].filter((el) => /^\d+$/.test(el.getAttribute("number") || ""));
  const osisVerses = [...document.getElementsByTagNameNS("*", "verse")].filter((el) => el.hasAttribute("osisID"));
  if (numberedBooks.length) {
    for (const bookEl of numberedBooks) {
      const book = bookEl.getAttribute("number") || "";
      for (const chapterEl of [...bookEl.children].filter((el) => el.localName.toLowerCase() === "chapter")) {
        const chapter = Number(chapterEl.getAttribute("number"));
        for (const verseEl of [...chapterEl.children].filter((el) => el.localName.toLowerCase() === "verse")) {
          addVerse(book, chapter, Number(verseEl.getAttribute("number")), verseEl.textContent || "");
        }
      }
    }
  } else if (osisVerses.length) {
    for (const verseEl of osisVerses) {
      const id = (verseEl.getAttribute("osisID") || "").split(/[\s,]/)[0].split(".");
      if (id.length >= 3 && /^\d+$/.test(id[1]) && /^\d+[a-z]?$/i.test(id[2])) addVerse(id[0], Number(id[1]), Number.parseInt(id[2], 10), verseEl.textContent || "");
    }
  } else {
    let book = "", chapter = 0, verse = 0, text = "";
    const flush = () => { if (book && chapter && verse) addVerse(book, chapter, verse, text); text = ""; };
    const walker = document.createTreeWalker(document.documentElement, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
    let node: Node | null;
    while ((node = walker.nextNode())) {
      if (node.nodeType === Node.TEXT_NODE) { if (verse) text += ` ${node.textContent || ""}`; continue; }
      const el = node as Element, local = el.localName.toLowerCase();
      if (local === "book") book = el.getAttribute("code") || el.getAttribute("osisID") || "";
      else if (local === "chapter") {
        flush();
        const n = el.getAttribute("number") || el.getAttribute("osisID")?.split(".").at(-1) || "";
        if (/^\d+$/.test(n)) chapter = Number(n);
      } else if (local === "verse") {
        flush();
        const n = el.getAttribute("number") || el.getAttribute("osisID")?.split(".").at(-1) || el.getAttribute("sID")?.split(".").at(-1) || "";
        if (/^\d+[a-z]?$/i.test(n)) verse = Number.parseInt(n, 10);
        else if (!el.hasAttribute("eID")) verse = 0;
      }
    }
    flush();
  }
  const result: BibleBook[] = [...books].map(([englishName, chapters]) => ({
    englishName,
    chapters: [...chapters].sort(([a], [b]) => a - b).map(([number, verses]) => ({ number, verses: [...verses].sort(([a], [b]) => a - b).map(([number, text]) => ({ number, text })) })),
  }));
  if (!result.length) throw new Error("No Bible verses were found. Import OSIS or USX XML, or app-format Bible JSON.");
  return { name: document.documentElement.getAttribute("translation") || document.documentElement.getAttribute("name") || "Imported Bible", books: result };
}

export async function importScriptureTranslation(file: File): Promise<AvailableScriptureTranslation> {
  const text = await file.text();
  const bible = file.name.toLowerCase().endsWith(".json") ? parseImportedBibleJson(JSON.parse(text)) : parseImportedBibleXml(text);
  const stem = file.name.replace(/\.[^.]+$/, "").replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "translation";
  let id = `custom-${stem}`;
  const existing = await listScriptureTranslations();
  let suffix = 2;
  while (existing.some((translation) => translation.id === id)) id = `custom-${stem}-${suffix++}`;
  const name = bible.name && bible.name !== "Imported Bible" ? bible.name : file.name.replace(/\.[^.]+$/, "");
  const verseCount = bible.books.reduce((sum, book) => sum + book.chapters.reduce((chapterSum, chapter) => chapterSum + chapter.verses.length, 0), 0);
  const db = await scriptureDb();
  await new Promise<void>((resolve, reject) => {
    const request = db.transaction(SCRIPTURE_STORE, "readwrite").objectStore(SCRIPTURE_STORE).put({ id, name, verseCount, bible: { name, books: bible.books } });
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error ?? new Error("Could not save the imported translation."));
  });
  return { id, name, verseCount, bundled: false };
}

const bookAliases: Record<string, string[]> = {
  Genesis: ["gen", "ge", "gn"], Exodus: ["ex", "exod"], Leviticus: ["lev", "lv"],
  Numbers: ["num", "nu", "nm"], Deuteronomy: ["deut", "dt"], Joshua: ["josh", "jos"],
  Judges: ["judg", "jdg"], Ruth: ["ru"], "1 Samuel": ["1 sam", "1 samuel", "i samuel", "1sa", "1 sm"],
  "2 Samuel": ["2 sam", "2 samuel", "ii samuel", "2sa", "2 sm"],
  "1 Kings": ["1 kgs", "1 kings", "i kings", "1ki"], "2 Kings": ["2 kgs", "2 kings", "ii kings", "2ki"],
  "1 Chronicles": ["1 chron", "1 chronicles", "i chronicles", "1chr", "1 ch"],
  "2 Chronicles": ["2 chron", "2 chronicles", "ii chronicles", "2chr", "2 ch"],
  Ezra: ["ezr"], Nehemiah: ["neh", "ne"], Esther: ["est", "esth"], Job: ["jb"],
  Psalms: ["ps", "psalm", "psa", "pss"], Proverbs: ["prov", "pr", "prv"],
  Ecclesiastes: ["eccl", "ecc", "qoheleth"], "Song of Solomon": ["song", "song of songs", "sos", "canticles", "cant"],
  Isaiah: ["isa", "is"], Jeremiah: ["jer", "je"], Lamentations: ["lam", "la"],
  Ezekiel: ["ezek", "eze"], Daniel: ["dan", "da"], Hosea: ["hos", "ho"], Joel: ["joe", "jl"],
  Amos: ["am"], Obadiah: ["obad", "ob"], Jonah: ["jon", "jnh"], Micah: ["mic", "mi"],
  Nahum: ["nah", "na"], Habakkuk: ["hab", "hb"], Zephaniah: ["zeph", "zep"],
  Haggai: ["hag", "hg"], Zechariah: ["zech", "zec"], Malachi: ["mal", "ml"],
  Matthew: ["matt", "mat", "mt"], Mark: ["mk", "mrk"], Luke: ["lk", "luk"], John: ["jn", "jhn"],
  Acts: ["act"], Romans: ["rom", "ro"], "1 Corinthians": ["1 cor", "1 corinthians", "i corinthians", "1co", "1 cor"],
  "2 Corinthians": ["2 cor", "2 corinthians", "ii corinthians", "2co"],
  Galatians: ["gal", "ga"], Ephesians: ["eph", "ep"], Philippians: ["phil", "php", "pp"],
  Colossians: ["col", "co"], "1 Thessalonians": ["1 thes", "1 thessalonians", "i thessalonians", "1th"],
  "2 Thessalonians": ["2 thes", "2 thessalonians", "ii thessalonians", "2th"],
  "1 Timothy": ["1 tim", "1 timothy", "i timothy", "1ti"], "2 Timothy": ["2 tim", "2 timothy", "ii timothy", "2ti"],
  Titus: ["tit", "ti"], Philemon: ["philem", "phm"], Hebrews: ["heb", "he"],
  James: ["jas", "jam", "jm"], "1 Peter": ["1 pet", "1 peter", "i peter", "1pe"],
  "2 Peter": ["2 pet", "2 peter", "ii peter", "2pe"], "1 John": ["1 jn", "1 john", "i john", "1jo", "1 jhn"],
  "2 John": ["2 jn", "2 john", "ii john", "2jo", "2 jhn"], "3 John": ["3 jn", "3 john", "iii john", "3jo", "3 jhn"],
  Jude: ["jud"], Revelation: ["rev", "re", "apocalypse"],
};

export function getTranslationName(translation: ScriptureTranslation): string {
  return translationNames[translation] ?? bundledTranslations.find((item) => item.id === translation)?.name ?? translation;
}

async function getBible(translation: ScriptureTranslation): Promise<BibleDocument> {
  let bible = bibleCache.get(translation);
  if (!bible) {
    bible = (bundledTranslations.some((item) => item.id === translation)
      ? listScriptureTranslations().then((available) => {
        if (!available.some((item) => item.id === translation)) throw new Error(`The ${translation.toUpperCase()} translation has been removed.`);
        return fetch(`${import.meta.env.BASE_URL}scripture/${translation.toUpperCase()}.xml`);
      }).then(async (response) => {
        if (!response.ok) throw new Error(`Could not load the bundled ${translation.toUpperCase()} scripture library.`);
        return parseImportedBibleXml(await response.text());
      })
      : customTranslation(translation).then((result) => {
        if (!result) throw new Error(`The ${translation} translation is not installed.`);
        return result;
      }));
    bibleCache.set(translation, bible);
    bible.catch(() => bibleCache.delete(translation));
  }
  return bible;
}

function normalize(value: string): string {
  return value.toLocaleLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim().replace(/\s+/g, " ");
}

function parseReference(query: string): ParsedReference | null {
  const input = query.trim().replace(/[–—]/g, "-").replace(/\s+/g, " ");
  const aliases = Object.entries(bookAliases).flatMap(([book, names]) => [book, ...names].map((alias) => ({ book, alias })))
    .sort((a, b) => b.alias.length - a.alias.length);
  let parsed: ParsedReference | null = null;
  let locator = "";
  for (const candidate of aliases) {
    const prefix = candidate.alias.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s+");
    const match = input.match(new RegExp(`^${prefix}(?=$|\\s)(.*)$`, "i"));
    if (match) {
      parsed = { book: candidate.book, ranges: [] };
      locator = match[1].trim();
      break;
    }
  }
  if (!parsed) return null;
  if (!locator) return parsed;

  // People commonly enter "jn 3 16" while presenting; treat the final number as a verse.
  locator = locator.replace(/^(\d+)\s+(\d+)$/, "$1:$2");
  let contextChapter: number | null = null;
  for (const rawSegment of locator.split(",")) {
    const segment = rawSegment.trim();
    if (!segment) return null;
    const pieces = segment.split(/\s*-\s*/);
    if (pieces.length > 2) return null;
    const endpoint = (part: string, chapterContext: number | null): { chapter: number; verse: number | null } | null => {
      const full = part.match(/^(\d+)(?::(\d+))?$/);
      const verseOnly = part.match(/^(\d+)$/);
      if (!full) return null;
      if (part.includes(":")) return { chapter: Number(full[1]), verse: Number(full[2]) };
      if (chapterContext !== null) return { chapter: chapterContext, verse: Number(verseOnly?.[1]) };
      return { chapter: Number(full[1]), verse: null };
    };
    const start = endpoint(pieces[0], contextChapter);
    if (!start || start.chapter < 1) return null;
    let end = start;
    if (pieces.length === 2) {
      const parsedEnd = endpoint(pieces[1], start.chapter);
      if (!parsedEnd) return null;
      end = parsedEnd;
      // A bare end after a chapter-only start means the last chapter in a chapter range.
      if (start.verse === null && !pieces[1].includes(":")) end = { chapter: Number(pieces[1]), verse: null };
    }
    if (end.chapter < start.chapter || (start.chapter === end.chapter && start.verse !== null && end.verse !== null && end.verse < start.verse)) return null;
    parsed.ranges.push({ firstChapter: start.chapter, firstVerse: start.verse, lastChapter: end.chapter, lastVerse: end.verse });
    contextChapter = end.chapter;
  }
  return parsed;
}

function findBook(bible: BibleDocument, query: string): BibleBook | undefined {
  const normalized = normalize(query);
  return bible.books.find(({ englishName }) => normalize(englishName) === normalized)
    ?? bible.books.find(({ englishName }) => normalize(englishName).startsWith(normalized));
}

function allVerses(bible: BibleDocument, translation: ScriptureTranslation): ScriptureVerse[] {
  const cached = verseCache.get(translation);
  if (cached) return cached;
  const verses = bible.books.flatMap((book) => book.chapters.flatMap((chapter) => chapter.verses.map((verse) => ({
    reference: `${book.englishName} ${chapter.number}:${verse.number}`,
    book: book.englishName,
    chapter: chapter.number,
    verse: verse.number,
    text: verse.text,
    translation,
  }))));
  verseCache.set(translation, verses);
  return verses;
}

function verseInRange(verse: ScriptureVerse, range: ReferenceRange): boolean {
  const key = verse.chapter * 1000 + verse.verse;
  const first = range.firstChapter * 1000 + (range.firstVerse ?? 0);
  const last = range.lastChapter * 1000 + (range.lastVerse ?? 999);
  return key >= first && key <= last;
}

function searchBible(bible: BibleDocument, translation: ScriptureTranslation, query: string, limit: number): ScriptureVerse[] {
  const reference = parseReference(query);
  if (reference) {
    const book = findBook(bible, reference.book);
    if (!book) return [];
    const bookVerses = allVerses(bible, translation).filter((verse) => normalize(verse.book) === normalize(book.englishName));
    if (!reference.ranges.length) return bookVerses.slice(0, limit);
    return bookVerses.filter((verse) => reference.ranges.some((range) => verseInRange(verse, range))).slice(0, 200);
  }

  const needle = normalize(query);
  if (needle.length < 2) return [];
  const words = needle.split(" ").filter(Boolean);
  const matches: ScriptureVerse[] = [];
  for (const verse of allVerses(bible, translation)) {
    const text = normalize(verse.text);
    if (text.includes(needle) || (words.length > 1 && words.every((word) => text.includes(word)))) {
      matches.push(verse);
      if (matches.length >= limit) return matches;
    }
  }
  return matches;
}

export function isScriptureReference(query: string): boolean {
  return parseReference(query) !== null;
}

export async function searchScripture(
  query: string,
  primary: ScriptureTranslation,
  comparison: ScriptureTranslation | null = null,
  limit = 12,
): Promise<ScriptureResult[]> {
  const trimmed = query.trim();
  if (!trimmed) return [];
  const [primaryBible, comparisonBible] = await Promise.all([
    getBible(primary),
    comparison && comparison !== primary ? getBible(comparison) : Promise.resolve(null),
  ]);
  const primaryResults = searchBible(primaryBible, primary, trimmed, limit);
  const comparisonResults = comparisonBible ? searchBible(comparisonBible, comparison!, trimmed, limit) : [];
  const byReference = new Map<string, ScriptureResult>();
  for (const verse of primaryResults) byReference.set(normalize(verse.reference), { ...verse, primaryText: verse.text, comparisonText: null });
  for (const verse of comparisonResults) {
    const key = normalize(verse.reference);
    const existing = byReference.get(key);
    if (existing) existing.comparisonText = verse.text;
    else byReference.set(key, { ...verse, primaryText: "", comparisonText: verse.text });
  }
  return [...byReference.values()].slice(0, parseReference(trimmed) ? 200 : limit);
}
