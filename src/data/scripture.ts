export type ScriptureTranslation = "web" | "kjv";

export interface ScriptureVerse {
  reference: string;
  book: string;
  chapter: number;
  verse: number;
  text: string;
}

export interface ScriptureResult extends ScriptureVerse {
  primaryText: string;
  comparisonText: string | null;
}

interface BibleBook {
  englishName: string;
  chapters: Array<{ number: number; verses: Array<{ number: number; text: string }> }>;
}

interface BibleDocument {
  name: string;
  books: BibleBook[];
}

const bibleCache = new Map<ScriptureTranslation, Promise<BibleDocument>>();

const translationNames: Record<ScriptureTranslation, string> = {
  web: "World English Bible",
  kjv: "King James Version",
};

export function getTranslationName(translation: ScriptureTranslation): string {
  return translationNames[translation];
}

async function getBible(translation: ScriptureTranslation): Promise<BibleDocument> {
  let bible = bibleCache.get(translation);
  if (!bible) {
    bible = fetch(`${import.meta.env.BASE_URL}scripture/${translation}.json`)
      .then(async (response) => {
        if (!response.ok) throw new Error(`Could not load the ${translation.toUpperCase()} scripture library.`);
        return await response.json() as BibleDocument;
      });
    bibleCache.set(translation, bible);
    bible.catch(() => bibleCache.delete(translation));
  }
  return bible;
}

function parseReference(query: string): { book: string; chapter: number; firstVerse: number | null; lastVerse: number | null } | null {
  const match = query.trim().match(/^(.+?)\s+(\d+)(?::(\d+)(?:\s*[-–]\s*(\d+))?)?$/);
  if (!match) return null;
  return {
    book: match[1].trim().toLocaleLowerCase(),
    chapter: Number(match[2]),
    firstVerse: match[3] ? Number(match[3]) : null,
    lastVerse: match[4] ? Number(match[4]) : match[3] ? Number(match[3]) : null,
  };
}

function normalize(value: string): string {
  return value.toLocaleLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim();
}

function findBook(bible: BibleDocument, query: string): BibleBook | undefined {
  const normalized = normalize(query);
  return bible.books.find(({ englishName }) => normalize(englishName) === normalized)
    ?? bible.books.find(({ englishName }) => normalize(englishName).startsWith(normalized));
}

function versesForBook(book: BibleBook, chapterNumber?: number): ScriptureVerse[] {
  return book.chapters
    .filter((chapter) => chapterNumber === undefined || chapter.number === chapterNumber)
    .flatMap((chapter) => chapter.verses.map((verse) => ({
      reference: `${book.englishName} ${chapter.number}:${verse.number}`,
      book: book.englishName,
      chapter: chapter.number,
      verse: verse.number,
      text: verse.text,
    })));
}

function searchBible(bible: BibleDocument, query: string, limit: number): ScriptureVerse[] {
  const reference = parseReference(query);
  if (reference) {
    const book = findBook(bible, reference.book);
    if (!book) return [];
    return versesForBook(book, reference.chapter).filter(({ verse }) => {
      if (reference.firstVerse === null) return true;
      return verse >= reference.firstVerse && verse <= (reference.lastVerse ?? reference.firstVerse);
    }).slice(0, 150);
  }

  const needle = normalize(query);
  if (needle.length < 2) return [];
  const matches: ScriptureVerse[] = [];
  for (const book of bible.books) {
    for (const verse of versesForBook(book)) {
      if (normalize(verse.text).includes(needle)) {
        matches.push(verse);
        if (matches.length >= limit) return matches;
      }
    }
  }
  return matches;
}

export async function searchScripture(
  query: string,
  primary: ScriptureTranslation,
  comparison: ScriptureTranslation | null,
  limit = 50,
): Promise<ScriptureResult[]> {
  const trimmed = query.trim();
  if (!trimmed) return [];
  const [primaryBible, comparisonBible] = await Promise.all([
    getBible(primary),
    comparison && comparison !== primary ? getBible(comparison) : Promise.resolve(null),
  ]);
  const primaryResults = searchBible(primaryBible, trimmed, limit);
  const comparisonResults = comparisonBible ? searchBible(comparisonBible, trimmed, limit) : [];
  const byReference = new Map<string, ScriptureResult>();
  for (const verse of primaryResults) {
    byReference.set(normalize(verse.reference), { ...verse, primaryText: verse.text, comparisonText: null });
  }
  for (const verse of comparisonResults) {
    const key = normalize(verse.reference);
    const existing = byReference.get(key);
    if (existing) existing.comparisonText = verse.text;
    else byReference.set(key, { ...verse, primaryText: "", comparisonText: verse.text });
  }
  return [...byReference.values()].slice(0, parseReference(trimmed) ? 200 : limit);
}
