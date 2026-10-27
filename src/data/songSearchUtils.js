/** Normalize song search text consistently for library filtering and online result ranking. */
export function normalizeSongSearchText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase()
    .replace(/[’'`]/g, "")
    .replace(/[-‐‑‒–—]/g, " ")
    .replace(/[^\p{L}\p{N}\s]/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function songSearchMatches(fields, query) {
  const normalizedQuery = normalizeSongSearchText(query);
  if (!normalizedQuery) return true;
  const compactQuery = normalizedQuery.replace(/\s+/g, "");
  return fields
    .map(normalizeSongSearchText)
    .filter(Boolean)
    .some((field) => field.includes(normalizedQuery) || field.replace(/\s+/g, "").includes(compactQuery));
}
