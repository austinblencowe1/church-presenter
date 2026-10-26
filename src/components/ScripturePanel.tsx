import { useEffect, useState } from "react";
import { getTranslationName, listScriptureTranslations, searchScripture, type AvailableScriptureTranslation, type ScriptureResult, type ScriptureTranslation } from "../data/scripture";

export interface SelectedScriptureVerse {
  reference: string;
  text: string;
  translation: string;
}

interface ScripturePanelProps {
  readonly onAddToService: (verses: SelectedScriptureVerse[]) => void;
  readonly onGoLive: (verse: SelectedScriptureVerse) => void;
}

function scriptureText(result: ScriptureResult): string {
  return result.primaryText || result.comparisonText || "";
}

function scriptureVersion(result: ScriptureResult, primary: ScriptureTranslation, comparison: ScriptureTranslation | null, translations: AvailableScriptureTranslation[]): string {
  const id = result.primaryText ? primary : (comparison ?? primary);
  return translations.find((translation) => translation.id === id)?.name ?? id.toUpperCase();
}

export function ScripturePanel({ onAddToService, onGoLive }: ScripturePanelProps) {
  const [query, setQuery] = useState("");
  const [translations, setTranslations] = useState<AvailableScriptureTranslation[]>([]);
  const [primary, setPrimary] = useState<ScriptureTranslation>("niv");
  const [comparison, setComparison] = useState<ScriptureTranslation | "none">("nirv");
  const [results, setResults] = useState<ScriptureResult[]>([]);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const comparisonVersion = comparison === "none" ? null : comparison;

  async function refreshResults(nextPrimary: ScriptureTranslation, nextComparison: ScriptureTranslation | null) {
    if (!query.trim()) return;
    setBusy(true);
    setError("");
    try {
      // Keep selected references visible even if a phrase ranks differently in
      // the new translation, and refresh their text in the selected version.
      const translatedSelected = await Promise.all(results
        .filter((result) => selected.has(result.reference))
        .map(async (result) => (await searchScripture(result.reference, nextPrimary, nextComparison, 1))[0] ?? result));
      const refreshed = await searchScripture(query, nextPrimary, nextComparison);
      const byReference = new Map(refreshed.map((result) => [result.reference, result]));
      for (const result of translatedSelected) byReference.set(result.reference, result);
      setResults([...byReference.values()]);
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setBusy(false);
    }
  }

  useEffect(() => {
    void listScriptureTranslations().then((available) => {
      const ids = available.map(({ id }) => id);
      const nextPrimary = ids.includes(primary) ? primary : ids[0] ?? "niv";
      setTranslations(available);
      setPrimary(nextPrimary);
      setComparison((current) => current === "none" || (ids.includes(current) && current !== nextPrimary)
        ? current
        : ids.find((id) => id !== nextPrimary) ?? "none");
    }).catch((reason) => setError(reason instanceof Error ? reason.message : String(reason)));
  }, []);

  async function search(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!query.trim()) return;
    setBusy(true);
    setError("");
    try {
      const found = await searchScripture(query, primary, comparisonVersion);
      setResults(found);
      setSelected(new Set(found.filter((result) => result.primaryText || result.comparisonText).map((result) => result.reference)));
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
      setResults([]);
    } finally {
      setBusy(false);
    }
  }

  function selectedVerses(): SelectedScriptureVerse[] {
    return results.filter((result) => selected.has(result.reference)).map((result) => ({
      reference: result.reference,
      text: scriptureText(result),
      translation: scriptureVersion(result, primary, comparisonVersion, translations),
    }));
  }

  return <section className="scripture-panel" aria-label="Scripture search">
    <header className="scripture-panel__header">
      <div><span className="eyebrow">OFFLINE SCRIPTURE</span><h2>Bible & Scripture</h2><p>Search by passage or phrase, compare translations, and send verses to a screen.</p></div>
      <span className="scripture-panel__offline">LOCAL LIBRARY</span>
    </header>
    <form className="scripture-search" onSubmit={(event) => void search(event)}>
      <input value={query} onChange={(event) => setQuery(event.currentTarget.value)} placeholder="John 3:16–18 or search a phrase" aria-label="Search scripture by passage or words" />
      <label>PRIMARY
        <select value={primary} onChange={(event) => {
          const nextPrimary = event.currentTarget.value as ScriptureTranslation;
          const nextComparison = comparison === nextPrimary
            ? translations.find((version) => version.id !== nextPrimary)?.id ?? "none"
            : comparison;
          setPrimary(nextPrimary);
          if (comparison !== nextComparison) setComparison(nextComparison);
          void refreshResults(nextPrimary, nextComparison === "none" ? null : nextComparison);
        }}>
          {translations.map((version) => <option key={version.id} value={version.id}>{version.name}</option>)}
        </select>
      </label>
      <label>COMPARE
        <select value={comparison} onChange={(event) => {
          const nextComparison = event.currentTarget.value as ScriptureTranslation | "none";
          setComparison(nextComparison);
          void refreshResults(primary, nextComparison === "none" ? null : nextComparison);
        }}>
          <option value="none">Off</option>
          {translations.filter((version) => version.id !== primary).map((version) => <option key={version.id} value={version.id}>{version.name}</option>)}
        </select>
      </label>
      <button type="submit" disabled={busy || !query.trim()}>{busy ? "SEARCHING..." : "SEARCH"}</button>
    </form>
    <div className="scripture-panel__toolbar">
      <span>{results.length ? `${results.length} results${results.length === 50 ? " · showing first 50" : ""}` : "Use a book and verse (John 3:16) or enter a keyword"}</span>
      {results.length > 0 && <div>
        <button onClick={() => setSelected(new Set(results.map((result) => result.reference)))}>SELECT ALL</button>
        <button onClick={() => onAddToService(selectedVerses())} disabled={selected.size === 0}>ADD SELECTED TO SERVICE ({selected.size})</button>
      </div>}
    </div>
    {error && <p className="scripture-panel__error" role="alert">{error}</p>}
    <div className="scripture-results">
      {results.map((result) => <article key={result.reference} className="scripture-result">
        <label className="scripture-result__select">
          <input type="checkbox" checked={selected.has(result.reference)} onChange={(event) => setSelected((current) => {
            const next = new Set(current);
            if (event.currentTarget.checked) next.add(result.reference);
            else next.delete(result.reference);
            return next;
          })} />
          <strong>{result.reference}</strong>
        </label>
        <div className="scripture-result__translations">
          {result.primaryText && <p><small>{getTranslationName(primary)}</small>{result.primaryText}</p>}
          {result.comparisonText && <p><small>{getTranslationName(comparisonVersion ?? primary)}</small>{result.comparisonText}</p>}
        </div>
        <button className="scripture-result__live" onClick={() => onGoLive({ reference: result.reference, text: scriptureText(result), translation: scriptureVersion(result, primary, comparisonVersion, translations) })}>GO LIVE</button>
      </article>)}
      {results.length === 0 && !busy && !error && <div className="scripture-results__empty">Search the installed Bible translations for a passage or phrase.</div>}
    </div>
  </section>;
}
