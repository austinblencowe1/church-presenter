import { useState } from "react";
import { getTranslationName, searchScripture, type ScriptureResult, type ScriptureTranslation } from "../data/scripture";

export interface SelectedScriptureVerse {
  reference: string;
  text: string;
  translation: string;
}

interface ScripturePanelProps {
  readonly onAddToService: (verses: SelectedScriptureVerse[]) => void;
  readonly onGoLive: (verse: SelectedScriptureVerse) => void;
}

const translations: ScriptureTranslation[] = ["web", "kjv"];

function scriptureText(result: ScriptureResult): string {
  return result.primaryText || result.comparisonText || "";
}

function scriptureVersion(result: ScriptureResult, primary: ScriptureTranslation, comparison: ScriptureTranslation | null): string {
  return result.primaryText ? primary.toUpperCase() : (comparison ?? primary).toUpperCase();
}

export function ScripturePanel({ onAddToService, onGoLive }: ScripturePanelProps) {
  const [query, setQuery] = useState("");
  const [primary, setPrimary] = useState<ScriptureTranslation>("web");
  const [comparison, setComparison] = useState<ScriptureTranslation | "none">("kjv");
  const [results, setResults] = useState<ScriptureResult[]>([]);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const comparisonVersion = comparison === "none" ? null : comparison;

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
      translation: scriptureVersion(result, primary, comparisonVersion),
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
          setPrimary(nextPrimary);
          if (comparison === nextPrimary) setComparison(nextPrimary === "web" ? "kjv" : "web");
        }}>
          {translations.map((version) => <option key={version} value={version}>{version.toUpperCase()}</option>)}
        </select>
      </label>
      <label>COMPARE
        <select value={comparison} onChange={(event) => setComparison(event.currentTarget.value as ScriptureTranslation | "none")}>
          <option value="none">Off</option>
          {translations.filter((version) => version !== primary).map((version) => <option key={version} value={version}>{version.toUpperCase()}</option>)}
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
        <button className="scripture-result__live" onClick={() => onGoLive({ reference: result.reference, text: scriptureText(result), translation: scriptureVersion(result, primary, comparisonVersion) })}>GO LIVE</button>
      </article>)}
      {results.length === 0 && !busy && !error && <div className="scripture-results__empty">The complete KJV and World English Bible are stored with the app for offline preparation.</div>}
    </div>
  </section>;
}
