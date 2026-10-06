import { useEffect, useState } from "react";
import type { SyntheticEvent } from "react";
import {
  SECTION_PALETTES,
  generateSlidesFromArrangement,
  parseLyricsToSections,
  type ReflowMode,
} from "../data/songLyrics";
import { peekCachedLyrics, searchLyrics, type LyricsResult } from "../data/lyricsLookup";
import type { ServiceItem, Song, SongArrangementPreset, SongSection } from "../types/presentation";

interface SongLibraryDialogProps {
  readonly songs: Song[];
  readonly docked?: boolean;
  readonly onClose: () => void;
  readonly onSave: (song: Song) => Promise<void>;
  readonly onDelete: (songId: string) => Promise<void>;
  readonly onAddToService: (item: ServiceItem) => void;
}

type SongDialogTab = "library" | "new";

function createSmartSongItem(song: Song): ServiceItem {
  const sections: SongSection[] = song.sections && song.sections.length > 0
    ? song.sections
    : parseLyricsToSections(song.lyrics, song.reflowMode ?? "markers");

  const arrangements: SongArrangementPreset[] = song.arrangements && song.arrangements.length > 0
    ? song.arrangements
    : [
        {
          id: "arr-default",
          name: "Standard Flow",
          sectionIds: sections.map((s) => s.id),
        },
      ];

  const activeArrangementId = song.defaultArrangementId || arrangements[0]?.id || "arr-default";
  const activeArr = arrangements.find((a) => a.id === activeArrangementId) || arrangements[0];
  const activeSequence = activeArr?.sectionIds || sections.map((s) => s.id);
  const slides = generateSlidesFromArrangement(sections, activeSequence);

  return {
    id: crypto.randomUUID(),
    title: song.title,
    type: "song",
    artist: song.artist,
    author: song.author,
    key: song.key,
    bpm: song.bpm,
    timeSignature: song.timeSignature,
    songId: song.id,
    sections,
    arrangements,
    activeArrangementId,
    activeSequence,
    slides,
    reference: "",
    bibleVersion: "WEB",
  };
}

export function SongLibraryDialog({
  songs,
  onClose,
  onSave,
  onDelete,
  onAddToService,
  docked = false,
}: SongLibraryDialogProps) {
  const [tab, setTab] = useState<SongDialogTab>("library");
  const [selectedSongId, setSelectedSongId] = useState(songs[0]?.id ?? "");
  const [title, setTitle] = useState("");
  const [artist, setArtist] = useState("");
  const [author, setAuthor] = useState("");
  const [key, setKey] = useState("A");
  const [bpm, setBpm] = useState("63");
  const [timeSignature, setTimeSignature] = useState("4/4");
  const [lyrics, setLyrics] = useState("");
  const [reflowMode, setReflowMode] = useState<ReflowMode>("markers");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const [libraryFilter, setLibraryFilter] = useState("");
  const [lyricQuery, setLyricQuery] = useState("");
  const [lyricResults, setLyricResults] = useState<LyricsResult[]>([]);
  const [lyricStatus, setLyricStatus] = useState<"idle" | "loading" | "done" | "error">("idle");
  const [lyricError, setLyricError] = useState("");
  const [lyricRetry, setLyricRetry] = useState(0);
  const [pickedLyricId, setPickedLyricId] = useState("");

  // Search as you type: wait for a pause, cancel the previous request, show results as soon as any arrive.
  useEffect(() => {
    const query = lyricQuery.trim();
    if (query.length < 3) {
      setLyricResults([]);
      setLyricStatus("idle");
      setLyricError("");
      return undefined;
    }
    const cached = peekCachedLyrics(query);
    if (cached) {
      setLyricResults(cached);
      setLyricStatus("done");
      setLyricError("");
      return undefined;
    }
    const controller = new AbortController();
    const timer = window.setTimeout(() => {
      setLyricStatus("loading");
      setLyricError("");
      searchLyrics(query, {
        signal: controller.signal,
        onResults: (partial) => {
          if (!controller.signal.aborted) setLyricResults(partial);
        },
      })
        .then((final) => {
          if (controller.signal.aborted) return;
          setLyricResults(final);
          setLyricStatus("done");
        })
        .catch((cause: unknown) => {
          if (controller.signal.aborted) return;
          setLyricResults([]);
          setLyricError(cause instanceof Error ? cause.message : "The lyric search failed.");
          setLyricStatus("error");
        });
    }, 400);
    return () => {
      window.clearTimeout(timer);
      controller.abort();
    };
  }, [lyricQuery, lyricRetry]);

  const filterText = libraryFilter.trim().toLocaleLowerCase();
  const visibleSongs = filterText
    ? songs.filter((song) => `${song.title} ${song.artist ?? ""} ${song.author ?? ""} ${song.lyrics}`.toLocaleLowerCase().includes(filterText))
    : songs;

  function pickLyricResult(result: LyricsResult) {
    setTitle(result.title);
    setArtist(result.artist);
    setLyrics(result.lyrics);
    // Online lyrics rarely have [Verse]/[Chorus] markers - split on blank lines instead.
    setReflowMode(/^\s*\[[^\]]+\]\s*$/m.test(result.lyrics) ? "markers" : "blank-lines");
    setPickedLyricId(result.id);
    setError("");
  }

  const selectedSong = songs.find((song) => song.id === selectedSongId) ?? null;
  const selectedSections = selectedSong
    ? selectedSong.sections && selectedSong.sections.length > 0
      ? selectedSong.sections
      : parseLyricsToSections(selectedSong.lyrics, selectedSong.reflowMode ?? "markers")
    : [];

  function addSavedSong() {
    if (!selectedSong) return;
    const item = createSmartSongItem(selectedSong);
    if (item.slides.length === 0) {
      setError("This song has no slide text. Add lyrics before using it in a service.");
      return;
    }
    onAddToService(item);
    onClose();
  }

  async function saveAndAdd(event: SyntheticEvent<HTMLFormElement>) {
    event.preventDefault();
    const cleanTitle = title.trim();
    const cleanLyrics = lyrics.trim();
    if (!cleanTitle || !cleanLyrics) {
      setError("Enter a song title and lyrics.");
      return;
    }

    const sections = parseLyricsToSections(cleanLyrics, reflowMode);
    if (sections.length === 0) {
      setError("No lyric sections were found. Add [Verse 1], [Chorus], or lyric text.");
      return;
    }

    const defaultArr: SongArrangementPreset = {
      id: "arr-default",
      name: "Standard Flow",
      sectionIds: sections.map((s) => s.id),
    };

    const song: Song = {
      id: crypto.randomUUID(),
      title: cleanTitle,
      artist: artist.trim() || undefined,
      author: author.trim() || undefined,
      key: key.trim() || undefined,
      bpm: bpm.trim() || undefined,
      timeSignature: timeSignature.trim() || undefined,
      lyrics: cleanLyrics,
      sections,
      arrangements: [defaultArr],
      defaultArrangementId: defaultArr.id,
      reflowMode,
    };

    setSaving(true);
    setError("");
    try {
      await onSave(song);
      onAddToService(createSmartSongItem(song));
      onClose();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
      setSaving(false);
    }
  }

  async function removeSelectedSong() {
    if (!selectedSong || !window.confirm(`Remove "${selectedSong.title}" from the song library?`)) return;
    try {
      await onDelete(selectedSong.id);
      setSelectedSongId(songs.find((song) => song.id !== selectedSong.id)?.id ?? "");
      setError("");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  return (
    <div className={`dialog-backdrop song-dialog-backdrop${docked ? " song-dialog-backdrop--dock" : ""}`}>
      <dialog open className="service-dialog song-library-dialog" aria-labelledby="song-dialog-title">
        <div className="dialog-heading">
          <div>
            <span className="eyebrow">SMART SONG LIBRARY</span>
            <h2 id="song-dialog-title">Song Library</h2>
          </div>
          <button type="button" className="dialog-close" aria-label="Close" onClick={onClose}>
            ×
          </button>
        </div>

        <div className="song-library-tabs" role="tablist" aria-label="Song library views">
          <button
            role="tab"
            aria-selected={tab === "library"}
            className={tab === "library" ? "is-active" : ""}
            onClick={() => {
              setTab("library");
              setError("");
            }}
          >
            Library <span>{songs.length}</span>
          </button>
          <button
            role="tab"
            aria-selected={tab === "new"}
            className={tab === "new" ? "is-active" : ""}
            onClick={() => {
              setTab("new");
              setError("");
            }}
          >
            + Add New Song
          </button>
        </div>

        {tab === "library" ? (
          <div className="song-library-content">
            <div className="song-library-list" aria-label="Saved songs">
              <input
                className="song-library-filter"
                type="search"
                value={libraryFilter}
                onChange={(event) => setLibraryFilter(event.currentTarget.value)}
                placeholder="Search saved songs or lyrics..."
                aria-label="Search saved songs"
              />
              {visibleSongs.map((song) => (
                <button
                  key={song.id}
                  className={`song-library-row${song.id === selectedSongId ? " is-selected" : ""}`}
                  onClick={() => {
                    setSelectedSongId(song.id);
                    setError("");
                  }}
                >
                  <div className="song-row-main">
                    <strong>{song.title}</strong>
                    {song.artist && <span className="song-row-artist">{song.artist}</span>}
                  </div>
                  <div className="song-row-meta">
                    {song.key && <span className="key-tag">{song.key}</span>}
                    {song.bpm && <span className="bpm-tag">{song.bpm} BPM</span>}
                  </div>
                </button>
              ))}
              {songs.length === 0 && (
                <p className="song-library-empty">No saved songs yet. Add lyrics in the New song tab.</p>
              )}
              {songs.length > 0 && visibleSongs.length === 0 && (
                <p className="song-library-empty">No saved song matches "{libraryFilter.trim()}".</p>
              )}
            </div>

            <div className="song-library-detail">
              {selectedSong ? (
                <>
                  <div className="song-library-detail__heading">
                    <div>
                      <div className="detail-eyebrow-row">
                        <span className="eyebrow">SMART SONG</span>
                        {selectedSong.key && <span className="music-tag">KEY: <strong>{selectedSong.key}</strong></span>}
                        {selectedSong.bpm && <span className="music-tag">BPM: <strong>{selectedSong.bpm}</strong></span>}
                        {selectedSong.timeSignature && <span className="music-tag">{selectedSong.timeSignature}</span>}
                      </div>
                      <h3>{selectedSong.title}</h3>
                      {(selectedSong.artist || selectedSong.author) && (
                        <p className="song-artist-line">
                          {selectedSong.artist} {selectedSong.author && `· ${selectedSong.author}`}
                        </p>
                      )}
                    </div>
                    <button className="danger-action" onClick={() => void removeSelectedSong()}>
                      Remove
                    </button>
                  </div>

                  {/* Section Bricks Overview */}
                  <div className="song-library-bricks-preview">
                    <span className="bricks-title">SECTIONS ({selectedSections.length}):</span>
                    <div className="bricks-list">
                      {selectedSections.map((sec, idx) => {
                        const meta = SECTION_PALETTES[sec.type] ?? SECTION_PALETTES.other;
                        return (
                          <span
                            key={`${sec.id}-${idx}`}
                            className="preview-brick-badge"
                            style={{ backgroundColor: meta.bgRgba, color: meta.color, borderColor: meta.color }}
                          >
                            {sec.label}: {sec.name}
                          </span>
                        );
                      })}
                    </div>
                  </div>

                  <div className="song-library-sections">
                    {selectedSections.map((section, index) => (
                      <div key={`${section.id}-${index}`} className="song-section-card">
                        <span className="section-label-chip">{section.name}</span>
                        <p>{section.lines || (section.slides ? section.slides.join("\n\n") : "")}</p>
                      </div>
                    ))}
                  </div>

                  <button className="present-button song-add-service-btn" onClick={addSavedSong}>
                    Add to service with Smart Arrangement
                  </button>
                </>
              ) : (
                <div className="song-library-empty">Select a saved song or add new lyrics.</div>
              )}
            </div>
          </div>
        ) : (
          <>
          <div className="lyrics-search">
            <label className="dialog-field">
              <span>Find lyrics online (optional)</span>
              <input
                type="search"
                value={lyricQuery}
                onChange={(event) => setLyricQuery(event.currentTarget.value)}
                placeholder="Song title, or title - artist"
                aria-label="Search for song lyrics online"
                autoComplete="off"
              />
            </label>
            {lyricStatus === "loading" && <p className="lyrics-search__status" role="status">Searching...</p>}
            {lyricStatus === "error" && (
              <p className="lyrics-search__status lyrics-search__status--error" role="alert">
                {lyricError}
                <button type="button" onClick={() => setLyricRetry((count) => count + 1)}>Try again</button>
              </p>
            )}
            {lyricStatus === "done" && lyricResults.length === 0 && (
              <p className="lyrics-search__status">No lyrics found. Try just the title, or paste the lyrics below.</p>
            )}
            {lyricResults.length > 0 && (
              <ul className="lyrics-search__results" aria-label="Lyric search results">
                {lyricResults.map((result) => (
                  <li key={result.id}>
                    <button
                      type="button"
                      className={`lyrics-search__result${result.id === pickedLyricId ? " is-picked" : ""}`}
                      onClick={() => pickLyricResult(result)}
                    >
                      <strong>{result.title}</strong>
                      <span>{result.artist || "Unknown artist"}{result.album ? ` · ${result.album}` : ""}</span>
                      <small>{result.snippet}</small>
                    </button>
                  </li>
                ))}
              </ul>
            )}
            {pickedLyricId && <p className="lyrics-search__status">Lyrics filled in below - check them and add [Verse]/[Chorus] markers if you like. Make sure your church's CCLI licence covers this song.</p>}
          </div>
          <form className="new-song-form" onSubmit={(event) => void saveAndAdd(event)}>
            <div className="form-two-col">
              <label className="dialog-field">
                <span>Song title</span>
                <input
                  maxLength={100}
                  required
                  value={title}
                  onChange={(event) => setTitle(event.currentTarget.value)}
                  placeholder="Goodness of God"
                />
              </label>

              <label className="dialog-field">
                <span>Artist / Lead</span>
                <input
                  maxLength={100}
                  value={artist}
                  onChange={(event) => setArtist(event.currentTarget.value)}
                  placeholder="Jenn Johnson"
                />
              </label>
            </div>

            <div className="form-three-col">
              <label className="dialog-field">
                <span>Musical Key</span>
                <input
                  maxLength={10}
                  value={key}
                  onChange={(event) => setKey(event.currentTarget.value)}
                  placeholder="A, G, C..."
                />
              </label>

              <label className="dialog-field">
                <span>BPM</span>
                <input
                  type="number"
                  min="30"
                  max="240"
                  value={bpm}
                  onChange={(event) => setBpm(event.currentTarget.value)}
                  placeholder="63"
                />
              </label>

              <label className="dialog-field">
                <span>Time Signature</span>
                <input
                  maxLength={10}
                  value={timeSignature}
                  onChange={(event) => setTimeSignature(event.currentTarget.value)}
                  placeholder="4/4, 6/8..."
                />
              </label>
            </div>

            <label className="dialog-field">
              <span>Author / Publisher / CCLI</span>
              <input
                maxLength={100}
                value={author}
                onChange={(event) => setAuthor(event.currentTarget.value)}
                placeholder="Bethel Music Publishing · CCLI #7117726"
              />
            </label>

            <label className="dialog-field">
              <span>Lyrics and section structure</span>
              <textarea
                required
                value={lyrics}
                onChange={(event) => setLyrics(event.currentTarget.value)}
                rows={8}
                placeholder={
                  "[Verse 1]\nI love You Lord\nOh Your mercy never fails me\n\n[Chorus]\nAll my life You have been faithful\nAll my life You have been so so good\n\n[Bridge]\nYour goodness is running after\nIt's running after me"
                }
              />
            </label>

            <p className="song-format-hint">
              💡 Use [Verse 1], [Chorus], [Bridge], [Tag], [Intro], [Outro] headers. Slides and arrangement bricks will automatically generate!
            </p>

            <div className="dialog-actions">
              <button type="button" className="secondary-button" onClick={() => setTab("library")}>
                Cancel
              </button>
              <button type="submit" className="present-button" disabled={saving}>
                {saving ? "Saving..." : "Save Smart Song & Add to Service"}
              </button>
            </div>
          </form>
          </>
        )}
        {error && (
          <p className="song-dialog-error" role="alert">
            {error}
          </p>
        )}
      </dialog>
    </div>
  );
}
