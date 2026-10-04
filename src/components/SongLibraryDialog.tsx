import { useState } from "react";
import type { SyntheticEvent } from "react";
import { parseSongLyrics, songSlides, type ReflowMode } from "../data/songLyrics";
import type { Arrangement, ServiceItem, Slide, SlideGroup, Song } from "../types/presentation";

interface SongLibraryDialogProps {
  readonly songs: Song[];
  readonly docked?: boolean;
  readonly onClose: () => void;
  readonly onSave: (song: Song) => Promise<void>;
  readonly onDelete: (songId: string) => Promise<void>;
  readonly onAddToService: (item: ServiceItem) => void;
}

type SongDialogTab = "library" | "new";

function createSongItem(song: Song, slides: Slide[]): ServiceItem {
  const groups: SlideGroup[] = [];
  const palette = ["#58a6d8", "#b48ae0", "#df9b62", "#58b99b", "#d77685", "#b5a54f"];
  const orderedGroups: string[] = [];
  const nameOccurrences = new Map<string, number>();
  let previousKey = "";
  let currentGroup: SlideGroup | null = null;
  slides.forEach((slide, index) => {
    const label = slide.sectionLabel?.trim() || `Section ${index + 1}`;
    const key = slide.sectionLabel?.trim() ? label.toLocaleLowerCase() : `${label}:${index}`;
    if (key !== previousKey || !currentGroup) {
      const occurrence = (nameOccurrences.get(key) ?? 0) + 1;
      nameOccurrences.set(key, occurrence);
      currentGroup = {
        id: crypto.randomUUID(),
        name: occurrence > 1 ? `${label} ${occurrence}` : label,
        color: palette[groups.length % palette.length],
        slideIds: [],
      };
      groups.push(currentGroup);
      previousKey = key;
    }
    if (currentGroup) {
      currentGroup.slideIds.push(slide.id);
      orderedGroups.push(currentGroup.id);
    }
  });
  const arrangement: Arrangement = { id: crypto.randomUUID(), name: "Original order", groupIds: orderedGroups };
  return {
    id: crypto.randomUUID(),
    title: song.title,
    type: "song",
    slides,
    groups,
    arrangements: [arrangement],
    activeArrangementId: arrangement.id,
    reference: "",
    bibleVersion: "WEB",
  };
}

export function SongLibraryDialog({ songs, onClose, onSave, onDelete, onAddToService, docked = false }: SongLibraryDialogProps) {
  const [tab, setTab] = useState<SongDialogTab>("library");
  const [selectedSongId, setSelectedSongId] = useState(songs[0]?.id ?? "");
  const [title, setTitle] = useState("");
  const [lyrics, setLyrics] = useState("");
  const [reflowMode, setReflowMode] = useState<ReflowMode>("markers");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const selectedSong = songs.find((song) => song.id === selectedSongId) ?? null;
  const selectedSections = selectedSong ? parseSongLyrics(selectedSong.lyrics, selectedSong.reflowMode ?? "markers") : [];

  function addSavedSong() {
    if (!selectedSong) return;
    const slides = songSlides(selectedSong.lyrics, selectedSong.reflowMode ?? "markers");
    if (slides.length === 0) {
      setError("This song has no slide text. Add lyrics before using it in a service.");
      return;
    }
    onAddToService(createSongItem(selectedSong, slides));
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
    const slides = songSlides(cleanLyrics, reflowMode);
    if (slides.length === 0) {
      setError("No lyric slides were found. Add a verse or other lyric text.");
      return;
    }
    const song: Song = { id: crypto.randomUUID(), title: cleanTitle, lyrics: cleanLyrics, reflowMode };
    setSaving(true);
    setError("");
    try {
      await onSave(song);
      onAddToService(createSongItem(song, slides));
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
            <span className="eyebrow">SONG LIBRARY</span>
            <h2 id="song-dialog-title">Add song</h2>
          </div>
          <button type="button" className="dialog-close" aria-label="Close" onClick={onClose}>×</button>
        </div>
        <div className="song-library-tabs" role="tablist" aria-label="Song library views">
          <button role="tab" aria-selected={tab === "library"} className={tab === "library" ? "is-active" : ""} onClick={() => { setTab("library"); setError(""); }}>Library <span>{songs.length}</span></button>
          <button role="tab" aria-selected={tab === "new"} className={tab === "new" ? "is-active" : ""} onClick={() => { setTab("new"); setError(""); }}>New song</button>
        </div>

        {tab === "library" ? (
          <div className="song-library-content">
            <div className="song-library-list" aria-label="Saved songs">
              {songs.map((song) => (
                <button
                  key={song.id}
                  className={`song-library-row${song.id === selectedSongId ? " is-selected" : ""}`}
                  onClick={() => { setSelectedSongId(song.id); setError(""); }}
                >
                  <strong>{song.title}</strong>
                  <span>{parseSongLyrics(song.lyrics).length} slides</span>
                </button>
              ))}
              {songs.length === 0 && <p className="song-library-empty">No saved songs yet. Add lyrics in the New song tab.</p>}
            </div>
            <div className="song-library-detail">
              {selectedSong ? (
                <>
                  <div className="song-library-detail__heading">
                    <div><span className="eyebrow">{selectedSong.title}</span><strong>{selectedSections.length} lyric slides</strong></div>
                    <button className="danger-action" onClick={() => void removeSelectedSong()}>Remove</button>
                  </div>
                  <div className="song-library-sections">
                    {selectedSections.map((section, index) => (
                      <div key={`${section.label ?? "section"}-${index}`}>
                        {section.label && <span>{section.label}</span>}
                        <p>{section.text}</p>
                      </div>
                    ))}
                  </div>
                  <button className="present-button" onClick={addSavedSong}>Add to service</button>
                </>
              ) : (
                <div className="song-library-empty">Select a saved song or add new lyrics.</div>
              )}
            </div>
          </div>
        ) : (
          <form className="new-song-form" onSubmit={(event) => void saveAndAdd(event)}>
            <label className="dialog-field">
              <span>Song title</span>
              <input maxLength={100} required value={title} onChange={(event) => setTitle(event.currentTarget.value)} placeholder="Amazing Grace" />
            </label>
            <label className="dialog-field">
              <span>Lyrics and structure</span>
              <textarea
                required
                value={lyrics}
                onChange={(event) => setLyrics(event.currentTarget.value)}
                placeholder={'[Verse 1]\nAmazing grace how sweet the sound\nThat saved a wretch like me\n\n---\n\n[Chorus]\nMy chains are gone, I\'ve been set free'}
              />
            </label>
            <label className="dialog-field song-reflow-mode">
              <span>Split into slides</span>
              <select value={reflowMode} onChange={(event) => setReflowMode(event.currentTarget.value as ReflowMode)}>
                <option value="markers">By section headings and --- markers</option>
                <option value="blank-lines">At blank lines</option>
                <option value="four-lines">Every four lyric lines</option>
              </select>
            </label>
            <p className="song-format-hint">Add [Verse 1], [Chorus], or [Bridge] headings. The reflow option controls how pasted lyrics become slides.</p>
            <div className="dialog-actions">
              <button type="button" className="secondary-button" onClick={() => setTab("library")}>Cancel</button>
              <button type="submit" className="present-button" disabled={saving}>{saving ? "Adding..." : "Save and add to service"}</button>
            </div>
          </form>
        )}
        {error && <p className="song-dialog-error" role="alert">{error}</p>}
      </dialog>
    </div>
  );
}
