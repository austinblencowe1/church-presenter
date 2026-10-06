import { useState } from "react";
import type { CSSProperties, DragEvent, ReactNode } from "react";
import type { ServiceItem, Slide, Song, SongArrangementPreset, SongSection, SongSectionType } from "../types/presentation";
import {
  SECTION_PALETTES,
  detectSectionBadge,
  detectSectionType,
  generateSlidesFromArrangement,
  parseLyricsToSections,
} from "../data/songLyrics";
import { SlideCanvas } from "./SlideCanvas";

interface SongStructureModalProps {
  readonly item: ServiceItem;
  readonly baseSlideStyle?: Partial<Slide>;
  readonly onClose: () => void;
  readonly onSave: (updatedItem: ServiceItem, saveToLibrary?: boolean) => void;
}

const MUSICAL_KEYS = ["C", "C# / Db", "D", "Eb", "E", "F", "F# / Gb", "G", "Ab", "A", "Bb", "B"];
const TIME_SIGNATURES = ["4/4", "6/8", "3/4", "2/4", "12/8"];

export function SongStructureModal({
  item,
  baseSlideStyle,
  onClose,
  onSave,
}: SongStructureModalProps): ReactNode {
  // Extract or initialize sections
  const initialSections: SongSection[] = (item.sections && item.sections.length > 0)
    ? item.sections
    : (item.slides && item.slides.length > 0
        ? parseLyricsToSections(item.slides.map((s) => `[${s.sectionLabel || "Section"}]\n${s.text}`).join("\n\n"))
        : parseLyricsToSections(""));

  const [sections, setSections] = useState<SongSection[]>(initialSections);

  // Arrangements & active sequence
  const initialArrangements: SongArrangementPreset[] = (item.arrangements && item.arrangements.length > 0)
    ? item.arrangements
    : [
        {
          id: "arr-default",
          name: "Standard Flow",
          sectionIds: initialSections.map((s) => s.id),
        },
      ];

  const [arrangements, setArrangements] = useState<SongArrangementPreset[]>(initialArrangements);
  const [selectedArrangementId, setSelectedArrangementId] = useState<string>(
    item.activeArrangementId || initialArrangements[0]?.id || "arr-default"
  );

  // Active sequence of section IDs
  const activeArrangement = arrangements.find((a) => a.id === selectedArrangementId) || arrangements[0];
  const [sequence, setSequence] = useState<string[]>(
    item.activeSequence && item.activeSequence.length > 0
      ? item.activeSequence
      : (activeArrangement?.sectionIds || initialSections.map((s) => s.id))
  );

  // Metadata
  const [title, setTitle] = useState(item.title || "Untitled Song");
  const [artist, setArtist] = useState(item.artist || "");
  const [author, setAuthor] = useState(item.author || "");
  const [key, setKey] = useState(item.key || "A");
  const [bpm, setBpm] = useState<string | number>(item.bpm || "63");
  const [timeSignature, setTimeSignature] = useState(item.timeSignature || "4/4");

  // Preset management
  const [isCreatingPreset, setIsCreatingPreset] = useState(false);
  const [newPresetName, setNewPresetName] = useState("");
  const [saveToLibrary, setSaveToLibrary] = useState(true);

  // New section creation
  const [isAddingSection, setIsAddingSection] = useState(false);
  const [newSectionType, setNewSectionType] = useState<SongSectionType>("verse");
  const [newSectionName, setNewSectionName] = useState("");
  const [newSectionText, setNewSectionText] = useState("");

  // Drag-and-drop state
  const [draggedBrickIndex, setDraggedBrickIndex] = useState<number | null>(null);
  const [draggedPaletteId, setDraggedPaletteId] = useState<string | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);

  // Preview tab / selection
  const [previewSlideIndex, setPreviewSlideIndex] = useState(0);

  // Generated slides in real-time
  const generatedSlides = generateSlidesFromArrangement(sections, sequence, baseSlideStyle);

  // Handlers for Brick Arrangement Sequence
  function addBrickToSequence(sectionId: string, atIndex?: number) {
    if (typeof atIndex === "number" && atIndex >= 0) {
      const next = [...sequence];
      next.splice(atIndex, 0, sectionId);
      setSequence(next);
    } else {
      setSequence((prev) => [...prev, sectionId]);
    }
  }

  function removeBrickFromSequence(index: number) {
    setSequence((prev) => prev.filter((_, i) => i !== index));
  }

  function duplicateBrick(index: number) {
    const next = [...sequence];
    next.splice(index + 1, 0, sequence[index]);
    setSequence(next);
  }

  function moveBrick(fromIndex: number, toIndex: number) {
    if (toIndex < 0 || toIndex >= sequence.length) return;
    const next = [...sequence];
    const [moved] = next.splice(fromIndex, 1);
    next.splice(toIndex, 0, moved);
    setSequence(next);
  }

  function selectArrangement(presetId: string) {
    setSelectedArrangementId(presetId);
    const preset = arrangements.find((a) => a.id === presetId);
    if (preset) {
      setSequence([...preset.sectionIds]);
    }
  }

  function saveCurrentAsPreset() {
    const cleanName = newPresetName.trim();
    if (!cleanName) return;
    const newPreset: SongArrangementPreset = {
      id: `arr-${crypto.randomUUID().slice(0, 8)}`,
      name: cleanName,
      sectionIds: [...sequence],
    };
    setArrangements((prev) => [...prev, newPreset]);
    setSelectedArrangementId(newPreset.id);
    setNewPresetName("");
    setIsCreatingPreset(false);
  }

  function updateActivePreset() {
    setArrangements((prev) =>
      prev.map((a) =>
        a.id === selectedArrangementId ? { ...a, sectionIds: [...sequence] } : a
      )
    );
  }

  function handleCreateSection() {
    const name = newSectionName.trim() || `${newSectionType[0].toUpperCase()}${newSectionType.slice(1)}`;
    const label = detectSectionBadge(name, newSectionType);
    const meta = SECTION_PALETTES[newSectionType] ?? SECTION_PALETTES.other;
    const lines = newSectionText.trim();
    const slides = lines ? lines.split(/\n\s*---+\s*\n/).map((s) => s.trim()).filter(Boolean) : (newSectionType === "blank" ? [""] : []);

    const newSec: SongSection = {
      id: `sec-${newSectionType}-${crypto.randomUUID().slice(0, 6)}`,
      type: newSectionType,
      name,
      label,
      lines,
      slides,
      color: meta.color,
    };

    setSections((prev) => [...prev, newSec]);
    addBrickToSequence(newSec.id);
    setIsAddingSection(false);
    setNewSectionName("");
    setNewSectionText("");
  }

  function handleSaveAll() {
    const updatedSlides = generateSlidesFromArrangement(sections, sequence, baseSlideStyle);
    const updatedItem: ServiceItem = {
      ...item,
      title,
      artist,
      author,
      key,
      bpm,
      timeSignature,
      sections,
      arrangements,
      activeArrangementId: selectedArrangementId,
      activeSequence: sequence,
      slides: updatedSlides,
    };

    onSave(updatedItem, saveToLibrary);
  }

  return (
    <div className="dialog-backdrop song-structure-modal-backdrop" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="song-structure-modal" role="dialog" aria-modal="true" aria-labelledby="modal-title">
        
        {/* Apple-style Top Bar */}
        <header className="song-structure-header">
          <div className="song-structure-header__left">
            <span className="apple-pill-tag">SMART SONG</span>
            <input
              id="modal-title"
              className="song-structure-title-input"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              placeholder="Song Title"
              aria-label="Song Title"
            />
          </div>

          {/* Quick Musical Metadata */}
          <div className="song-structure-meta-pills">
            <div className="meta-pill" title="Musical Key">
              <span className="meta-pill__label">KEY</span>
              <select value={key} onChange={(e) => setKey(e.target.value)} aria-label="Musical Key">
                {MUSICAL_KEYS.map((k) => (
                  <option key={k} value={k}>{k}</option>
                ))}
              </select>
            </div>

            <div className="meta-pill" title="Beats Per Minute">
              <span className="meta-pill__label">BPM</span>
              <input
                type="number"
                min="30"
                max="240"
                value={bpm}
                onChange={(e) => setBpm(e.target.value)}
                aria-label="BPM"
              />
            </div>

            <div className="meta-pill" title="Time Signature">
              <span className="meta-pill__label">TIME</span>
              <select value={timeSignature} onChange={(e) => setTimeSignature(e.target.value)} aria-label="Time Signature">
                {TIME_SIGNATURES.map((ts) => (
                  <option key={ts} value={ts}>{ts}</option>
                ))}
              </select>
            </div>
          </div>

          <button className="apple-modal-close" onClick={onClose} aria-label="Close modal">
            ✕
          </button>
        </header>

        {/* Sub-header with Credits & Arrangement Presets */}
        <div className="song-structure-subbar">
          <div className="song-structure-credits">
            <input
              value={artist}
              onChange={(e) => setArtist(e.target.value)}
              placeholder="Artist (e.g. Jenn Johnson)"
              aria-label="Artist"
            />
            <span className="dot-sep">·</span>
            <input
              value={author}
              onChange={(e) => setAuthor(e.target.value)}
              placeholder="Author / Publisher (e.g. Bethel Music)"
              aria-label="Author / Publisher"
            />
          </div>

          {/* Arrangement Preset Bar */}
          <div className="arrangement-preset-control">
            <span className="preset-label">ARRANGEMENT:</span>
            <select
              value={selectedArrangementId}
              onChange={(e) => selectArrangement(e.target.value)}
              aria-label="Arrangement Preset"
            >
              {arrangements.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.name} ({a.sectionIds.length} bricks)
                </option>
              ))}
            </select>

            <button
              type="button"
              className="apple-button-subtle"
              onClick={updateActivePreset}
              title="Update current preset with active brick arrangement"
            >
              Update Preset
            </button>

            {!isCreatingPreset ? (
              <button
                type="button"
                className="apple-button-subtle"
                onClick={() => setIsCreatingPreset(true)}
              >
                + New Preset
              </button>
            ) : (
              <div className="new-preset-inline">
                <input
                  autoFocus
                  value={newPresetName}
                  onChange={(e) => setNewPresetName(e.target.value)}
                  placeholder="Preset name (e.g. Acoustic)"
                  onKeyDown={(e) => e.key === "Enter" && saveCurrentAsPreset()}
                />
                <button type="button" className="apple-button-accent" onClick={saveCurrentAsPreset}>
                  Save
                </button>
                <button type="button" className="apple-button-subtle" onClick={() => setIsCreatingPreset(false)}>
                  Cancel
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Main Work Area: Split Palette (Left) & Brick Timeline (Right/Center) */}
        <div className="song-structure-body">
          
          {/* Section Palette / Bank */}
          <section className="section-palette-shelf">
            <div className="shelf-heading">
              <span className="eyebrow">SECTION PALETTE</span>
              <button
                className="apple-button-micro"
                onClick={() => setIsAddingSection((v) => !v)}
              >
                {isAddingSection ? "Cancel" : "+ Add Section"}
              </button>
            </div>

            {isAddingSection && (
              <div className="add-section-panel">
                <select
                  value={newSectionType}
                  onChange={(e) => setNewSectionType(e.target.value as SongSectionType)}
                  aria-label="Section Type"
                >
                  <option value="verse">Verse</option>
                  <option value="chorus">Chorus</option>
                  <option value="bridge">Bridge</option>
                  <option value="tag">Tag</option>
                  <option value="pre-chorus">Pre-Chorus</option>
                  <option value="intro">Intro</option>
                  <option value="outro">Outro</option>
                  <option value="vamp">Vamp / Interlude</option>
                  <option value="blank">Blank / Instrumental</option>
                </select>

                <input
                  value={newSectionName}
                  onChange={(e) => setNewSectionName(e.target.value)}
                  placeholder="Section Name (e.g. Verse 3)"
                />

                {newSectionType !== "blank" && (
                  <textarea
                    rows={3}
                    value={newSectionText}
                    onChange={(e) => setNewSectionText(e.target.value)}
                    placeholder="Enter lyric lines for this section..."
                  />
                )}

                <button
                  type="button"
                  className="apple-button-accent"
                  onClick={handleCreateSection}
                >
                  Add to Song
                </button>
              </div>
            )}

            <div className="palette-bricks-list">
              {sections.map((sec) => {
                const meta = SECTION_PALETTES[sec.type] ?? SECTION_PALETTES.other;
                return (
                  <div
                    key={sec.id}
                    className="palette-brick"
                    style={{ "--brick-color": meta.color, "--brick-bg": meta.bgRgba } as CSSProperties}
                    draggable
                    onDragStart={(e) => {
                      e.dataTransfer.setData("text/plain", sec.id);
                      setDraggedPaletteId(sec.id);
                    }}
                    onDragEnd={() => setDraggedPaletteId(null)}
                    onClick={() => addBrickToSequence(sec.id)}
                    title="Click or drag to add to arrangement"
                  >
                    <span className="palette-brick__badge">{sec.label}</span>
                    <div className="palette-brick__info">
                      <strong>{sec.name}</strong>
                      <span className="palette-brick__preview">
                        {sec.lines ? sec.lines.split("\n")[0] : "(Instrumental / Blank)"}
                      </span>
                    </div>
                    <span className="palette-brick__plus">+</span>
                  </div>
                );
              })}
            </div>
            
            <p className="palette-hint">
              💡 Click any section brick or drag it into the arrangement timeline.
            </p>
          </section>

          {/* Center: Brick Timeline (The Arrangement) */}
          <section
            className="arrangement-timeline-panel"
            onDragOver={(e) => {
              e.preventDefault();
              e.dataTransfer.dropEffect = "move";
            }}
            onDrop={(e) => {
              e.preventDefault();
              if (draggedPaletteId) {
                addBrickToSequence(draggedPaletteId);
                setDraggedPaletteId(null);
              }
            }}
          >
            <div className="timeline-heading">
              <div>
                <span className="eyebrow">ARRANGEMENT TIMELINE</span>
                <h3>{sequence.length} Sections · {generatedSlides.length} Live Slides</h3>
              </div>

              <div className="timeline-quick-actions">
                <button
                  type="button"
                  className="apple-button-subtle"
                  onClick={() => setSequence([])}
                  disabled={sequence.length === 0}
                >
                  Clear All
                </button>
                <button
                  type="button"
                  className="apple-button-subtle"
                  onClick={() => setSequence(sections.map((s) => s.id))}
                >
                  Reset to Album Order
                </button>
              </div>
            </div>

            {/* The Connected Bricks Builder */}
            <div className="bricks-timeline-canvas">
              {sequence.map((sectionId, index) => {
                const section = sections.find((s) => s.id === sectionId);
                if (!section) return null;
                const meta = SECTION_PALETTES[section.type] ?? SECTION_PALETTES.other;
                const isOver = dragOverIndex === index;

                return (
                  <div
                    key={`${sectionId}-${index}`}
                    className={`arranged-brick${isOver ? " is-drag-over" : ""}`}
                    style={{ "--brick-color": meta.color, "--brick-bg": meta.bgRgba } as CSSProperties}
                    draggable
                    onDragStart={(e) => {
                      e.dataTransfer.setData("text/plain", String(index));
                      setDraggedBrickIndex(index);
                    }}
                    onDragOver={(e) => {
                      e.preventDefault();
                      setDragOverIndex(index);
                    }}
                    onDragLeave={() => {
                      if (dragOverIndex === index) setDragOverIndex(null);
                    }}
                    onDrop={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      setDragOverIndex(null);

                      if (draggedBrickIndex !== null && draggedBrickIndex !== index) {
                        moveBrick(draggedBrickIndex, index);
                        setDraggedBrickIndex(null);
                      } else if (draggedPaletteId) {
                        addBrickToSequence(draggedPaletteId, index);
                        setDraggedPaletteId(null);
                      }
                    }}
                    onDragEnd={() => {
                      setDraggedBrickIndex(null);
                      setDragOverIndex(null);
                    }}
                  >
                    <div className="arranged-brick__handle" title="Drag to reorder">
                      ☰
                    </div>
                    <span className="arranged-brick__index">{index + 1}</span>
                    <span className="arranged-brick__badge">{section.label}</span>
                    <div className="arranged-brick__content">
                      <strong>{section.name}</strong>
                      <small>{section.lines ? section.lines.split("\n")[0] : "Blank Slide"}</small>
                    </div>

                    <div className="arranged-brick__actions">
                      <button
                        type="button"
                        className="brick-btn"
                        onClick={() => moveBrick(index, index - 1)}
                        disabled={index === 0}
                        title="Move left"
                        aria-label="Move left"
                      >
                        ◀
                      </button>
                      <button
                        type="button"
                        className="brick-btn"
                        onClick={() => moveBrick(index, index + 1)}
                        disabled={index === sequence.length - 1}
                        title="Move right"
                        aria-label="Move right"
                      >
                        ▶
                      </button>
                      <button
                        type="button"
                        className="brick-btn brick-btn--dup"
                        onClick={() => duplicateBrick(index)}
                        title="Duplicate (×2)"
                        aria-label="Duplicate section"
                      >
                        +
                      </button>
                      <button
                        type="button"
                        className="brick-btn brick-btn--del"
                        onClick={() => removeBrickFromSequence(index)}
                        title="Remove section"
                        aria-label="Remove section"
                      >
                        ✕
                      </button>
                    </div>
                  </div>
                );
              })}

              {sequence.length === 0 && (
                <div className="empty-timeline-state">
                  <p>Drag section bricks here from the left palette to build the arrangement.</p>
                </div>
              )}
            </div>

            {/* Live Resulting Slide Strip */}
            <div className="generated-slides-strip">
              <div className="strip-heading">
                <span className="eyebrow">GENERATED PRESENTATION SLIDES ({generatedSlides.length})</span>
                <span className="strip-hint">Updated live from your arrangement above</span>
              </div>

              <div className="slides-carousel">
                {generatedSlides.map((slide, slideIdx) => (
                  <div
                    key={slide.id}
                    className={`carousel-slide-card${slideIdx === previewSlideIndex ? " is-active" : ""}`}
                    onClick={() => setPreviewSlideIndex(slideIdx)}
                  >
                    <div className="carousel-slide-thumb">
                      <SlideCanvas slide={slide} screenMode="slide" variant="thumbnail" />
                    </div>
                    <div className="carousel-slide-meta">
                      <span className="slide-num">{slideIdx + 1}</span>
                      <span className="slide-label">{slide.sectionLabel}</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>

          </section>
        </div>

        {/* Modal Footer */}
        <footer className="song-structure-footer">
          <label className="save-library-toggle">
            <input
              type="checkbox"
              checked={saveToLibrary}
              onChange={(e) => setSaveToLibrary(e.target.checked)}
            />
            <span>Update this song in the permanent Song Library</span>
          </label>

          <div className="footer-actions">
            <button type="button" className="apple-button-secondary" onClick={onClose}>
              Cancel
            </button>
            <button type="button" className="apple-button-primary" onClick={handleSaveAll}>
              Apply Structure to Service
            </button>
          </div>
        </footer>

      </div>
    </div>
  );
}
