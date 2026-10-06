import { useState } from "react";
import type { CSSProperties, ReactNode } from "react";
import type { ServiceItem, Slide } from "../types/presentation";
import { getItemSequenceSlides } from "../data/presentationNavigation";
import { SECTION_PALETTES } from "../data/songLyrics";
import { SlideCanvas } from "./SlideCanvas";

interface ServiceItemOverviewProps {
  readonly item: ServiceItem;
  readonly selectedSlideId: string;
  readonly liveItemPosition: number | null;
  readonly nextItemPosition: number | null;
  readonly size: number;
  readonly onSizeChange: (size: number) => void;
  readonly onSelectSlide: (slideId: string) => void;
  readonly onEditSlide: (slideId: string) => void;
  readonly onGoLive: (slide: Slide, itemPosition: number) => void;
  readonly onAddSlide: () => void;
  readonly onItemChange: (update: Partial<ServiceItem>) => void;
  readonly onOpenStructure?: () => void;
}

const bibleVersions = ["WEB", "KJV", "ASV", "BSB", "NIV"];

export function ServiceItemOverview({
  item,
  selectedSlideId,
  liveItemPosition,
  nextItemPosition,
  size,
  onSizeChange,
  onSelectSlide,
  onEditSlide,
  onGoLive,
  onAddSlide,
  onItemChange,
  onOpenStructure,
}: ServiceItemOverviewProps): ReactNode {
  const listMode = size <= 4;
  const arrangedSlides = getItemSequenceSlides(item);
  const activeArrangement = item.arrangements?.find((a) => a.id === item.activeArrangementId) ?? item.arrangements?.[0];

  return (
    <section className="item-overview" aria-label={`${item.title} overview`}>
      {/* Apple-style Heading with Context Badges */}
      <header className="item-overview__heading">
        <div className="item-overview__info">
          <div className="item-overview__eyebrow-row">
            <span className="eyebrow">{item.type.toUpperCase()}</span>
            {item.key && <span className="music-tag">KEY: <strong>{item.key}</strong></span>}
            {item.bpm && <span className="music-tag">BPM: <strong>{item.bpm}</strong></span>}
            {item.timeSignature && <span className="music-tag">{item.timeSignature}</span>}
          </div>

          <h2>{item.title}</h2>
          
          {(item.artist || item.author) && (
            <div className="item-overview__credits">
              {item.artist && <span>{item.artist}</span>}
              {item.artist && item.author && <span className="dot-sep">·</span>}
              {item.author && <span>{item.author}</span>}
            </div>
          )}

          <span className="item-overview__count">
            {arrangedSlides.length} {arrangedSlides.length === 1 ? "slide" : "slides"}
          </span>
        </div>

        <div className="item-overview__top-actions">
          {item.type === "song" && onOpenStructure && (
            <button
              type="button"
              className="structure-arrange-btn"
              onClick={onOpenStructure}
              title="Edit song structure, bricks, and arrangement flow"
            >
              ☰ Structure & Arrangement
            </button>
          )}
          <button className="add-overview-slide" onClick={onAddSlide}>
            + Add slide
          </button>
        </div>
      </header>

      {/* Song Arrangement Bar (The Visual Sequence Bricks) */}
      {item.type === "song" && item.sections && item.sections.length > 0 && (
        <section className="song-overview-arrangement-bar">
          <div className="arrangement-bar-header">
            <span className="eyebrow">ARRANGEMENT: <strong>{activeArrangement?.name || "Active Order"}</strong></span>
            {onOpenStructure && (
              <button
                type="button"
                className="apple-link-btn"
                onClick={onOpenStructure}
              >
                Change flow...
              </button>
            )}
          </div>

          <div className="arrangement-bricks-strip">
            {(item.activeSequence || item.sections.map((s) => s.id)).map((sectionId, index) => {
              const sec = item.sections?.find((s) => s.id === sectionId);
              if (!sec) return null;
              const meta = SECTION_PALETTES[sec.type] ?? SECTION_PALETTES.other;

              return (
                <div
                  key={`${sectionId}-${index}`}
                  className="mini-sequence-brick"
                  style={{ "--brick-color": meta.color, "--brick-bg": meta.bgRgba } as CSSProperties}
                  title={`${index + 1}. ${sec.name}`}
                  onClick={onOpenStructure}
                >
                  <span className="mini-brick-badge">{sec.label}</span>
                  <span className="mini-brick-name">{sec.name}</span>
                </div>
              );
            })}
          </div>
        </section>
      )}

      {/* Bible settings if Bible reading */}
      {item.type === "bible" && (
        <div className="bible-item-settings">
          <label>
            <span>PASSAGE</span>
            <input
              value={item.reference ?? ""}
              onChange={(event) => onItemChange({ reference: event.currentTarget.value })}
              placeholder="John 3:16-21"
            />
          </label>
          <label>
            <span>VERSION</span>
            <select
              value={item.bibleVersion ?? "WEB"}
              onChange={(event) => onItemChange({ bibleVersion: event.currentTarget.value })}
            >
              {bibleVersions.map((version) => (
                <option key={version}>{version}</option>
              ))}
            </select>
          </label>
          <span className="bible-settings-note">
            Passage and version are labels only; text remains manually editable.
          </span>
        </div>
      )}

      {/* Slide Cards Grid / List */}
      <div
        className={`item-overview__slides${listMode ? " item-overview__slides--list" : ""}`}
        style={{ "--overview-card-size": `${140 + size * 2}px` } as CSSProperties}
      >
        {arrangedSlides.map((slide, index) => {
          const secMeta = slide.sectionType ? SECTION_PALETTES[slide.sectionType] : null;
          const badgeColor = secMeta?.color ?? "#79847d";

          return (
            <article
              className={`overview-slide${slide.id === selectedSlideId ? " is-selected" : ""}${index + 1 === liveItemPosition ? " is-live" : ""}${index + 1 === nextItemPosition ? " is-next" : ""}`}
              key={`${slide.id}-${index}`}
              style={{ "--group-color": badgeColor } as CSSProperties}
            >
              <button
                className="overview-slide__select"
                aria-label={`Select slide ${index + 1}`}
                onClick={() => onSelectSlide(slide.id)}
                onDoubleClick={() => onGoLive(slide, index + 1)}
              >
                {!listMode && <SlideCanvas slide={slide} screenMode="slide" variant="overview" />}
                {(slide.cueMacroIds?.length ?? 0) > 0 && (
                  <span className="slide-cue-badge" title={`${slide.cueMacroIds?.length} automatic cue(s)`}>
                    ↯ {slide.cueMacroIds?.length}
                  </span>
                )}
                <span className="overview-slide__meta">
                  <span className="overview-slide__number">{String(index + 1).padStart(2, "0")}</span>
                  {slide.sectionLabel && <span className="overview-slide__section">{slide.sectionLabel}</span>}
                  <span className="overview-slide__text">{slide.text.replace(/\s+/g, " ").trim() || "Empty slide"}</span>
                </span>
              </button>
              <button className="overview-slide__edit" onClick={() => onEditSlide(slide.id)}>
                Edit
              </button>
              <button
                className="overview-slide__go-live"
                onClick={() => onGoLive(slide, index + 1)}
                disabled={index + 1 === liveItemPosition}
              >
                {index + 1 === liveItemPosition ? "ON AIR" : "GO LIVE"}
              </button>
            </article>
          );
        })}
        {item.slides.length === 0 && (
          <div className="overview-empty">This item does not have any slides yet.</div>
        )}
      </div>

      <footer className="item-overview__zoom">
        <span>TEXT</span>
        <input
          aria-label="Slide overview size"
          type="range"
          min="0"
          max="100"
          value={size}
          onChange={(event) => onSizeChange(Number(event.currentTarget.value))}
        />
        <span>SLIDES</span>
      </footer>
    </section>
  );
}
