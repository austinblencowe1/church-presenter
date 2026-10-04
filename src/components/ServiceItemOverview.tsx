import { useState } from "react";
import type { CSSProperties, SyntheticEvent } from "react";
import type { ServiceItem, Slide } from "../types/presentation";
import { getItemSequenceSlides } from "../data/presentationNavigation";
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
}

const bibleVersions = ["WEB", "KJV", "ASV", "BSB", "NIV"];
const groupColors = ["#58a6d8", "#b48ae0", "#df9b62", "#58b99b", "#d77685", "#b5a54f"];

function groupColor(label: string | null | undefined): string {
  if (!label) return "#79847d";
  const hash = [...label.toLocaleLowerCase()].reduce((value, character) => (value * 31 + character.charCodeAt(0)) >>> 0, 7);
  return groupColors[hash % groupColors.length];
}

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
}: ServiceItemOverviewProps) {
  const listMode = size <= 4;
  const [groupToRepeat, setGroupToRepeat] = useState(item.groups?.[0]?.id ?? "");
  const [arrangementName, setArrangementName] = useState("");
  const arrangements = item.arrangements ?? [];
  const activeArrangement = arrangements.find((entry) => entry.id === item.activeArrangementId) ?? arrangements[0] ?? null;
  const arrangedSlides = getItemSequenceSlides(item);
  const groupedItems = item.groups ?? [];

  function changeActiveArrangement(update: { groupIds: string[] }) {
    if (!activeArrangement) return;
    onItemChange({ arrangements: arrangements.map((arrangement) => arrangement.id === activeArrangement.id ? { ...arrangement, ...update } : arrangement) });
  }

  function createArrangement(event: SyntheticEvent<HTMLFormElement>) {
    event.preventDefault();
    const name = arrangementName.trim();
    if (!name || !activeArrangement) return;
    const created = { id: crypto.randomUUID(), name, groupIds: [...activeArrangement.groupIds] };
    onItemChange({ arrangements: [...arrangements, created], activeArrangementId: created.id });
    setArrangementName("");
  }

  return (
    <section className="item-overview" aria-label={`${item.title} overview`}>
      <header className="item-overview__heading">
        <div>
          <span className="eyebrow">{item.type.toUpperCase()} · SERVICE ITEM</span>
          <h2>{item.title}</h2>
          <span className="item-overview__count">{item.slides.length} {item.slides.length === 1 ? "slide" : "slides"}</span>
        </div>
        <button className="add-overview-slide" onClick={onAddSlide}>+ Add slide</button>
      </header>

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
            <select value={item.bibleVersion ?? "WEB"} onChange={(event) => onItemChange({ bibleVersion: event.currentTarget.value })}>
              {bibleVersions.map((version) => <option key={version}>{version}</option>)}
            </select>
          </label>
          <span className="bible-settings-note">Passage and version are labels only; text remains manually editable.</span>
        </div>
      )}

      {item.type === "song" && activeArrangement && groupedItems.length > 0 && (
        <section className="arrangement-editor" aria-label="Song arrangement">
          <div className="arrangement-editor__topline">
            <label>PLAY ORDER
              <select value={activeArrangement.id} onChange={(event) => onItemChange({ activeArrangementId: event.currentTarget.value })}>
                {arrangements.map((arrangement) => <option key={arrangement.id} value={arrangement.id}>{arrangement.name}</option>)}
              </select>
            </label>
            <form onSubmit={createArrangement}>
              <input value={arrangementName} onChange={(event) => setArrangementName(event.currentTarget.value)} placeholder="New arrangement name" aria-label="New arrangement name" />
              <button type="submit" disabled={!arrangementName.trim()}>Save copy</button>
            </form>
          </div>
          <div className="arrangement-editor__repeat">
            <label>REPEAT A SECTION
              <select value={groupToRepeat} onChange={(event) => setGroupToRepeat(event.currentTarget.value)}>
                {groupedItems.map((group) => <option key={group.id} value={group.id}>{group.name}</option>)}
              </select>
            </label>
            <button onClick={() => changeActiveArrangement({ groupIds: [...activeArrangement.groupIds, groupToRepeat] })} disabled={!groupToRepeat}>Add to play order</button>
          </div>
          <ol className="arrangement-editor__order">
            {activeArrangement.groupIds.map((groupId, index) => {
              const group = groupedItems.find((entry) => entry.id === groupId);
              return <li key={`${groupId}-${index}`} style={{ "--group-color": group?.color ?? "#87909c" } as CSSProperties}>
                <span>{group?.name ?? "Missing section"}</span>
                <button title="Move earlier" aria-label={`Move ${group?.name ?? "section"} earlier`} disabled={index === 0} onClick={() => {
                  const groupIds = [...activeArrangement.groupIds];
                  [groupIds[index - 1], groupIds[index]] = [groupIds[index], groupIds[index - 1]];
                  changeActiveArrangement({ groupIds });
                }}>?</button>
                <button title="Move later" aria-label={`Move ${group?.name ?? "section"} later`} disabled={index === activeArrangement.groupIds.length - 1} onClick={() => {
                  const groupIds = [...activeArrangement.groupIds];
                  [groupIds[index], groupIds[index + 1]] = [groupIds[index + 1], groupIds[index]];
                  changeActiveArrangement({ groupIds });
                }}>?</button>
                <button title="Remove from this arrangement" aria-label={`Remove ${group?.name ?? "section"} from play order`} onClick={() => changeActiveArrangement({ groupIds: activeArrangement.groupIds.filter((_, position) => position !== index) })}>?</button>
              </li>;
            })}
          </ol>
        </section>
      )}

      <div
        className={`item-overview__slides${listMode ? " item-overview__slides--list" : ""}`}
        style={{ "--overview-card-size": `${140 + size * 2}px` } as CSSProperties}
      >
        {arrangedSlides.map((slide, index) => (
          <article
            className={`overview-slide${slide.id === selectedSlideId ? " is-selected" : ""}${index + 1 === liveItemPosition ? " is-live" : ""}${index + 1 === nextItemPosition ? " is-next" : ""}`}
            key={`${slide.id}-${index}`}
            style={{ "--group-color": item.groups?.find((group) => group.slideIds.includes(slide.id))?.color ?? groupColor(slide.sectionLabel) } as CSSProperties}
          >
            <button
              className="overview-slide__select"
              aria-label={`Select slide ${index + 1}`}
              onClick={() => onSelectSlide(slide.id)}
              onDoubleClick={() => onGoLive(slide, index + 1)}
            >
              {!listMode && <SlideCanvas slide={slide} screenMode="slide" variant="overview" />}
              {(slide.cueMacroIds?.length ?? 0) > 0 && <span className="slide-cue-badge" title={`${slide.cueMacroIds?.length} automatic cue(s)`}>↯ {slide.cueMacroIds?.length}</span>}
              <span className="overview-slide__meta">
                <span className="overview-slide__number">{String(index + 1).padStart(2, "0")}</span>
                {slide.sectionLabel && <span className="overview-slide__section">{slide.sectionLabel}</span>}
                <span className="overview-slide__text">{slide.text.replace(/\s+/g, " ").trim() || "Empty slide"}</span>
              </span>
            </button>
            <button className="overview-slide__edit" onClick={() => onEditSlide(slide.id)}>Edit</button>
            <button className="overview-slide__go-live" onClick={() => onGoLive(slide, index + 1)} disabled={index + 1 === liveItemPosition}>
              {index + 1 === liveItemPosition ? "ON AIR" : "GO LIVE"}
            </button>
          </article>
        ))}
        {item.slides.length === 0 && <div className="overview-empty">This item does not have any slides yet.</div>}
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
