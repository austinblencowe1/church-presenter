import { useEffect, useState } from "react";
import type { FormEvent } from "react";
import { SlideCanvas } from "./SlideCanvas";
import { getSlideSequence } from "../data/presentationNavigation";
import type {
  LiveSlide,
  ScreenMode,
  Service,
  ServiceItem,
  ServiceItemType,
  Slide,
  TextAlignment,
} from "../types/presentation";

export interface PresentationDisplay {
  id: string;
  label: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

interface ServiceEditorProps {
  service: Service;
  saveStatus: "saving" | "saved" | "error";
  isPresenting: boolean;
  liveSlide: LiveSlide | null;
  screenMode: ScreenMode;
  displays: PresentationDisplay[];
  selectedDisplay: string;
  onServiceChange: (service: Service) => void;
  onBack: () => void;
  onOpenPresentation: () => void;
  onClosePresentation: () => void;
  onGoLive: (slide: LiveSlide) => void;
  onNavigate: (direction: -1 | 1) => void;
  onJumpTo: (edge: "first" | "last") => void;
  onScreenMode: (mode: ScreenMode) => void;
  onDisplayChange: (displayId: string) => void;
  onForceSave: () => void;
}

const itemTypes: Array<{ value: ServiceItemType; label: string }> = [
  { value: "welcome", label: "Welcome" },
  { value: "song", label: "Song" },
  { value: "bible", label: "Bible Reading" },
  { value: "sermon", label: "Sermon" },
  { value: "announcement", label: "Announcement" },
  { value: "other", label: "Other" },
];

const alignmentGlyphs: Record<TextAlignment, string> = { left: "☰", center: "≡", right: "☷" };

function displayDate(value: string): string {
  return value
    ? new Intl.DateTimeFormat(undefined, { day: "numeric", month: "long", year: "numeric" }).format(new Date(`${value}T12:00:00`))
    : "Date not set";
}

function createSlide(): Slide {
  return {
    id: crypto.randomUUID(),
    type: "text",
    text: "New slide",
    background: "#17211f",
    textAlign: "center",
    fontSize: 56,
  };
}

function itemDefaultTitle(type: ServiceItemType): string {
  return itemTypes.find((item) => item.value === type)?.label ?? "Other";
}

export function ServiceEditor({
  service,
  saveStatus,
  isPresenting,
  liveSlide,
  screenMode,
  displays,
  selectedDisplay,
  onServiceChange,
  onBack,
  onOpenPresentation,
  onClosePresentation,
  onGoLive,
  onNavigate,
  onJumpTo,
  onScreenMode,
  onDisplayChange,
  onForceSave,
}: ServiceEditorProps) {
  const [selectedItemId, setSelectedItemId] = useState(service.items[0]?.id ?? "");
  const [selectedSlideId, setSelectedSlideId] = useState(service.items[0]?.slides[0]?.id ?? "");
  const [expandedItems, setExpandedItems] = useState<Set<string>>(() => new Set(service.items.map((item) => item.id)));
  const [newItemType, setNewItemType] = useState<ServiceItemType>("song");
  const [editingItemId, setEditingItemId] = useState("");
  const [editingItemTitle, setEditingItemTitle] = useState("");
  const sequence = getSlideSequence(service);
  const selectedItem = service.items.find((item) => item.id === selectedItemId) ?? null;
  const selectedSlide = selectedItem?.slides.find((slide) => slide.id === selectedSlideId) ?? null;
  const selectedItemSlideIndex = selectedItem?.slides.findIndex((slide) => slide.id === selectedSlideId) ?? -1;
  const previewSlide = selectedSlide && selectedItem
    ? sequence.find((step) => step.itemId === selectedItem.id && step.slide.id === selectedSlide.id) ?? null
    : null;
  const isPreviewLive = previewSlide?.slide.id === liveSlide?.slide.id && previewSlide?.serviceId === liveSlide?.serviceId;

  useEffect(() => {
    setSelectedItemId(service.items[0]?.id ?? "");
    setSelectedSlideId(service.items[0]?.slides[0]?.id ?? "");
    setExpandedItems(new Set(service.items.map((item) => item.id)));
  }, [service.id]);

  useEffect(() => {
    if (selectedItem && !selectedItem.slides.some((slide) => slide.id === selectedSlideId)) {
      setSelectedSlideId(selectedItem.slides[0]?.id ?? "");
    }
  }, [selectedItem, selectedSlideId]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const target = event.target;
      if (target instanceof HTMLElement && (target.isContentEditable || ["INPUT", "TEXTAREA", "SELECT"].includes(target.tagName))) return;
      if (target instanceof HTMLButtonElement && [" ", "Enter"].includes(event.key)) return;

      if (event.ctrlKey && event.key.toLowerCase() === "s") {
        event.preventDefault();
        onForceSave();
      } else if (event.key === "ArrowRight" || event.key === " ") {
        event.preventDefault();
        onNavigate(1);
      } else if (event.key === "ArrowLeft") {
        event.preventDefault();
        onNavigate(-1);
      } else if (event.key === "Home") {
        event.preventDefault();
        onJumpTo("first");
      } else if (event.key === "End") {
        event.preventDefault();
        onJumpTo("last");
      } else if (event.key === "Enter" && previewSlide) {
        onGoLive(previewSlide);
      } else if (event.key.toLowerCase() === "b") {
        onScreenMode("black");
      } else if (event.key.toLowerCase() === "w") {
        onScreenMode("white");
      } else if (event.key === "Escape" && isPresenting) {
        onClosePresentation();
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [isPresenting, onForceSave, onGoLive, onJumpTo, onNavigate, onScreenMode, onClosePresentation, previewSlide]);

  function updateService(update: Partial<Service>) {
    onServiceChange({ ...service, ...update });
  }

  function changeItem(itemId: string, update: Partial<ServiceItem>) {
    updateService({ items: service.items.map((item) => item.id === itemId ? { ...item, ...update } : item) });
  }

  function addItem() {
    const item: ServiceItem = {
      id: crypto.randomUUID(),
      title: itemDefaultTitle(newItemType),
      type: newItemType,
      slides: [createSlide()],
    };
    updateService({ items: [...service.items, item] });
    setSelectedItemId(item.id);
    setSelectedSlideId(item.slides[0].id);
    setExpandedItems((current) => new Set(current).add(item.id));
  }

  function removeItem(itemId: string) {
    const items = service.items.filter((item) => item.id !== itemId);
    updateService({ items });
    if (selectedItemId === itemId) {
      setSelectedItemId(items[0]?.id ?? "");
      setSelectedSlideId(items[0]?.slides[0]?.id ?? "");
    }
  }

  function moveItem(itemIndex: number, direction: -1 | 1) {
    const targetIndex = itemIndex + direction;
    if (targetIndex < 0 || targetIndex >= service.items.length) return;
    const items = [...service.items];
    [items[itemIndex], items[targetIndex]] = [items[targetIndex], items[itemIndex]];
    updateService({ items });
  }

  function addSlide(itemId: string) {
    const slide = createSlide();
    const item = service.items.find((entry) => entry.id === itemId);
    if (!item) return;
    changeItem(itemId, { slides: [...item.slides, slide] });
    setSelectedItemId(itemId);
    setSelectedSlideId(slide.id);
    setExpandedItems((current) => new Set(current).add(itemId));
  }

  function changeSlide(update: Partial<Slide>) {
    if (!selectedItem || !selectedSlide) return;
    changeItem(selectedItem.id, {
      slides: selectedItem.slides.map((slide) => slide.id === selectedSlide.id ? { ...slide, ...update } : slide),
    });
  }

  function removeSlide() {
    if (!selectedItem || !selectedSlide) return;
    const slides = selectedItem.slides.filter((slide) => slide.id !== selectedSlide.id);
    changeItem(selectedItem.id, { slides });
    setSelectedSlideId(slides[Math.min(selectedItemSlideIndex, slides.length - 1)]?.id ?? "");
  }

  function duplicateSlide() {
    if (!selectedItem || !selectedSlide) return;
    const copy = { ...selectedSlide, id: crypto.randomUUID() };
    const slides = [...selectedItem.slides];
    slides.splice(selectedItemSlideIndex + 1, 0, copy);
    changeItem(selectedItem.id, { slides });
    setSelectedSlideId(copy.id);
  }

  function moveSlide(direction: -1 | 1) {
    if (!selectedItem) return;
    const targetIndex = selectedItemSlideIndex + direction;
    if (selectedItemSlideIndex < 0 || targetIndex < 0 || targetIndex >= selectedItem.slides.length) return;
    const slides = [...selectedItem.slides];
    [slides[selectedItemSlideIndex], slides[targetIndex]] = [slides[targetIndex], slides[selectedItemSlideIndex]];
    changeItem(selectedItem.id, { slides });
  }

  function renameItem(event: FormEvent<HTMLFormElement>, itemId: string) {
    event.preventDefault();
    if (editingItemTitle.trim()) changeItem(itemId, { title: editingItemTitle.trim() });
    setEditingItemId("");
  }

  return (
    <main className="app-shell service-editor-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <button className="back-button" onClick={onBack} title="Back to services" aria-label="Back to services">←</button>
          <div className="brand-mark" aria-hidden="true">CP</div>
          <div className="brand-name">Church Presenter</div>
          <div className="topbar-divider" />
          <div className="service-heading">
            <strong>{service.title}</strong>
            <span>{displayDate(service.date)}</span>
          </div>
        </div>
        <div className="topbar-actions">
          <div className={`live-status${isPresenting ? " live-status--active" : ""}`}>
            <span className="live-status__dot" />
            {isPresenting ? "LIVE" : "OFFLINE"}
          </div>
          <span className={`save-status save-status--${saveStatus}`}>
            <span className="save-status__check">{saveStatus === "saving" ? "◌" : saveStatus === "error" ? "!" : "✓"}</span>
            {saveStatus === "saving" ? "Saving..." : saveStatus === "error" ? "Save failed" : "Saved"}
          </span>
          <label className="display-picker">
            <span>DISPLAY</span>
            <select value={selectedDisplay} onChange={(event) => onDisplayChange(event.currentTarget.value)}>
              <option value="automatic">Automatic</option>
              {displays.map((display) => <option value={display.id} key={display.id}>{display.label}</option>)}
            </select>
          </label>
          {isPresenting ? (
            <button className="present-button present-button--stop" onClick={onClosePresentation}>■ End</button>
          ) : (
            <button className="present-button" onClick={onOpenPresentation}><span className="present-button__icon">▶</span> Present</button>
          )}
        </div>
      </header>

      <div className="workspace service-workspace">
        <aside className="slide-sidebar service-sidebar">
          <div className="sidebar-heading service-sidebar-heading">
            <div>
              <span className="eyebrow">SERVICE ORDER</span>
              <h1>{service.items.length} {service.items.length === 1 ? "item" : "items"}</h1>
            </div>
            <span className="slide-count">{sequence.length}</span>
          </div>
          <div className="service-tree" aria-label="Service items and slides">
            {service.items.map((item, itemIndex) => {
              const expanded = expandedItems.has(item.id);
              const isLiveItem = item.id === liveSlide?.itemId;
              return (
                <section className={`service-item${isLiveItem ? " service-item--live" : ""}`} key={item.id}>
                  <div className={`service-item-heading${selectedItemId === item.id ? " is-selected" : ""}`}>
                    <button
                      className="item-expand"
                      title={expanded ? "Collapse item" : "Expand item"}
                      aria-label={`${expanded ? "Collapse" : "Expand"} ${item.title}`}
                      aria-expanded={expanded}
                      onClick={() => setExpandedItems((current) => {
                        const next = new Set(current);
                        if (next.has(item.id)) next.delete(item.id);
                        else next.add(item.id);
                        return next;
                      })}
                    >{expanded ? "⌄" : "›"}</button>
                    <button
                      className="item-title-button"
                      onClick={() => {
                        setSelectedItemId(item.id);
                        setSelectedSlideId(item.slides[0]?.id ?? "");
                      }}
                    >
                      <span className="item-type">{item.type.toUpperCase()}</span>
                      <strong>{item.title}</strong>
                    </button>
                    {isLiveItem && <span className="item-live-tag"><i /> LIVE</span>}
                    <details className="item-menu">
                      <summary aria-label={`Actions for ${item.title}`} title="Item actions">···</summary>
                      <div className="item-menu-options">
                        <button onClick={() => moveItem(itemIndex, -1)} disabled={itemIndex === 0}>Move up</button>
                        <button onClick={() => moveItem(itemIndex, 1)} disabled={itemIndex === service.items.length - 1}>Move down</button>
                        <button onClick={() => {
                          setEditingItemId(item.id);
                          setEditingItemTitle(item.title);
                        }}>Rename</button>
                        <button className="danger-action" onClick={() => removeItem(item.id)}>Delete</button>
                      </div>
                    </details>
                  </div>
                  {editingItemId === item.id && (
                    <form className="item-rename-form" onSubmit={(event) => renameItem(event, item.id)}>
                      <input autoFocus aria-label="Service item name" value={editingItemTitle} onChange={(event) => setEditingItemTitle(event.currentTarget.value)} />
                      <button type="submit" aria-label="Save name">✓</button>
                      <button type="button" aria-label="Cancel rename" onClick={() => setEditingItemId("")}>×</button>
                    </form>
                  )}
                  {expanded && (
                    <div className="service-item-slides">
                      {item.slides.map((slide, slideIndex) => {
                        const slideIsLive = liveSlide?.itemId === item.id && liveSlide.slide.id === slide.id;
                        return (
                          <button
                            key={slide.id}
                            className={`service-slide${slide.id === selectedSlideId && item.id === selectedItemId ? " is-preview" : ""}${slideIsLive ? " is-live" : ""}`}
                            onClick={() => { setSelectedItemId(item.id); setSelectedSlideId(slide.id); }}
                          >
                            <SlideCanvas slide={slide} screenMode="slide" variant="thumbnail" />
                            <span className="service-slide__index">{String(slideIndex + 1).padStart(2, "0")}</span>
                            <span className="service-slide__text">{slide.text.replace(/\s+/g, " ").trim() || "Empty slide"}</span>
                            {slideIsLive && <span className="service-slide__live-dot" title="Currently live" />}
                          </button>
                        );
                      })}
                      <button className="add-item-slide" onClick={() => addSlide(item.id)}>+ Add slide</button>
                    </div>
                  )}
                </section>
              );
            })}
            {service.items.length === 0 && <div className="empty-service-order">Add an item to start building this service.</div>}
          </div>
          <div className="add-item-control">
            <select aria-label="New item type" value={newItemType} onChange={(event) => setNewItemType(event.currentTarget.value as ServiceItemType)}>
              {itemTypes.map(({ value, label }) => <option value={value} key={value}>{label}</option>)}
            </select>
            <button className="add-slide-button" onClick={addItem}><span>+</span> Add item</button>
          </div>
        </aside>

        <section className="editor-pane" aria-label="Slide editor">
          <div className="pane-heading">
            <div>
              <span className="eyebrow">PREVIEW · {selectedItem?.title ?? "NO ITEM"}</span>
              <h2>{selectedSlide ? `Slide ${selectedItemSlideIndex + 1}` : "No slide selected"}</h2>
            </div>
            <span className="editor-hint">Changes stay in preview until Go Live</span>
          </div>
          {selectedSlide ? (
            <>
              <textarea
                className="slide-text-input"
                aria-label="Slide text"
                value={selectedSlide.text}
                onChange={(event) => changeSlide({ text: event.currentTarget.value })}
                placeholder="Enter slide text"
                spellCheck
              />
              <div className="format-toolbar">
                <div className="format-group">
                  <span className="control-label">ALIGN</span>
                  <div className="segmented-control" aria-label="Text alignment">
                    {(["left", "center", "right"] as const).map((alignment) => (
                      <button
                        key={alignment}
                        className={selectedSlide.textAlign === alignment ? "is-active" : ""}
                        title={`${alignment[0].toUpperCase()}${alignment.slice(1)} align`}
                        aria-label={`${alignment[0].toUpperCase()}${alignment.slice(1)} align`}
                        aria-pressed={selectedSlide.textAlign === alignment}
                        onClick={() => changeSlide({ textAlign: alignment })}
                      >{alignmentGlyphs[alignment]}</button>
                    ))}
                  </div>
                </div>
                <div className="format-group font-size-control">
                  <label className="control-label" htmlFor="font-size">TEXT SIZE <strong>{selectedSlide.fontSize}px</strong></label>
                  <input id="font-size" type="range" min="24" max="96" step="2" value={selectedSlide.fontSize}
                    onChange={(event) => changeSlide({ fontSize: Number(event.currentTarget.value) })} />
                </div>
                <div className="format-group background-control">
                  <label className="control-label" htmlFor="background-color">BACKGROUND</label>
                  <label className="color-picker" htmlFor="background-color" title="Choose background color">
                    <input id="background-color" type="color" value={selectedSlide.background}
                      onChange={(event) => changeSlide({ background: event.currentTarget.value })} />
                    <span>{selectedSlide.background.toUpperCase()}</span>
                  </label>
                </div>
              </div>
              <div className="slide-edit-actions">
                <span>{selectedItem?.title} · {selectedItemSlideIndex + 1} / {selectedItem?.slides.length}</span>
                <button onClick={() => moveSlide(-1)} disabled={selectedItemSlideIndex <= 0}>↑ Move up</button>
                <button onClick={() => moveSlide(1)} disabled={!selectedItem || selectedItemSlideIndex >= selectedItem.slides.length - 1}>↓ Move down</button>
                <button onClick={duplicateSlide}>▣ Duplicate</button>
                <button className="danger-action" onClick={removeSlide}>× Delete</button>
              </div>
            </>
          ) : (
            <div className="empty-editor-wrap">
              <div className="empty-editor-copy">{selectedItem ? "This item has no slides." : "Choose a service item to edit."}</div>
              {selectedItem && <button className="add-slide-button" onClick={() => addSlide(selectedItem.id)}><span>+</span> Add slide</button>}
            </div>
          )}
        </section>

        <aside className="preview-pane" aria-label="Slide preview and live output">
          <div className="pane-heading preview-heading">
            <div><span className="eyebrow">PREVIEW</span><h2>{selectedItem?.title ?? "No item"}</h2></div>
            <span className={`output-state${isPresenting ? " is-live" : ""}`}><i />{isPresenting ? "LIVE" : "OFFLINE"}</span>
          </div>
          <div className="preview-frame">
            {selectedSlide
              ? <SlideCanvas slide={selectedSlide} screenMode="slide" variant="preview" />
              : <div className="preview-empty">No slide selected</div>}
          </div>
          <div className="preview-detail">
            <span>PREVIEW SLIDE</span>
            <strong>{previewSlide ? `${previewSlide.itemTitle} · ${previewSlide.itemSlidePosition} / ${previewSlide.itemSlideTotal}` : "No slide selected"}</strong>
            <span>{previewSlide ? `Overall ${previewSlide.overallPosition} / ${previewSlide.overallTotal}` : "Add a slide to this item"}</span>
          </div>
          <div className="live-preview-block">
            <div className="live-preview-heading"><span><i /> LIVE OUTPUT</span><span>{isPresenting ? "CONNECTED" : "OFFLINE"}</span></div>
            <div className="live-preview-detail">
              <strong>{liveSlide?.itemTitle ?? "No live slide"}</strong>
              <span>{liveSlide ? `${liveSlide.itemSlidePosition} / ${liveSlide.itemSlideTotal} · Overall ${liveSlide.overallPosition} / ${liveSlide.overallTotal}` : "Select a preview slide and go live"}</span>
            </div>
            <button className="go-live-button" disabled={!previewSlide || isPreviewLive} onClick={() => previewSlide && onGoLive(previewSlide)}>
              {isPreviewLive ? "LIVE SLIDE" : "GO LIVE"}
            </button>
          </div>
          <div className="preview-footer">
            <span className="preview-resolution">16:9</span>
            <span className="preview-output"><span className={isPresenting ? "is-live" : ""} />{isPresenting ? "Output connected" : "Output idle"}</span>
          </div>
        </aside>
      </div>

      <footer className="bottom-bar service-bottom-bar">
        <div className="live-item-summary">
          <span>{liveSlide?.itemType.toUpperCase() ?? "LIVE"}</span>
          <strong>{liveSlide?.itemTitle ?? "No live slide"}</strong>
        </div>
        <button className="navigation-button" onClick={() => onNavigate(-1)} disabled={!liveSlide || liveSlide.overallPosition <= 1}>
          <span aria-hidden="true">←</span> Previous
        </button>
        <div className="slide-position"><strong>{liveSlide ? String(liveSlide.overallPosition).padStart(2, "0") : "--"}<span> / </span>{String(liveSlide?.overallTotal ?? sequence.length).padStart(2, "0")}</strong></div>
        <button className="navigation-button" onClick={() => onNavigate(1)} disabled={!liveSlide || liveSlide.overallPosition >= liveSlide.overallTotal}>
          Next <span aria-hidden="true">→</span>
        </button>
        <div className="screen-mode-controls">
          <button className={screenMode === "black" ? "is-active" : ""} onClick={() => onScreenMode(screenMode === "black" ? "slide" : "black")}>BLACK</button>
          <button className={screenMode === "white" ? "is-active" : ""} onClick={() => onScreenMode(screenMode === "white" ? "slide" : "white")}>WHITE</button>
        </div>
      </footer>
    </main>
  );
}