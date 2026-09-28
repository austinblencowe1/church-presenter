import { useEffect, useState } from "react";
import { isTauri } from "@tauri-apps/api/core";
import { emitTo, listen } from "@tauri-apps/api/event";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";
import { SlideCanvas } from "./components/SlideCanvas";
import { PresentationWindow } from "./components/PresentationWindow";
import { defaultPresentation } from "./data/defaultPresentation";
import type { Presentation, ScreenMode, Slide, TextAlignment } from "./types/presentation";
import "./App.css";

const presentationWindow = new URLSearchParams(window.location.search).get("mode") === "presentation";
const alignmentGlyphs: Record<TextAlignment, string> = { left: "☰", center: "≡", right: "☷" };

function App() {
  return presentationWindow ? <PresentationWindow /> : <EditorApp />;
}

function EditorApp() {
  const [presentation, setPresentation] = useState<Presentation>(defaultPresentation);
  const [selectedSlideId, setSelectedSlideId] = useState(defaultPresentation.slides[0]?.id ?? "");
  const [screenMode, setScreenMode] = useState<ScreenMode>("slide");
  const [isPresenting, setIsPresenting] = useState(false);
  const selectedIndex = presentation.slides.findIndex((slide) => slide.id === selectedSlideId);
  const selectedSlide = presentation.slides[selectedIndex] ?? null;

  useEffect(() => {
    if (!isTauri()) return;

    let active = true;
    const unlisteners: Array<() => void> = [];
    const register = async () => {
      const registrations = await Promise.all([
        listen("presentation-ready", () => {
          void emitTo("presentation", "presentation-update", { slide: selectedSlide, screenMode });
        }),
        listen<{ direction: "next" | "previous" }>("presentation-navigation", (event) => {
          if (event.payload.direction === "next") moveSelection(1);
          else moveSelection(-1);
        }),
        listen<{ mode: ScreenMode }>("presentation-screen-mode", (event) => {
          setScreenMode(event.payload.mode);
        }),
        listen("presentation-closed", () => setIsPresenting(false)),
      ]);

      if (active) unlisteners.push(...registrations);
      else registrations.forEach((unlisten) => unlisten());
    };

    void register();
    return () => {
      active = false;
      unlisteners.forEach((unlisten) => unlisten());
    };
  }, [selectedSlide, screenMode]);

  useEffect(() => {
    if (isTauri()) {
      void emitTo("presentation", "presentation-update", { slide: selectedSlide, screenMode }).catch(() => {});
    }
  }, [selectedSlide, screenMode]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const target = event.target;
      if (
        target instanceof HTMLElement &&
        (target.isContentEditable || ["INPUT", "TEXTAREA", "SELECT"].includes(target.tagName))
      ) return;
      if (event.key === " " && target instanceof HTMLButtonElement) return;

      if (event.key === "ArrowRight" || event.key === " ") {
        event.preventDefault();
        moveSelection(1);
      } else if (event.key === "ArrowLeft") {
        event.preventDefault();
        moveSelection(-1);
      } else if (event.key.toLowerCase() === "b") {
        setScreenMode("black");
      } else if (event.key.toLowerCase() === "w") {
        setScreenMode("white");
      } else if (event.key === "Escape" && isPresenting) {
        void closePresentationWindow();
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [presentation.slides, selectedIndex, isPresenting]);

  function moveSelection(direction: -1 | 1) {
    if (presentation.slides.length === 0) return;
    const nextIndex = Math.min(Math.max(selectedIndex + direction, 0), presentation.slides.length - 1);
    setSelectedSlideId(presentation.slides[nextIndex].id);
    setScreenMode("slide");
  }

  function updateSelectedSlide(update: Partial<Slide>) {
    if (!selectedSlide) return;
    setPresentation((current) => ({
      ...current,
      slides: current.slides.map((slide) => slide.id === selectedSlide.id ? { ...slide, ...update } : slide),
    }));
  }

  function addSlide() {
    const slide = createSlide();
    setPresentation((current) => ({ ...current, slides: [...current.slides, slide] }));
    setSelectedSlideId(slide.id);
    setScreenMode("slide");
  }

  function duplicateSlide() {
    if (!selectedSlide) return;
    const copy = { ...selectedSlide, id: crypto.randomUUID() };
    setPresentation((current) => {
      const slides = [...current.slides];
      slides.splice(selectedIndex + 1, 0, copy);
      return { ...current, slides };
    });
    setSelectedSlideId(copy.id);
  }

  function deleteSlide() {
    if (!selectedSlide) return;
    const slides = presentation.slides.filter((slide) => slide.id !== selectedSlide.id);
    setPresentation((current) => ({ ...current, slides }));
    setSelectedSlideId(slides[Math.min(selectedIndex, slides.length - 1)]?.id ?? "");
  }

  function reorderSlide(direction: -1 | 1) {
    const targetIndex = selectedIndex + direction;
    if (selectedIndex < 0 || targetIndex < 0 || targetIndex >= presentation.slides.length) return;
    setPresentation((current) => {
      const slides = [...current.slides];
      [slides[selectedIndex], slides[targetIndex]] = [slides[targetIndex], slides[selectedIndex]];
      return { ...current, slides };
    });
  }

  async function openPresentationWindow() {
    if (!isTauri()) return;
    const existing = await WebviewWindow.getByLabel("presentation");
    if (existing) {
      setIsPresenting(true);
      await existing.setFocus();
      return;
    }

    const stage = new WebviewWindow("presentation", {
      url: "index.html?mode=presentation",
      title: "Church Presenter - Live",
      fullscreen: true,
      decorations: false,
      resizable: false,
    });
    stage.once("tauri://created", () => setIsPresenting(true));
    stage.once("tauri://error", (event) => {
      console.error("Could not open the presentation window", event.payload);
      setIsPresenting(false);
    });
  }

  async function closePresentationWindow() {
    if (!isTauri()) return;
    const stage = await WebviewWindow.getByLabel("presentation");
    if (stage) await stage.close();
    setIsPresenting(false);
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <div className="brand-mark" aria-hidden="true">CP</div>
          <div className="brand-name">Church Presenter</div>
          <div className="topbar-divider" />
          <input
            className="presentation-title"
            aria-label="Presentation title"
            value={presentation.title}
            onChange={(event) => setPresentation((current) => ({ ...current, title: event.currentTarget.value }))}
          />
        </div>
        <div className="topbar-actions">
          <div className={`live-status${isPresenting ? " live-status--active" : ""}`}>
            <span className="live-status__dot" />
            {isPresenting ? "LIVE" : "NOT LIVE"}
          </div>
          <span className="save-status"><span className="save-status__check">✓</span> Saved</span>
          <button className="present-button" onClick={() => void openPresentationWindow()}>
            <span className="present-button__icon">▶</span> Present
          </button>
        </div>
      </header>

      <div className="workspace">
        <aside className="slide-sidebar">
          <div className="sidebar-heading">
            <div>
              <span className="eyebrow">SERVICE</span>
              <h1>Presentation</h1>
            </div>
            <span className="slide-count">{presentation.slides.length}</span>
          </div>
          <div className="slide-list" aria-label="Slides">
            {presentation.slides.map((slide, index) => (
              <div
                key={slide.id}
                className={`slide-row${slide.id === selectedSlideId ? " slide-row--selected" : ""}`}
              >
                <button
                  className="slide-row__select"
                  aria-label={`Select slide ${index + 1}`}
                  onClick={() => { setSelectedSlideId(slide.id); setScreenMode("slide"); }}
                >
                  <SlideCanvas slide={slide} screenMode="slide" variant="thumbnail" />
                  <span className="slide-row__details">
                    <span className="slide-row__number">{String(index + 1).padStart(2, "0")}</span>
                    <span className="slide-row__text">{slide.text.replace(/\s+/g, " ").trim() || "Empty slide"}</span>
                  </span>
                </button>
                {slide.id === selectedSlideId && (
                  <div className="slide-row__actions">
                    <button title="Move up" aria-label="Move slide up" disabled={index === 0} onClick={() => reorderSlide(-1)}>↑</button>
                    <button title="Move down" aria-label="Move slide down" disabled={index === presentation.slides.length - 1} onClick={() => reorderSlide(1)}>↓</button>
                    <button title="Duplicate" aria-label="Duplicate slide" onClick={duplicateSlide}>▣</button>
                    <button title="Delete" aria-label="Delete slide" onClick={deleteSlide}>×</button>
                  </div>
                )}
              </div>
            ))}
          </div>
          <button className="add-slide-button" onClick={addSlide}><span>+</span> Add slide</button>
        </aside>

        <section className="editor-pane" aria-label="Slide editor">
          <div className="pane-heading">
            <div>
              <span className="eyebrow">SLIDE EDITOR</span>
              <h2>{selectedSlide ? `Slide ${selectedIndex + 1}` : "No slide selected"}</h2>
            </div>
            <span className="editor-hint">Changes update the preview instantly</span>
          </div>
          {selectedSlide ? (
            <>
              <textarea
                className="slide-text-input"
                aria-label="Slide text"
                value={selectedSlide.text}
                onChange={(event) => updateSelectedSlide({ text: event.currentTarget.value })}
                placeholder="Enter slide text"
                spellCheck
              />
              <div className="format-toolbar">
                <div className="format-group">
                  <span className="control-label">ALIGN</span>
                  <div className="segmented-control">
                    {(["left", "center", "right"] as const).map((alignment) => (
                      <button
                        key={alignment}
                        className={selectedSlide.textAlign === alignment ? "is-active" : ""}
                        title={`${alignment[0].toUpperCase()}${alignment.slice(1)} align`}
                        aria-label={`${alignment[0].toUpperCase()}${alignment.slice(1)} align`}
                        aria-pressed={selectedSlide.textAlign === alignment}
                        onClick={() => updateSelectedSlide({ textAlign: alignment })}
                      >
                        {alignmentGlyphs[alignment]}
                      </button>
                    ))}
                  </div>
                </div>
                <div className="format-group font-size-control">
                  <label className="control-label" htmlFor="font-size">TEXT SIZE <strong>{selectedSlide.fontSize}px</strong></label>
                  <input
                    id="font-size"
                    type="range"
                    min="24"
                    max="96"
                    step="2"
                    value={selectedSlide.fontSize}
                    onChange={(event) => updateSelectedSlide({ fontSize: Number(event.currentTarget.value) })}
                  />
                </div>
                <div className="format-group background-control">
                  <label className="control-label" htmlFor="background-color">BACKGROUND</label>
                  <label className="color-picker" htmlFor="background-color" title="Choose background color">
                    <input
                      id="background-color"
                      type="color"
                      value={selectedSlide.background}
                      onChange={(event) => updateSelectedSlide({ background: event.currentTarget.value })}
                    />
                    <span>{selectedSlide.background.toUpperCase()}</span>
                  </label>
                </div>
              </div>
            </>
          ) : (
            <button className="empty-editor" onClick={addSlide}>+ Add your first slide</button>
          )}
        </section>

        <aside className="preview-pane" aria-label="Slide preview">
          <div className="pane-heading preview-heading">
            <div>
              <span className="eyebrow">OUTPUT</span>
              <h2>Preview</h2>
            </div>
            {isPresenting && <span className="preview-live"><span /> LIVE</span>}
          </div>
          <div className="preview-frame">
            {selectedSlide
              ? <SlideCanvas slide={selectedSlide} screenMode={screenMode} variant="preview" />
              : <div className="preview-empty">No slide selected</div>}
          </div>
          <div className="preview-footer">
            <span className="preview-resolution">16:9</span>
            <span className="preview-output"><span className={isPresenting ? "is-live" : ""} />{isPresenting ? "Output connected" : "Output idle"}</span>
          </div>
        </aside>
      </div>

      <footer className="bottom-bar">
        <button className="navigation-button" onClick={() => moveSelection(-1)} disabled={selectedIndex <= 0}>
          <span aria-hidden="true">←</span> Previous
        </button>
        <div className="slide-position">
          <strong>{selectedSlide ? String(selectedIndex + 1).padStart(2, "0") : "--"}</strong>
          <span>/</span>
          <span>{String(presentation.slides.length).padStart(2, "0")}</span>
        </div>
        <button className="navigation-button" onClick={() => moveSelection(1)} disabled={selectedIndex < 0 || selectedIndex >= presentation.slides.length - 1}>
          Next <span aria-hidden="true">→</span>
        </button>
        <div className="keyboard-hint"><kbd>←</kbd><kbd>→</kbd> Navigate <span>·</span> <kbd>Space</kbd> Next</div>
      </footer>
    </main>
  );
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

export default App;
