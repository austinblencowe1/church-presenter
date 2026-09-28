import { useEffect, useState } from "react";
import { emitTo, listen } from "@tauri-apps/api/event";
import { getCurrentWebviewWindow } from "@tauri-apps/api/webviewWindow";
import { SlideCanvas } from "./SlideCanvas";
import type { PresentationUpdate } from "../types/presentation";

export function PresentationWindow() {
  const [current, setCurrent] = useState<PresentationUpdate>({ liveSlide: null, screenMode: "slide" });

  useEffect(() => {
    let unlisten: (() => void) | undefined;
    let active = true;

    void listen<PresentationUpdate>("presentation-update", (event) => {
      setCurrent(event.payload);
    }).then((stopListening) => {
      if (!active) {
        stopListening();
        return;
      }
      unlisten = stopListening;
      void emitTo("main", "presentation-ready");
    });

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "ArrowRight" || event.key === " ") {
        event.preventDefault();
        void emitTo("main", "presentation-navigation", { direction: "next" });
      } else if (event.key === "ArrowLeft") {
        event.preventDefault();
        void emitTo("main", "presentation-navigation", { direction: "previous" });
      } else if (event.key === "Home") {
        event.preventDefault();
        void emitTo("main", "presentation-jump", { edge: "first" });
      } else if (event.key === "End") {
        event.preventDefault();
        void emitTo("main", "presentation-jump", { edge: "last" });
      } else if (event.key.toLowerCase() === "b") {
        void emitTo("main", "presentation-screen-mode", { mode: "black" });
      } else if (event.key.toLowerCase() === "w") {
        void emitTo("main", "presentation-screen-mode", { mode: "white" });
      } else if (event.key === "Escape") {
        void emitTo("main", "presentation-closed");
        void getCurrentWebviewWindow().close();
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => {
      active = false;
      unlisten?.();
      window.removeEventListener("keydown", onKeyDown);
    };
  }, []);

  return (
    <main className="presentation-stage" aria-label="Presentation output">
      {current.liveSlide
        ? <SlideCanvas slide={current.liveSlide.slide} screenMode={current.screenMode} variant="stage" />
        : <div className="presentation-stage__waiting" />}
    </main>
  );
}