import { useEffect, useRef, useState } from "react";
import { emitTo, listen } from "@tauri-apps/api/event";
import { getCurrentWebviewWindow } from "@tauri-apps/api/webviewWindow";
import { SlideCanvas } from "./SlideCanvas";
import type { PresentationUpdate } from "../types/presentation";

export function PresentationWindow() {
  const [current, setCurrent] = useState<PresentationUpdate>({ liveSlide: null, liveMedia: null, videoPlayback: { playing: false, volume: 1, playbackRate: 1, loop: false, seekTo: 0, seekRequest: 0, endBehavior: "stop" }, screenMode: "slide", clearForeground: false, clearBackground: false, message: null });
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const lastSeekRequest = useRef(-1);

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

  useEffect(() => {
    const video = videoRef.current;
    if (!video || current.liveMedia?.kind !== "video") return;
    video.volume = Math.max(0, Math.min(1, current.videoPlayback.volume));
    video.playbackRate = current.videoPlayback.playbackRate;
    video.loop = current.videoPlayback.loop;
    if (lastSeekRequest.current !== current.videoPlayback.seekRequest && video.readyState > 0) {
      video.currentTime = Math.max(0, current.videoPlayback.seekTo);
      lastSeekRequest.current = current.videoPlayback.seekRequest;
    }
    if (current.videoPlayback.playing) void video.play().catch(() => undefined);
    else video.pause();
  }, [current.liveMedia?.id, current.videoPlayback, current.screenMode, current.clearForeground]);

  return (
    <main className="presentation-stage" aria-label="Presentation output">
      {current.liveSlide
        ? <SlideCanvas slide={current.liveSlide.slide} screenMode={current.screenMode} variant="stage" clearSlide={current.clearForeground} clearBackground={current.clearBackground} />
        : <div className="presentation-stage__waiting" />}
      {current.liveMedia?.kind === "video" && <video
        ref={videoRef}
        className={`presentation-live-media${current.screenMode !== "slide" || current.clearForeground ? " is-hidden" : ""}`}
        src={current.liveMedia.source}
        autoPlay
        playsInline
        onLoadedMetadata={(event) => {
          if (lastSeekRequest.current !== current.videoPlayback.seekRequest) {
            event.currentTarget.currentTime = Math.max(0, current.videoPlayback.seekTo);
            lastSeekRequest.current = current.videoPlayback.seekRequest;
          }
          void emitTo("main", "presentation-video-progress", { currentTime: event.currentTarget.currentTime, duration: event.currentTarget.duration }).catch(() => undefined);
        }}
        onTimeUpdate={(event) => void emitTo("main", "presentation-video-progress", { currentTime: event.currentTarget.currentTime, duration: event.currentTarget.duration }).catch(() => undefined)}
        onEnded={() => void emitTo("main", "presentation-video-ended", { mediaId: current.liveMedia?.id }).catch(() => undefined)}
      />}
      {current.message && !current.clearForeground && <div className="presentation-message">{current.message}</div>}
    </main>
  );
}
