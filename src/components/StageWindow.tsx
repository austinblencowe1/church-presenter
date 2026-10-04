import { useEffect, useState } from "react";
import { PhysicalPosition, PhysicalSize } from "@tauri-apps/api/dpi";
import { emitTo, listen } from "@tauri-apps/api/event";
import { getCurrentWebviewWindow } from "@tauri-apps/api/webviewWindow";
import { availableMonitors } from "@tauri-apps/api/window";
import { SlideCanvas } from "./SlideCanvas";
import type { StagePresentationUpdate } from "../types/presentation";

const emptyStage: StagePresentationUpdate = { current: null, next: null, message: null, timer: null, audienceDisplayId: null };

interface StageMonitor {
  id: string;
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

export function StageWindow() {
  const [stage, setStage] = useState(emptyStage);
  const [now, setNow] = useState(Date.now());
  const [monitors, setMonitors] = useState<StageMonitor[]>([]);
  const [selectedMonitor, setSelectedMonitor] = useState("windowed");
  const [routingError, setRoutingError] = useState("");

  useEffect(() => {
    let active = true;
    let unlisten: (() => void) | undefined;
    void listen<StagePresentationUpdate>("stage-update", (event) => setStage(event.payload)).then((stopListening) => {
      if (!active) {
        stopListening();
        return;
      }
      unlisten = stopListening;
      void emitTo("main", "stage-ready");
    });
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        void emitTo("main", "stage-closed");
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
    void availableMonitors().then((available) => {
      setMonitors(available.map((monitor, index) => ({
        id: String(index),
        name: monitor.name || `Display ${index + 1}`,
        x: monitor.position.x,
        y: monitor.position.y,
        width: monitor.size.width,
        height: monitor.size.height,
      })));
    }).catch(() => setMonitors([]));
  }, []);

  useEffect(() => {
    if (selectedMonitor !== "windowed" && selectedMonitor === stage.audienceDisplayId) {
      void routeToMonitor("windowed").then(() => setRoutingError("The audience output took this display. Stage output returned to a movable window."));
    }
  }, [selectedMonitor, stage.audienceDisplayId]);

  async function routeToMonitor(value: string) {
    setSelectedMonitor(value);
    const stageWindow = getCurrentWebviewWindow();
    try {
      if (value === "windowed") {
        await stageWindow.setFullscreen(false);
        await stageWindow.setResizable(true);
        setRoutingError("");
        return;
      }
      const monitor = monitors.find((candidate) => candidate.id === value);
      if (!monitor) return;
      if (monitor.id === stage.audienceDisplayId) {
        setRoutingError("The audience output is already using this display.");
        return;
      }
      await stageWindow.setFullscreen(false);
      await stageWindow.setPosition(new PhysicalPosition(monitor.x, monitor.y));
      await stageWindow.setSize(new PhysicalSize(monitor.width, monitor.height));
      await stageWindow.setFullscreen(true);
      setRoutingError("");
    } catch (error) {
      setSelectedMonitor("windowed");
      setRoutingError(error instanceof Error ? error.message : String(error));
    }
  }

  useEffect(() => {
    if (!stage.timer) return;
    const endsAt = stage.timer.endsAt;
    const tick = window.setInterval(() => {
      const currentTime = Date.now();
      setNow(currentTime);
      if (currentTime >= endsAt) window.clearInterval(tick);
    }, 250);
    return () => window.clearInterval(tick);
  }, [stage.timer?.endsAt]);

  const secondsRemaining = stage.timer ? Math.max(0, Math.ceil((stage.timer.endsAt - now) / 1000)) : 0;
  const countdown = `${String(Math.floor(secondsRemaining / 60)).padStart(2, "0")}:${String(secondsRemaining % 60).padStart(2, "0")}`;

  return (
    <main className="stage-window" aria-label="Stage display">
      <header className="stage-window__header">
        <strong>STAGE DISPLAY</strong>
        <span><i /> {stage.current ? "LIVE" : "WAITING FOR OUTPUT"}</span>
        <span>{stage.current?.serviceTitle ?? "Church Presenter"}</span>
        <label className="stage-window__monitor-picker">
          <span>OUTPUT</span>
          <select value={selectedMonitor} onChange={(event) => void routeToMonitor(event.currentTarget.value)}>
            <option value="windowed">Windowed</option>
            {monitors.map((monitor) => <option key={monitor.id} value={monitor.id} disabled={monitor.id === stage.audienceDisplayId}>{monitor.name} · {monitor.width} × {monitor.height}{monitor.id === stage.audienceDisplayId ? " · Audience" : ""}</option>)}
          </select>
        </label>
      </header>
      {routingError && <div className="stage-window__routing-error" role="alert">Could not route output: {routingError}</div>}
      <div className="stage-window__body">
        <section className="stage-window__current">
          <div className="stage-window__label"><i /> CURRENT</div>
          <div className="stage-window__canvas">
            {stage.current
              ? <SlideCanvas slide={stage.current.slide} screenMode="slide" variant="stage" />
              : <div className="stage-window__empty">Current slide will appear here</div>}
          </div>
          <div className="stage-window__notes">
            <span>STAGE NOTES</span>
            <p>{stage.current?.slide.notes || "No speaker notes for this slide."}</p>
          </div>
          <footer>
            <strong>{stage.current?.itemTitle ?? "No item"}</strong>
            <span>{stage.current ? `${stage.current.itemSlidePosition} / ${stage.current.itemSlideTotal} · Overall ${stage.current.overallPosition} / ${stage.current.overallTotal}` : ""}</span>
          </footer>
        </section>
        <section className="stage-window__next">
          <div className="stage-window__label">UP NEXT</div>
          <div className="stage-window__canvas">
            {stage.next
              ? <SlideCanvas slide={stage.next.slide} screenMode="slide" variant="stage" />
              : <div className="stage-window__empty">End of service</div>}
          </div>
          <div className="stage-window__next-lyrics">
            <span>NEXT LYRICS</span>
            <p>{stage.next?.slide.text ?? "The service has no more slides."}</p>
          </div>
          <footer>
            <strong>{stage.next?.slide.sectionLabel || stage.next?.itemTitle || "Nothing queued"}</strong>
            <span>{stage.next?.itemTitle ?? ""}</span>
          </footer>
        </section>
      </div>
      <div className="stage-window__status">
        <span>{stage.message ?? "No message overlay"}</span>
        <strong className={secondsRemaining === 0 && stage.timer ? "is-finished" : ""}>
          {stage.timer ? <>{stage.timer.label} <time>{countdown}</time></> : "TIMER IDLE"}
        </strong>
      </div>
    </main>
  );
}
