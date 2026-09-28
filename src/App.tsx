import { useEffect, useRef, useState } from "react";
import { isTauri } from "@tauri-apps/api/core";
import { emitTo, listen } from "@tauri-apps/api/event";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";
import { availableMonitors, currentMonitor } from "@tauri-apps/api/window";
import { PresentationWindow } from "./components/PresentationWindow";
import { ServiceEditor, type PresentationDisplay } from "./components/ServiceEditor";
import { ServiceHome } from "./components/ServiceHome";
import { getSlideSequence } from "./data/presentationNavigation";
import { deleteService, listServices, loadService, saveService } from "./data/serviceStore";
import type { LiveSlide, PresentationUpdate, ScreenMode, Service, ServiceSummary } from "./types/presentation";
import "./App.css";

const isPresentationWindow = new URLSearchParams(window.location.search).get("mode") === "presentation";

function App() {
  return isPresentationWindow ? <PresentationWindow /> : <ChurchPresenter />;
}

function ChurchPresenter() {
  const [services, setServices] = useState<ServiceSummary[]>([]);
  const [currentService, setCurrentService] = useState<Service | null>(null);
  const [isLoaded, setIsLoaded] = useState(false);
  const [saveStatus, setSaveStatus] = useState<"saving" | "saved" | "error">("saved");
  const [homeError, setHomeError] = useState("");
  const [liveSlide, setLiveSlide] = useState<LiveSlide | null>(null);
  const [screenMode, setScreenMode] = useState<ScreenMode>("slide");
  const [isPresenting, setIsPresenting] = useState(false);
  const [displays, setDisplays] = useState<PresentationDisplay[]>([]);
  const [selectedDisplay, setSelectedDisplay] = useState("automatic");
  const updateRef = useRef<PresentationUpdate>({ liveSlide: null, screenMode: "slide" });
  const serviceRef = useRef<Service | null>(null);
  const navigateRef = useRef<(direction: -1 | 1) => void>(() => undefined);

  useEffect(() => {
    updateRef.current = { liveSlide, screenMode };
    serviceRef.current = currentService;
  }, [currentService, liveSlide, screenMode]);

  useEffect(() => {
    let active = true;
    void (async () => {
      try {
        const summaries = await listServices();
        if (active) setServices(summaries);
      } catch (error) {
        if (active) setHomeError(errorMessage(error));
      } finally {
        if (active) setIsLoaded(true);
      }
    })();
    return () => { active = false; };
  }, []);

  useEffect(() => {
    if (!currentService || !isLoaded) return;
    const timer = window.setTimeout(() => {
      void saveService(currentService).then(() => {
        setServices((current) => upsertSummary(current, currentService));
        setSaveStatus("saved");
      }).catch(() => setSaveStatus("error"));
    }, 450);
    return () => window.clearTimeout(timer);
  }, [currentService, isLoaded]);

  useEffect(() => {
    if (!isTauri()) return;
    void availableMonitors().then((monitors) => {
      setDisplays(monitors.map((monitor, index) => {
        const position = monitor.position.toLogical(monitor.scaleFactor);
        const size = monitor.size.toLogical(monitor.scaleFactor);
        const displayName = monitor.name || `Display ${index + 1}`;
        return {
          id: String(index),
          label: `${displayName} · ${monitor.size.width} × ${monitor.size.height}`,
          x: position.x,
          y: position.y,
          width: size.width,
          height: size.height,
        };
      }));
    }).catch(() => setDisplays([]));
  }, []);

  useEffect(() => {
    if (!isTauri()) return;
    let active = true;
    const unlisteners: Array<() => void> = [];
    void Promise.all([
      listen("presentation-ready", () => {
        void emitTo("presentation", "presentation-update", updateRef.current);
      }),
      listen<{ direction: "next" | "previous" }>("presentation-navigation", (event) => {
        const service = serviceRef.current;
        const sequence = service ? getSlideSequence(service) : [];
      }),
      listen<{ edge: "first" | "last" }>("presentation-jump", (event) => {
        const sequence = currentService ? getSlideSequence(currentService) : [];
        const target = event.payload.edge === "first" ? sequence[0] : sequence[sequence.length - 1];
        if (target) setLiveSlide(target);
        setScreenMode("slide");
      }),
      listen<{ mode: ScreenMode }>("presentation-screen-mode", (event) => setScreenMode(event.payload.mode)),
      listen("presentation-closed", () => setIsPresenting(false)),
    ]).then((registrations) => {
      if (active) unlisteners.push(...registrations);
      else registrations.forEach((unlisten) => unlisten());
    });
    return () => {
      active = false;
      unlisteners.forEach((unlisten) => unlisten());
    };
  }, []);

  useEffect(() => {
    navigateRef.current = (direction) => {
      if (!currentService) return;
      const sequence = getSlideSequence(currentService);
      if (sequence.length === 0) return;
      const currentIndex = sequence.findIndex((entry) => entry.itemId === liveSlide?.itemId && entry.slide.id === liveSlide.slide.id);
      const targetIndex = Math.min(Math.max(currentIndex + direction, 0), sequence.length - 1);
      setLiveSlide(sequence[targetIndex]);
      setScreenMode("slide");
    };
  }, [currentService, liveSlide]);

  useEffect(() => {
    if (isTauri()) {
      void emitTo("presentation", "presentation-update", updateRef.current).catch(() => undefined);
    }
  }, [liveSlide, screenMode]);

  async function createService(service: Service) {
    try {
      await saveService(service);
      setServices((current) => upsertSummary(current, service));
      setCurrentService(service);
      setLiveSlide(null);
      setScreenMode("slide");
      setSaveStatus("saved");
      setHomeError("");
    } catch (error) {
      setHomeError(errorMessage(error));
    }
  }

  async function openService(serviceId: string) {
    try {
      setCurrentService(await loadService(serviceId));
      setLiveSlide(null);
      setScreenMode("slide");
      setSaveStatus("saved");
      setHomeError("");
    } catch (error) {
      setHomeError(errorMessage(error));
    }
  }

  async function renameService(serviceId: string, title: string) {
    try {
      const service = await loadService(serviceId);
      await saveService({ ...service, title });
      setServices((current) => current.map((item) => item.id === serviceId ? { ...item, title } : item));
    } catch (error) {
      setHomeError(errorMessage(error));
    }
  }

  async function duplicateService(serviceId: string) {
    try {
      const original = await loadService(serviceId);
      const copy: Service = {
        ...original,
        id: crypto.randomUUID(),
        title: `${original.title} (copy)`,
        items: original.items.map((item) => ({
          ...item,
          id: crypto.randomUUID(),
          slides: item.slides.map((slide) => ({ ...slide, id: crypto.randomUUID() })),
        })),
      };
      await saveService(copy);
      setServices((current) => upsertSummary(current, copy));
    } catch (error) {
      setHomeError(errorMessage(error));
    }
  }

  async function removeService(service: ServiceSummary) {
    if (!window.confirm(`Delete "${service.title}" and all of its slides?`)) return;
    try {
      await deleteService(service.id);
      setServices((current) => current.filter((item) => item.id !== service.id));
      if (currentService?.id === service.id) setCurrentService(null);
    } catch (error) {
      setHomeError(errorMessage(error));
    }
  }

  async function forceSave() {
    if (!currentService) return;
    setSaveStatus("saving");
    try {
      await saveService(currentService);
      setServices((current) => upsertSummary(current, currentService));
      setSaveStatus("saved");
    } catch {
      setSaveStatus("error");
    }
  }

  function updateCurrentService(service: Service) {
    setSaveStatus("saving");
    setCurrentService(service);
  }

  async function openPresentation() {
    if (!isTauri()) {
      setHomeError("The presentation output is available in the Tauri desktop app.");
      return;
    }
    try {
      const existing = await WebviewWindow.getByLabel("presentation");
      if (existing) {
        setIsPresenting(true);
        await existing.setFocus();
        return;
      }
      let automaticDisplay: PresentationDisplay | undefined;
      if (selectedDisplay === "automatic" && displays.length > 1) {
        const monitors = await availableMonitors();
        const activeMonitor = await currentMonitor();
        const activeIndex = activeMonitor
          ? monitors.findIndex((monitor) => monitor.position.x === activeMonitor.position.x && monitor.position.y === activeMonitor.position.y)
          : 0;
        automaticDisplay = displays.find((display) => display.id !== String(activeIndex)) ?? displays[0];
      }
      const target = selectedDisplay === "automatic"
        ? automaticDisplay ?? displays[0]
        : displays.find((display) => display.id === selectedDisplay) ?? displays[0];
      const stage = new WebviewWindow("presentation", {
        url: "index.html?mode=presentation",
        title: "Church Presenter - Live Output",
        fullscreen: true,
        decorations: false,
        resizable: false,
        ...(target ? { x: target.x, y: target.y, width: target.width, height: target.height } : {}),
      });
      stage.once("tauri://created", () => setIsPresenting(true));
      stage.once("tauri://error", (event) => {
        setHomeError(`Could not open the presentation window: ${String(event.payload)}`);
        setIsPresenting(false);
      });
    } catch (error) {
      setHomeError(errorMessage(error));
    }
  }

  async function closePresentation() {
    if (isTauri()) {
      const stage = await WebviewWindow.getByLabel("presentation");
      if (stage) await stage.close();
    }
    setIsPresenting(false);
  }

  function goLive(slide: LiveSlide) {
    setLiveSlide(slide);
    setScreenMode("slide");
  }

  function navigate(direction: -1 | 1) {
    navigateRef.current(direction);
  }

  function jumpTo(edge: "first" | "last") {
    if (!currentService) return;
    const sequence = getSlideSequence(currentService);
    const target = edge === "first" ? sequence[0] : sequence[sequence.length - 1];
    if (target) setLiveSlide(target);
    setScreenMode("slide");
  }

  if (!isLoaded) return <div className="loading-screen"><span className="brand-mark">CP</span><span>Church Presenter</span></div>;

  if (!currentService) {
    return <ServiceHome
      services={services}
      error={homeError}
      onOpen={(id) => void openService(id)}
      onCreate={(service) => void createService(service)}
      onRename={(id, title) => void renameService(id, title)}
      onDuplicate={(id) => void duplicateService(id)}
      onDelete={(service) => void removeService(service)}
    />;
  }

  return <ServiceEditor
    service={currentService}
    saveStatus={saveStatus}
    isPresenting={isPresenting}
    liveSlide={liveSlide}
    screenMode={screenMode}
    displays={displays}
    selectedDisplay={selectedDisplay}
    onServiceChange={updateCurrentService}
    onBack={() => { setCurrentService(null); setLiveSlide(null); }}
    onOpenPresentation={() => void openPresentation()}
    onClosePresentation={() => void closePresentation()}
    onGoLive={goLive}
    onNavigate={navigate}
    onJumpTo={jumpTo}
    onScreenMode={setScreenMode}
    onDisplayChange={setSelectedDisplay}
    onForceSave={() => void forceSave()}
  />;
}

function upsertSummary(summaries: ServiceSummary[], service: Service): ServiceSummary[] {
  const summary = { id: service.id, title: service.title, date: service.date };
  return [summary, ...summaries.filter((item) => item.id !== service.id)]
    .sort((first, second) => second.date.localeCompare(first.date));
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export default App;