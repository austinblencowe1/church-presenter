import { useEffect, useRef, useState } from "react";
import { isTauri } from "@tauri-apps/api/core";
import { emitTo, listen } from "@tauri-apps/api/event";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";
import { availableMonitors, currentMonitor } from "@tauri-apps/api/window";
import { PresentationWindow } from "./components/PresentationWindow";
import { StageWindow } from "./components/StageWindow";
import { ServiceEditor, type PresentationDisplay } from "./components/ServiceEditor";
import { ServiceHome } from "./components/ServiceHome";
import { getSlideSequence } from "./data/presentationNavigation";
import { deleteMacro, listMacros, saveMacro } from "./data/macroStore";
import { deleteMedia, deleteService, deleteSong, deleteSlideTheme, listMedia, listServices, listSlideThemes, listSongs, loadService, saveMedia, saveService, saveSlideTheme, saveSong } from "./data/serviceStore";
import type { LiveSlide, MacroAction, MediaAsset, PresentationUpdate, ScreenMode, Service, ServiceSummary, ShowMacro, Song, SlideTheme, StagePresentationUpdate, VideoPlayback, VideoProgress } from "./types/presentation";
import { typographyFromTheme } from "./data/fonts";
import "./App.css";

const windowMode = new URLSearchParams(window.location.search).get("mode");

function App() {
  if (windowMode === "presentation") return <PresentationWindow />;
  if (windowMode === "stage") return <StageWindow />;
  return <ChurchPresenter />;
}

function ChurchPresenter() {
  const [services, setServices] = useState<ServiceSummary[]>([]);
  const [songs, setSongs] = useState<Song[]>([]);
  const [mediaAssets, setMediaAssets] = useState<MediaAsset[]>([]);
  const [macros, setMacros] = useState<ShowMacro[]>(() => listMacros());
  const [isMacroRunning, setIsMacroRunning] = useState(false);
  const [slideThemes, setSlideThemes] = useState<SlideTheme[]>([]);
  const [currentService, setCurrentService] = useState<Service | null>(null);
  const [isLoaded, setIsLoaded] = useState(false);
  const [saveStatus, setSaveStatus] = useState<"saving" | "saved" | "error">("saved");
  const [homeError, setHomeError] = useState("");
  const [liveSlide, setLiveSlide] = useState<LiveSlide | null>(null);
  const [liveMedia, setLiveMedia] = useState<MediaAsset | null>(null);
  const [liveAudio, setLiveAudio] = useState<MediaAsset | null>(null);
  const [videoPlayback, setVideoPlayback] = useState<VideoPlayback>({ playing: false, volume: 1, playbackRate: 1, loop: false, seekTo: 0, seekRequest: 0, endBehavior: "stop" });
  const [videoProgress, setVideoProgress] = useState<VideoProgress>({ currentTime: 0, duration: 0 });
  const [screenMode, setScreenMode] = useState<ScreenMode>("slide");
  const [clearForeground, setClearForeground] = useState(false);
  const [clearBackground, setClearBackground] = useState(false);
  const [liveMessage, setLiveMessage] = useState<string | null>(null);
  const [stageTimer, setStageTimer] = useState<{ label: string; endsAt: number } | null>(null);
  const [isPresenting, setIsPresenting] = useState(false);
  const [audienceDisplayId, setAudienceDisplayId] = useState<string | null>(null);
  const [isStageOpen, setIsStageOpen] = useState(false);
  const [displays, setDisplays] = useState<PresentationDisplay[]>([]);
  const [selectedDisplay, setSelectedDisplay] = useState("automatic");
  const updateRef = useRef<PresentationUpdate>({ liveSlide: null, liveMedia: null, videoPlayback: { playing: false, volume: 1, playbackRate: 1, loop: false, seekTo: 0, seekRequest: 0, endBehavior: "stop" }, screenMode: "slide", clearForeground: false, clearBackground: false, message: null });
  const stageUpdateRef = useRef<StagePresentationUpdate>({ current: null, next: null, message: null, timer: null, audienceDisplayId: null });
  const serviceRef = useRef<Service | null>(null);
  const navigateRef = useRef<(direction: -1 | 1) => void>(() => undefined);
  const goLiveRef = useRef<(slide: LiveSlide) => void>(() => undefined);
  const liveMediaRef = useRef<MediaAsset | null>(null);
  const videoPlaybackRef = useRef(videoPlayback);
  const macroSequenceRef = useRef(0);
  const activeMacroRef = useRef<number | null>(null);
  const macroWaitRef = useRef<(() => void) | null>(null);

  useEffect(() => {
    updateRef.current = { liveSlide, liveMedia, videoPlayback, screenMode, clearForeground, clearBackground, message: liveMessage };
    serviceRef.current = currentService;
    liveMediaRef.current = liveMedia;
    videoPlaybackRef.current = videoPlayback;
  }, [currentService, liveSlide, liveMedia, videoPlayback, screenMode, clearForeground, clearBackground, liveMessage]);

  useEffect(() => {
    const sequence = currentService ? getSlideSequence(currentService) : [];
    const currentIndex = liveSlide && sequence[liveSlide.overallPosition - 1]?.itemId === liveSlide.itemId
      && sequence[liveSlide.overallPosition - 1]?.slide.id === liveSlide.slide.id
      ? liveSlide.overallPosition - 1
      : sequence.findIndex((entry) => entry.itemId === liveSlide?.itemId && entry.slide.id === liveSlide?.slide.id);
    stageUpdateRef.current = {
      current: liveSlide,
      next: currentIndex >= 0 ? sequence[currentIndex + 1] ?? null : null,
      message: liveMessage,
      timer: stageTimer,
      audienceDisplayId,
    };
  }, [currentService, liveSlide, liveMessage, stageTimer, audienceDisplayId]);

  useEffect(() => {
    let active = true;
    void (async () => {
      const [serviceResult, songResult, mediaResult, themeResult] = await Promise.allSettled([listServices(), listSongs(), listMedia(), listSlideThemes()]);
      if (!active) return;
      if (serviceResult.status === "fulfilled") setServices(serviceResult.value);
      if (songResult.status === "fulfilled") setSongs(songResult.value);
      if (mediaResult.status === "fulfilled") setMediaAssets(mediaResult.value);
      if (themeResult.status === "fulfilled") setSlideThemes(themeResult.value);
      const failures = [serviceResult, songResult, mediaResult, themeResult]
        .filter((result): result is PromiseRejectedResult => result.status === "rejected");
      if (failures.length > 0) setHomeError(failures.map(({ reason }) => errorMessage(reason)).join(" · "));
      setIsLoaded(true);
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
      listen("stage-ready", () => {
        void emitTo("stage", "stage-update", stageUpdateRef.current);
      }),
      listen<{ direction: "next" | "previous" }>("presentation-navigation", (event) => {
        navigateRef.current(event.payload.direction === "next" ? 1 : -1);
      }),
      listen<{ edge: "first" | "last" }>("presentation-jump", (event) => {
        const service = serviceRef.current;
        const sequence = service ? getSlideSequence(service) : [];
        const target = event.payload.edge === "first" ? sequence[0] : sequence[sequence.length - 1];
        if (target) goLiveRef.current(target);
      }),
      listen<{ mode: ScreenMode }>("presentation-screen-mode", (event) => setScreenMode(event.payload.mode)),
      listen<VideoProgress>("presentation-video-progress", (event) => setVideoProgress(event.payload)),
      listen<{ mediaId: string }>("presentation-video-ended", (event) => {
        if (liveMediaRef.current?.id !== event.payload.mediaId) return;
        if (videoPlaybackRef.current.endBehavior === "advance") navigateRef.current(1);
        else setVideoPlayback((current) => ({ ...current, playing: false }));
      }),
      listen("presentation-closed", () => { setIsPresenting(false); setAudienceDisplayId(null); }),
      listen("stage-closed", () => setIsStageOpen(false)),
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
      const currentIndex = liveSlide && sequence[liveSlide.overallPosition - 1]?.itemId === liveSlide.itemId
        && sequence[liveSlide.overallPosition - 1]?.slide.id === liveSlide.slide.id
        ? liveSlide.overallPosition - 1
        : sequence.findIndex((entry) => entry.itemId === liveSlide?.itemId && entry.slide.id === liveSlide?.slide.id);
      const targetIndex = Math.min(Math.max(currentIndex + direction, 0), sequence.length - 1);
      goLiveRef.current(sequence[targetIndex]);
    };
  }, [currentService, liveSlide]);

  useEffect(() => {
    if (isTauri()) {
      void emitTo("presentation", "presentation-update", updateRef.current).catch(() => undefined);
      void emitTo("stage", "stage-update", stageUpdateRef.current).catch(() => undefined);
    }
  }, [liveSlide, liveMedia, videoPlayback, screenMode, clearForeground, clearBackground, liveMessage, stageTimer, currentService, audienceDisplayId]);

  async function createService(service: Service) {
    try {
      await saveService(service);
      setServices((current) => upsertSummary(current, service));
      setCurrentService(service);
      setLiveSlide(null);
      setLiveMedia(null);
      setScreenMode("slide");
      setClearForeground(false);
      setClearBackground(false);
      setLiveMessage(null);
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
      setLiveMedia(null);
      setScreenMode("slide");
      setClearForeground(false);
      setClearBackground(false);
      setLiveMessage(null);
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

  async function moveServiceToFolder(serviceId: string, folder: string) {
    try {
      const service = await loadService(serviceId);
      const updated = { ...service, folder };
      await saveService(updated);
      setServices((current) => upsertSummary(current, updated));
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

  function exportCurrentService() {
    if (!currentService) return;
    const blob = new Blob([JSON.stringify(currentService, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const safeName = currentService.title.toLocaleLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "service";
    link.href = url;
    link.download = `${safeName}-${currentService.date || "backup"}.json`;
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function importServiceBackup(imported: Service) {
    const service: Service = {
      ...imported,
      id: crypto.randomUUID(),
      title: `${imported.title} (imported)`,
      items: imported.items.map((item) => ({
        ...item,
        id: crypto.randomUUID(),
        slides: item.slides.map((slide) => ({ ...slide, id: crypto.randomUUID(), themeId: null })),
      })),
    };
    await createService(service);
  }

  async function persistSong(song: Song) {
    await saveSong(song);
    setSongs((current) => [song, ...current.filter((item) => item.id !== song.id)]
      .sort((first, second) => first.title.localeCompare(second.title)));
  }

  async function removeSong(songId: string) {
    await deleteSong(songId);
    setSongs((current) => current.filter((song) => song.id !== songId));
  }

  async function persistMedia(asset: MediaAsset) {
    await saveMedia(asset);
    setMediaAssets((current) => [asset, ...current.filter((entry) => entry.id !== asset.id)]);
  }

  async function removeMedia(mediaId: string) {
    await deleteMedia(mediaId);
    setMediaAssets((current) => current.filter((asset) => asset.id !== mediaId));
  }

  async function persistSlideTheme(theme: SlideTheme) {
    await saveSlideTheme(theme);
    setSlideThemes((current) => [theme, ...current.filter((entry) => entry.id !== theme.id)]
      .sort((first, second) => first.name.localeCompare(second.name)));
    if (currentService) setSaveStatus("saving");
    setCurrentService((current) => {
      if (!current) return current;
      const updated = {
        ...current,
        items: current.items.map((item) => ({
          ...item,
          slides: item.slides.map((slide) => slide.themeId === theme.id ? {
            ...slide,
            background: theme.background,
            fontSize: theme.fontSize,
            textAlign: theme.textAlign,
            transition: theme.transition,
            transitionDuration: theme.transitionDuration,
            ...typographyFromTheme(theme),
          } : slide),
        })),
      };
      return updated;
    });
  }

  async function removeSlideTheme(themeId: string) {
    await deleteSlideTheme(themeId);
    setSlideThemes((current) => current.filter((theme) => theme.id !== themeId));
    if (currentService) setSaveStatus("saving");
    setCurrentService((current) => current ? {
      ...current,
      items: current.items.map((item) => ({
        ...item,
        slides: item.slides.map((slide) => slide.themeId === themeId ? { ...slide, themeId: null } : slide),
      })),
    } : current);
  }

  function updateCurrentService(service: Service) {
    setSaveStatus("saving");
    setCurrentService(service);
    setLiveSlide((active) => {
      if (!active || active.serviceId !== service.id) return active;
      const sequence = getSlideSequence(service);
      const currentOccurrence = sequence[active.overallPosition - 1];
      if (currentOccurrence?.itemId === active.itemId && currentOccurrence.slide.id === active.slide.id) return currentOccurrence;
      return sequence.find((entry) => entry.itemId === active.itemId && entry.slide.id === active.slide.id) ?? null;
    });
  }

  async function openPresentation(displayOverride?: string) {
    if (!isTauri()) {
      setHomeError("The presentation output is available in the Tauri desktop app.");
      return;
    }
    try {
      const requestedDisplay = displayOverride ?? selectedDisplay;
      let automaticDisplay: PresentationDisplay | undefined;
      if (requestedDisplay === "automatic" && displays.length > 1) {
        const monitors = await availableMonitors();
        const activeMonitor = await currentMonitor();
        const activeIndex = activeMonitor
          ? monitors.findIndex((monitor) => monitor.position.x === activeMonitor.position.x && monitor.position.y === activeMonitor.position.y)
          : 0;
        automaticDisplay = displays.find((display) => display.id !== String(activeIndex)) ?? displays[0];
      }
      const target = requestedDisplay === "automatic"
        ? automaticDisplay ?? displays[0]
        : displays.find((display) => display.id === requestedDisplay) ?? displays[0];
      const existing = await WebviewWindow.getByLabel("presentation");
      if (existing && target?.id === audienceDisplayId) {
        setIsPresenting(true);
        await existing.setFocus();
        return;
      }
      if (existing) await existing.close();
      setSelectedDisplay(requestedDisplay);
      setAudienceDisplayId(target?.id ?? null);
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
        setAudienceDisplayId(null);
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
    setAudienceDisplayId(null);
  }

  async function openStageDisplay() {
    if (!isTauri()) {
      setHomeError("The separate stage display is available in the Tauri desktop app.");
      return;
    }
    try {
      const existing = await WebviewWindow.getByLabel("stage");
      if (existing) {
        setIsStageOpen(true);
        await existing.setFocus();
        return;
      }
      const stage = new WebviewWindow("stage", {
        url: "index.html?mode=stage",
        title: "Church Presenter - Stage Display",
        width: 1280,
        height: 720,
        minWidth: 720,
        minHeight: 405,
        resizable: true,
        decorations: true,
      });
      stage.once("tauri://created", () => setIsStageOpen(true));
      stage.once("tauri://destroyed", () => setIsStageOpen(false));
      stage.once("tauri://error", (event) => {
        setHomeError(`Could not open the stage display: ${String(event.payload)}`);
        setIsStageOpen(false);
      });
    } catch (error) {
      setHomeError(errorMessage(error));
    }
  }

  async function closeStageDisplay() {
    if (isTauri()) {
      const stage = await WebviewWindow.getByLabel("stage");
      if (stage) await stage.close();
    }
    setIsStageOpen(false);
  }

  function goLive(slide: LiveSlide) {
    setLiveMedia(null);
    setVideoPlayback((current) => ({ ...current, playing: false }));
    setLiveSlide(slide);
    setScreenMode("slide");
    setClearForeground(false);
    setClearBackground(false);
    const cues = (slide.slide.cueMacroIds ?? [])
      .map((macroId) => macros.find((macro) => macro.id === macroId))
      .filter((macro): macro is ShowMacro => Boolean(macro));
    if (cues.length > 0) {
      void (async () => {
        for (const macro of cues) await runMacro(macro);
      })();
    }
  }

  goLiveRef.current = goLive;

  function goLiveMedia(asset: MediaAsset) {
    setLiveMedia(asset.kind === "video" ? asset : null);
    if (asset.kind === "video") {
      setVideoProgress({ currentTime: 0, duration: 0 });
      setVideoPlayback((current) => ({ ...current, playing: true, seekTo: 0, seekRequest: current.seekRequest + 1 }));
    }
    
    setScreenMode("slide");
    setClearForeground(false);
    setClearBackground(false);
  }

  function goLiveAudio(asset: MediaAsset) {
    setLiveAudio(asset.kind === "audio" ? asset : null);
  }

  function toggleClearForeground() {
    setClearForeground((current) => {
      const cleared = !current;
      if (cleared) setVideoPlayback((playback) => ({ ...playback, playing: false }));
      return cleared;
    });
  }

  async function runMacro(macro: ShowMacro) {
    if (activeMacroRef.current !== null) return;
    const runId = ++macroSequenceRef.current;
    activeMacroRef.current = runId;
    setIsMacroRunning(true);
    try {
      for (const action of macro.actions) {
        if (action.delayMs > 0) await new Promise<void>((resolve) => {
          const timer = window.setTimeout(() => { macroWaitRef.current = null; resolve(); }, action.delayMs);
          macroWaitRef.current = () => { window.clearTimeout(timer); macroWaitRef.current = null; resolve(); };
        });
        if (activeMacroRef.current !== runId) return;
        runMacroAction(action);
      }
    } finally {
      if (activeMacroRef.current === runId) {
        activeMacroRef.current = null;
        setIsMacroRunning(false);
      }
    }
  }

  function stopMacro() {
    activeMacroRef.current = null;
    macroWaitRef.current?.();
    setIsMacroRunning(false);
  }

  function runMacroAction(action: MacroAction) {
    switch (action.kind) {
      case "next-slide": navigateRef.current(1); break;
      case "previous-slide": navigateRef.current(-1); break;
      case "black-screen": setScreenMode("black"); break;
      case "white-screen": setScreenMode("white"); break;
      case "show-slide": setScreenMode("slide"); setClearForeground(false); setClearBackground(false); setLiveMedia(null);  break;
      case "clear-foreground": setClearForeground(true); setVideoPlayback((current) => ({ ...current, playing: false })); break;
      case "clear-background": setClearBackground(true); break;
      case "clear-all":
        setScreenMode("slide"); setClearForeground(true); setClearBackground(true); 
        setLiveMedia(null); setLiveAudio(null); setLiveMessage(null);
        setVideoPlayback((current) => ({ ...current, playing: false }));
        break;
      case "send-message": setLiveMessage(action.value?.trim() || null); break;

      case "play-video": {
        const asset = mediaAssets.find((entry) => entry.id === action.value && entry.kind === "video");
        if (asset) goLiveMedia(asset);
        break;
      }
      case "play-audio": {
        const asset = mediaAssets.find((entry) => entry.id === action.value && entry.kind === "audio");
        if (asset) setLiveAudio(asset);
        break;
      }
    }
  }

  function navigate(direction: -1 | 1) {
    navigateRef.current(direction);
  }

  function jumpTo(edge: "first" | "last") {
    if (!currentService) return;
    const sequence = getSlideSequence(currentService);
    const target = edge === "first" ? sequence[0] : sequence[sequence.length - 1];
    if (target) goLive(target);
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
      onImport={(service) => void importServiceBackup(service)}
      onMoveToFolder={(serviceId, folder) => void moveServiceToFolder(serviceId, folder)}
    />;
  }

  return <ServiceEditor
    service={currentService}
    songs={songs}
    mediaAssets={mediaAssets}
    macros={macros}
    isMacroRunning={isMacroRunning}
    slideThemes={slideThemes}
    saveStatus={saveStatus}
    isPresenting={isPresenting}
    isStageOpen={isStageOpen}
    liveSlide={liveSlide}
    liveMedia={liveMedia}
    liveAudio={liveAudio}
    videoPlayback={videoPlayback}
    videoProgress={videoProgress}
    screenMode={screenMode}
    clearForeground={clearForeground}
    clearBackground={clearBackground}
    displays={displays}
    selectedDisplay={selectedDisplay}
    onServiceChange={updateCurrentService}
    onBack={() => { setCurrentService(null); setLiveSlide(null); setLiveMedia(null); }}
    onOpenPresentation={() => void openPresentation()}
    onClosePresentation={() => void closePresentation()}
    onOpenStage={() => void openStageDisplay()}
    onCloseStage={() => void closeStageDisplay()}
    onGoLive={goLive}
    onGoLiveMedia={goLiveMedia}
    onGoLiveAudio={goLiveAudio}
    onVideoPlaybackChange={(update) => {
      if (update.playing) setClearForeground(false);
      setVideoPlayback((current) => ({ ...current, ...update }));
    }}
    onSeekVideo={(time) => setVideoPlayback((current) => ({ ...current, seekTo: time, seekRequest: current.seekRequest + 1 }))}
    onClearAudio={() => setLiveAudio(null)}
    onSaveMacro={(macro) => setMacros(saveMacro(macro))}
    onDeleteMacro={(macroId) => setMacros(deleteMacro(macroId))}
    onRunMacro={(macro) => void runMacro(macro)}
    onStopMacro={stopMacro}
    onNavigate={navigate}
    onJumpTo={jumpTo}
    onScreenMode={setScreenMode}
    onClearForeground={toggleClearForeground}
    onClearBackground={() => setClearBackground((current) => !current)}
    onClearAll={() => {
      setScreenMode("slide");
      setClearForeground(true);
      setClearBackground(true);
      setLiveMedia(null);
      setLiveAudio(null);
      setLiveMessage(null);
      setVideoPlayback((current) => ({ ...current, playing: false }));
      
    }}
    onTriggerMessage={(message) => setLiveMessage(message.trim() || null)}
    liveMessage={liveMessage}
    onStartTimer={(label, durationSeconds) => setStageTimer({ label, endsAt: Date.now() + durationSeconds * 1000 })}
    onClearTimer={() => setStageTimer(null)}
    stageTimer={stageTimer}
    onDisplayChange={setSelectedDisplay}
    onForceSave={() => void forceSave()}
    onExportService={exportCurrentService}
    onSongSave={persistSong}
    onSongDelete={removeSong}
    onMediaSave={persistMedia}
    onMediaDelete={removeMedia}
    onSlideThemeSave={persistSlideTheme}
    onSlideThemeDelete={removeSlideTheme}
  />;
}

function upsertSummary(summaries: ServiceSummary[], service: Service): ServiceSummary[] {
  const summary = { id: service.id, title: service.title, date: service.date, folder: service.folder ?? "" };
  return [summary, ...summaries.filter((item) => item.id !== service.id)]
    .sort((first, second) => second.date.localeCompare(first.date));
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export default App;
