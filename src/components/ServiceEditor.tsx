import { useEffect, useRef, useState } from "react";
import type { ChangeEvent, MouseEvent, SyntheticEvent } from "react";
import { SlideCanvas } from "./SlideCanvas";
import { ServiceItemOverview } from "./ServiceItemOverview";
import { SongLibraryDialog } from "./SongLibraryDialog";
import { ScripturePanel, type SelectedScriptureVerse } from "./ScripturePanel";
import { getItemSequenceSlides, getSlideSequence } from "../data/presentationNavigation";
import { deleteSlideTemplate, listSlideTemplates, saveSlideTemplate } from "../data/templateStore";
import { importDeckFile } from "../data/importSlides";


import type {
  LiveSlide,
  MediaAsset,
  MediaKind,
  ScreenMode,
  Service,
  ServiceItem,
  ServiceItemType,
  Slide,
  SlideObject,
  SlideObjectKind,
  SlideTemplate,
  MacroAction,
  MacroActionKind,
  ShowMacro,
  SlideTransition,
  Song,
  SlideTheme,
  TextAlignment,
  VideoPlayback,
  VideoProgress,
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
  songs: Song[];
  mediaAssets: MediaAsset[];
  macros: ShowMacro[];
  isMacroRunning: boolean;
  slideThemes: SlideTheme[];
  saveStatus: "saving" | "saved" | "error";
  isPresenting: boolean;
  isStageOpen: boolean;
  liveSlide: LiveSlide | null;
  liveMedia: MediaAsset | null;
  liveAudio: MediaAsset | null;
  videoPlayback: VideoPlayback;
  videoProgress: VideoProgress;
  screenMode: ScreenMode;
  clearForeground: boolean;
  clearBackground: boolean;
  displays: PresentationDisplay[];
  selectedDisplay: string;
  onServiceChange: (service: Service) => void;
  onBack: () => void;
  onOpenPresentation: () => void;
  onClosePresentation: () => void;
  onOpenStage: () => void;
  onCloseStage: () => void;
  onGoLive: (slide: LiveSlide) => void;
  onGoLiveMedia: (asset: MediaAsset) => void;
  onGoLiveAudio: (asset: MediaAsset) => void;
  onVideoPlaybackChange: (update: Partial<VideoPlayback>) => void;
  onSeekVideo: (time: number) => void;
  onClearAudio: () => void;
  onSaveMacro: (macro: ShowMacro) => void;
  onDeleteMacro: (macroId: string) => void;
  onRunMacro: (macro: ShowMacro) => void;
  onStopMacro: () => void;
  onNavigate: (direction: -1 | 1) => void;
  onJumpTo: (edge: "first" | "last") => void;
  onScreenMode: (mode: ScreenMode) => void;
  onClearForeground: () => void;
  onClearBackground: () => void;
  onClearAll: () => void;
  onTriggerMessage: (message: string) => void;
  liveMessage: string | null;
  onStartTimer: (label: string, durationSeconds: number) => void;
  onClearTimer: () => void;
  stageTimer: { label: string; endsAt: number } | null;
  onDisplayChange: (displayId: string) => void;
  onForceSave: () => void;
  onExportService: () => void;
  onSongSave: (song: Song) => Promise<void>;
  onSongDelete: (songId: string) => Promise<void>;
  onMediaSave: (asset: MediaAsset) => Promise<void>;
  onMediaDelete: (mediaId: string) => Promise<void>;
  onSlideThemeSave: (theme: SlideTheme) => Promise<void>;
  onSlideThemeDelete: (themeId: string) => Promise<void>;
}

interface SlideContextMenuState {
  x: number;
  y: number;
  itemId: string;
  slideId: string;
}

interface DraggedSlide {
  itemId: string;
  slideId: string;
}

const itemTypes: Array<{ value: ServiceItemType; label: string }> = [
  { value: "welcome", label: "Welcome" },
  { value: "song", label: "Song" },
  { value: "bible", label: "Bible Reading" },
  { value: "sermon", label: "Sermon" },
  { value: "announcement", label: "Announcement" },
  { value: "other", label: "Other" },
];

const alignmentGlyphs: Record<TextAlignment, string> = { left: "◀", center: "≡", right: "▶" };

const macroActions: Array<{ value: MacroActionKind; label: string }> = [
  { value: "next-slide", label: "Next slide" }, { value: "previous-slide", label: "Previous slide" },
  { value: "black-screen", label: "Black screen" }, { value: "white-screen", label: "White screen" }, { value: "show-slide", label: "Restore slide output" },
  { value: "clear-foreground", label: "Clear foreground"}, { value: "clear-background", label: "Clear background" },
 { value: "clear-all", label: "Clear all" },
  { value: "send-message", label: "Send message" },
  { value: "play-video", label: "Play video" }, { value: "play-audio", label: "Play audio" },
];

const macroValueKinds: MacroActionKind[] = ["send-message", "play-video", "play-audio"];

function formatFileSize(size: number): string {
  if (size < 1024 * 1024) return `${Math.max(1, Math.round(size / 1024))} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function formatTime(seconds: number): string {
  if (!Number.isFinite(seconds) || seconds < 0) return "0:00";
  const minutes = Math.floor(seconds / 60);
  const remainder = Math.floor(seconds % 60).toString().padStart(2, "0");
  return `${minutes}:${remainder}`;
}

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
  songs,
  mediaAssets,
  macros,
  isMacroRunning,
  slideThemes,
  saveStatus,
  isPresenting,
  isStageOpen,
  liveSlide,
  liveMedia,
  liveAudio,
  videoPlayback,
  videoProgress,
  screenMode,
  clearForeground,
  clearBackground,
  displays,
  selectedDisplay,
  onServiceChange,
  onBack,
  onOpenPresentation,
  onClosePresentation,
  onOpenStage,
  onCloseStage,
  onGoLive,
  onGoLiveMedia,
  onGoLiveAudio,
  onVideoPlaybackChange,
  onSeekVideo,
  onClearAudio,
  onSaveMacro,
  onDeleteMacro,
  onRunMacro,
  onStopMacro,
  onNavigate,
  onJumpTo,
  onScreenMode,
  onClearForeground,
  onClearBackground,
  onClearAll,
  onTriggerMessage,
  liveMessage,
  onStartTimer,
  onClearTimer,
  stageTimer,
  onDisplayChange,
  onForceSave,
  onExportService,
  onSongSave,
  onSongDelete,
  onMediaSave,
  onMediaDelete,
  onSlideThemeSave,
  onSlideThemeDelete,
}: ServiceEditorProps) {
  const [panelWidths, setPanelWidths] = useState(() => {
    try {
      const stored = JSON.parse(localStorage.getItem("church-presenter-panel-widths") ?? "{}");
      return {
        sidebar: typeof stored.sidebar === "number" ? stored.sidebar : 300,
        preview: typeof stored.preview === "number" ? stored.preview : 330,
      };
    } catch {
      return { sidebar: 300, preview: 330 };
    }
  });
  const [importStatus, setImportStatus] = useState("");
  const [importError, setImportError] = useState("");
  const importInputRef = useRef<HTMLInputElement | null>(null);

  const draggingPanel = useRef<"sidebar" | "preview" | null>(null);
  const previewVideoRef = useRef<HTMLVideoElement | null>(null);
  const [selectedItemId, setSelectedItemId] = useState(service.items[0]?.id ?? "");
  const [selectedSlideId, setSelectedSlideId] = useState(service.items[0]?.slides[0]?.id ?? "");
  const [selectedObjectId, setSelectedObjectId] = useState<string | null>(null);
  const [expandedItems, setExpandedItems] = useState<Set<string>>(() => new Set(service.items.map((item) => item.id)));
  const [newItemType, setNewItemType] = useState<ServiceItemType>("song");
  const [editingItemId, setEditingItemId] = useState("");
  const [editingItemTitle, setEditingItemTitle] = useState("");
  const [recentlyRemovedItem, setRecentlyRemovedItem] = useState<{ item: ServiceItem; index: number } | null>(null);
  const [viewMode, setViewMode] = useState<"overview" | "editor" | "bible">("overview");
  const [overviewSize, setOverviewSize] = useState(42);
  const [songDialogOpen, setSongDialogOpen] = useState(false);
  const [songTargetItemId, setSongTargetItemId] = useState<string | null>(null);
  const [utilityTab, setUtilityTab] = useState<"songs" | "media" | "messages" | "timers" | "macros" | null>(null);
  const [slideContextMenu, setSlideContextMenu] = useState<SlideContextMenuState | null>(null);
  const [draggedItemId, setDraggedItemId] = useState("");
  const [mediaError, setMediaError] = useState("");
  const [mediaSearch, setMediaSearch] = useState("");
  const [mediaKindFilter, setMediaKindFilter] = useState<"all" | MediaKind>("all");
  const [timerLabel, setTimerLabel] = useState("Sermon");
  const [timerMinutes, setTimerMinutes] = useState("20");
  const [macroName, setMacroName] = useState("");
  const [macroDraftActions, setMacroDraftActions] = useState<MacroAction[]>([]);
  const [macroActionKind, setMacroActionKind] = useState<MacroActionKind>("next-slide");
  const [macroActionValue, setMacroActionValue] = useState("");
  const [macroActionDelay, setMacroActionDelay] = useState("0");
  const [editingMacroId, setEditingMacroId] = useState<string | null>(null);
  const [cueMacroSelection, setCueMacroSelection] = useState("");
  const [selectedThemeId, setSelectedThemeId] = useState(slideThemes[0]?.id ?? "");
  const [newThemeName, setNewThemeName] = useState("");
  const [themeError, setThemeError] = useState("");
  const [slideTemplates, setSlideTemplates] = useState<SlideTemplate[]>(() => listSlideTemplates());
  const [selectedTemplateId, setSelectedTemplateId] = useState(() => listSlideTemplates()[0]?.id ?? "");
  const [templateName, setTemplateName] = useState("");
  const [templateDraft, setTemplateDraft] = useState("{{title}}\n\n{{body}}");
  const [templateValues, setTemplateValues] = useState<Record<string, string>>({});
  const [templateError, setTemplateError] = useState("");
  const draggedSlideRef = useRef<DraggedSlide | null>(null);
  const sequence = getSlideSequence(service);
  const selectedItem = service.items.find((item) => item.id === selectedItemId) ?? null;
  const selectedSlide = selectedItem?.slides.find((slide) => slide.id === selectedSlideId) ?? null;
  const selectedObject = selectedSlide?.objects?.find((object) => object.id === selectedObjectId) ?? null;
  const selectedTheme = slideThemes.find((theme) => theme.id === selectedThemeId) ?? null;
  const selectedTemplate = slideTemplates.find((template) => template.id === selectedTemplateId) ?? null;
  const templateFields = selectedTemplate
    ? [...new Set(Array.from(selectedTemplate.content.matchAll(/{{\s*([\w-]+)\s*}}/g), (match) => match[1]))]
    : [];
  const selectedItemSlideIndex = selectedItem?.slides.findIndex((slide) => slide.id === selectedSlideId) ?? -1;
  const previewSlide = selectedSlide && selectedItem
    ? sequence.find((step) => step.itemId === selectedItem.id && step.slide.id === selectedSlide.id) ?? null
    : null;
  const filteredMediaAssets = mediaAssets
    .filter((asset) => (mediaKindFilter === "all" || asset.kind === mediaKindFilter)
      && asset.name.toLocaleLowerCase().includes(mediaSearch.trim().toLocaleLowerCase()))
    .sort((first, second) => second.addedAt.localeCompare(first.addedAt));
  const liveSequenceIndex = liveSlide && sequence[liveSlide.overallPosition - 1]?.slide.id === liveSlide.slide.id ? liveSlide.overallPosition - 1 : sequence.findIndex((step) => step.serviceId === liveSlide?.serviceId && step.slide.id === liveSlide?.slide.id);
  const isPreviewLive = previewSlide?.slide.id === liveSlide?.slide.id && previewSlide?.serviceId === liveSlide?.serviceId;

  useEffect(() => {
    setSelectedItemId(service.items[0]?.id ?? "");
    setSelectedSlideId(service.items[0]?.slides[0]?.id ?? "");
    setExpandedItems(new Set(service.items.map((item) => item.id)));
    setViewMode("overview");
  }, [service.id]);

  useEffect(() => {
    const video = previewVideoRef.current;
    if (!video || liveMedia?.kind !== "video") return;
    video.muted = true;
    if (Math.abs(video.currentTime - videoProgress.currentTime) > 0.75) video.currentTime = videoProgress.currentTime;
    if (videoPlayback.playing) void video.play().catch(() => undefined);
    else video.pause();
  }, [liveMedia?.id, videoPlayback.playing, videoProgress.currentTime]);

  useEffect(() => {
    if (selectedItem && !selectedItem.slides.some((slide) => slide.id === selectedSlideId)) {
      setSelectedSlideId(selectedItem.slides[0]?.id ?? "");
    }
  }, [selectedItem, selectedSlideId]);

  useEffect(() => {
    localStorage.setItem("church-presenter-panel-widths", JSON.stringify(panelWidths));
  }, [panelWidths]);

  useEffect(() => {
    if (!slideThemes.some((theme) => theme.id === selectedThemeId)) setSelectedThemeId(slideThemes[0]?.id ?? "");
  }, [slideThemes, selectedThemeId]);

  useEffect(() => {
    setTemplateDraft(selectedSlide?.text ?? "{{title}}\n\n{{body}}");
    setTemplateName("");
    setSelectedObjectId(null);
    setCueMacroSelection(macros[0]?.id ?? "");
  }, [selectedSlide?.id]);

  useEffect(() => {
    setTemplateValues({});
    setTemplateError("");
  }, [selectedTemplateId]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const target = event.target;
      if (target instanceof HTMLElement && (target.isContentEditable || ["INPUT", "TEXTAREA", "SELECT"].includes(target.tagName))) return;
      if (target instanceof HTMLButtonElement && [" ", "Enter"].includes(event.key)) return;

      if (event.ctrlKey && event.key.toLowerCase() === "s") {
        event.preventDefault();
        onForceSave();
      } else if (event.altKey && event.key.toLowerCase() === "b") {
        event.preventDefault();
        onClearBackground();
      } else if (event.altKey && event.key.toLowerCase() === "x") {
        event.preventDefault();
        onClearAll();
      } else if (event.altKey && event.key === "Backspace") {
        event.preventDefault();
        onClearForeground();
      } else if (event.key === "Backspace" && selectedItem && selectedSlide) {
        event.preventDefault();
        removeSlideFromItem(selectedItem.id, selectedSlide.id);
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
  }, [isPresenting, onForceSave, onGoLive, onJumpTo, onNavigate, onScreenMode, onClosePresentation, onClearAll, onClearBackground, onClearForeground, previewSlide, selectedItem, selectedSlide]);

  useEffect(() => {
    const closeMenu = () => setSlideContextMenu(null);
    window.addEventListener("click", closeMenu);
    window.addEventListener("scroll", closeMenu, true);
    return () => {
      window.removeEventListener("click", closeMenu);
      window.removeEventListener("scroll", closeMenu, true);
    };
  }, []);

  function updateService(update: Partial<Service>) {
    onServiceChange({ ...service, ...update });
  }

  function changeItem(itemId: string, update: Partial<ServiceItem>) {
    updateService({ items: service.items.map((item) => item.id === itemId ? { ...item, ...update } : item) });
  }

  function addItem() {
    const item: ServiceItem = {
      id: crypto.randomUUID(),
      title: newItemType === "song" ? "Choose a song?" : newItemType === "bible" ? "Choose a passage?" : itemDefaultTitle(newItemType),
      type: newItemType,
      slides: newItemType === "song" || newItemType === "bible" ? [] : [createSlide()],
      reference: "",
      bibleVersion: "WEB",
    };
    updateService({ items: [...service.items, item] });
    setSelectedItemId(item.id);
    setSelectedSlideId(item.slides[0]?.id ?? "");
    setExpandedItems((current) => new Set(current).add(item.id));
    if (newItemType === "song") {
      setSongTargetItemId(item.id);
      setSongDialogOpen(true);
      setUtilityTab("songs");
    } else if (newItemType === "bible") {
      setViewMode("bible");
    } else {
      setViewMode("overview");
    }
  }

  function removeItem(itemId: string) {
    const index = service.items.findIndex((item) => item.id === itemId);
    const removed = service.items[index];
    if (!removed) return;
    setRecentlyRemovedItem({ item: removed, index });
    window.setTimeout(() => setRecentlyRemovedItem((current) => current?.item.id === itemId ? null : current), 8000);
    const items = service.items.filter((item) => item.id !== itemId);
    updateService({ items });
    if (selectedItemId === itemId) {
      setSelectedItemId(items[0]?.id ?? "");
      setSelectedSlideId(items[0]?.slides[0]?.id ?? "");
    }
  }

  function undoRemoveItem() {
    if (!recentlyRemovedItem) return;
    const items = [...service.items];
    items.splice(Math.min(recentlyRemovedItem.index, items.length), 0, recentlyRemovedItem.item);
    updateService({ items });
    setSelectedItemId(recentlyRemovedItem.item.id);
    setSelectedSlideId(recentlyRemovedItem.item.slides[0]?.id ?? "");
    setRecentlyRemovedItem(null);
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
      slides: selectedItem.slides.map((slide) => {
        if (slide.id !== selectedSlide.id) return slide;
        const styleChange = ["background", "fontSize", "textAlign", "transition", "transitionDuration"]
          .some((key) => key in update) && !("themeId" in update);
        return { ...slide, ...update, ...(styleChange ? { themeId: null } : {}) };
      }),
    });
  }

  function addSlideObject(kind: SlideObjectKind, source?: string) {
    if (!selectedSlide) return;
    const object: SlideObject = {
      id: crypto.randomUUID(),
      kind,
      x: kind === "text" ? 18 : 32,
      y: kind === "text" ? 35 : 32,
      width: kind === "text" ? 64 : 36,
      height: kind === "text" ? 25 : 36,
      rotation: 0,
      opacity: 1,
      fill: kind === "rectangle" || kind === "ellipse" ? "#367bf5" : "transparent",
      stroke: "#ffffff",
      strokeWidth: 0,
      color: "#ffffff",
      text: kind === "text" ? "New text" : "",
      fontSize: 48,
      textAlign: "center",
      ...(source ? { source } : {}),
    };
    changeSlide({ objects: [...(selectedSlide.objects ?? []), object] });
    setSelectedObjectId(object.id);
  }

  function changeSlideObject(objectId: string, update: Partial<SlideObject>) {
    if (!selectedSlide) return;
    changeSlide({ objects: (selectedSlide.objects ?? []).map((object) => object.id === objectId ? { ...object, ...update } : object) });
  }

  function addMediaObject(asset: MediaAsset) {
    if (asset.kind === "image" || asset.kind === "video") addSlideObject(asset.kind, asset.source);
  }

  function slideThemeFromCurrent(id: string, name: string): SlideTheme | null {
    if (!selectedSlide) return null;
    return {
      id,
      name,
      background: selectedSlide.background,
      fontSize: selectedSlide.fontSize,
      textAlign: selectedSlide.textAlign,
      transition: selectedSlide.transition ?? "cut",
      transitionDuration: selectedSlide.transitionDuration ?? 300,
    };
  }

  function applyThemeToSlide(theme: SlideTheme) {
    changeSlide({
      background: theme.background,
      fontSize: theme.fontSize,
      textAlign: theme.textAlign,
      transition: theme.transition,
      transitionDuration: theme.transitionDuration,
      themeId: theme.id,
    });
    setSelectedThemeId(theme.id);
  }

  function applyThemeToService(theme: SlideTheme) {
    onServiceChange({
      ...service,
      items: service.items.map((item) => ({
        ...item,
        slides: item.slides.map((slide) => ({
          ...slide,
          background: theme.background,
          fontSize: theme.fontSize,
          textAlign: theme.textAlign,
          transition: theme.transition,
          transitionDuration: theme.transitionDuration,
          themeId: theme.id,
        })),
      })),
    });
    setSelectedThemeId(theme.id);
  }

  async function saveCurrentTheme(event: SyntheticEvent<HTMLFormElement>) {
    event.preventDefault();
    const name = newThemeName.trim();
    if (!name || !selectedSlide) return;
    const theme = slideThemeFromCurrent(crypto.randomUUID(), name);
    if (!theme) return;
    try {
      await onSlideThemeSave(theme);
      setSelectedThemeId(theme.id);
      changeSlide({
        background: theme.background,
        fontSize: theme.fontSize,
        textAlign: theme.textAlign,
        transition: theme.transition,
        transitionDuration: theme.transitionDuration,
        themeId: theme.id,
      });
      setNewThemeName("");
      setThemeError("");
    } catch (error) {
      setThemeError(error instanceof Error ? error.message : String(error));
    }
  }

  async function updateThemeFromCurrentSlide() {
    if (!selectedTheme || !selectedSlide) return;
    const updated = slideThemeFromCurrent(selectedTheme.id, selectedTheme.name);
    if (!updated) return;
    try {
      await onSlideThemeSave(updated);
      setSelectedThemeId(updated.id);
      changeSlide({ themeId: updated.id });
      setThemeError("");
    } catch (error) {
      setThemeError(error instanceof Error ? error.message : String(error));
    }
  }

  function uploadSlideMedia(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.currentTarget.files ?? []);
    event.currentTarget.value = "";
    files.forEach(importMediaFile);
  }

  function importMediaFile(file: File) {
    if (!file) return;
    const kind = file.type.startsWith("image/") ? "image" : file.type.startsWith("video/") ? "video" : file.type.startsWith("audio/") ? "audio" : null;
    if (!kind) {
      setMediaError("Choose an image, video, or audio file.");
      return;
    }
    const reader = new FileReader();
    reader.addEventListener("load", () => {
      if (typeof reader.result !== "string") return;
      const asset: MediaAsset = {
        id: crypto.randomUUID(),
        name: file.name,
        kind,
        source: reader.result,
        size: file.size,
        addedAt: new Date().toISOString(),
      };
      void onMediaSave(asset).then(() => {
        setMediaError("");
        if (selectedSlide && kind !== "audio") changeSlide({ backgroundMedia: asset.source });
      }).catch((error: unknown) => {
        setMediaError(error instanceof Error ? error.message : String(error));
      });
    });
    reader.addEventListener("error", () => setMediaError("The selected file could not be read."));
    reader.readAsDataURL(file);
  }

  async function importSlidesFile(file: File) {
  setImportError("");
  setImportStatus("Starting import…");
  try {
    const { item, mediaAssets } = await importDeckFile(file, {
      onProgress: (message) => setImportStatus(message),
    });
    for (const asset of mediaAssets) {
      await onMediaSave(asset);
    }
    updateService({ items: [...service.items, item] });
    setSelectedItemId(item.id);
    setSelectedSlideId(item.slides[0]?.id ?? "");
    setExpandedItems((current) => new Set(current).add(item.id));
    setViewMode("overview");
    setImportStatus(`Imported ${item.slides.length} slide${item.slides.length === 1 ? "" : "s"}.`);
  } catch (error) {
    setImportStatus("");
    setImportError(error instanceof Error ? error.message : String(error));
  }
}

function onImportSlidesChange(event: ChangeEvent<HTMLInputElement>) {
  const file = event.currentTarget.files?.[0];
  event.currentTarget.value = "";
  if (file) void importSlidesFile(file);
}

  function removeSlideFromItem(itemId: string, slideId: string) {
    const item = service.items.find((entry) => entry.id === itemId);
    if (!item) return;
    const slideIndex = item.slides.findIndex((slide) => slide.id === slideId);
    const slides = item.slides.filter((slide) => slide.id !== slideId);
    changeItem(itemId, { slides });
    if (selectedItemId === itemId && selectedSlideId === slideId) {
      setSelectedSlideId(slides[Math.min(slideIndex, slides.length - 1)]?.id ?? "");
    }
  }

  function removeSlide() {
    if (!selectedItem || !selectedSlide) return;
    removeSlideFromItem(selectedItem.id, selectedSlide.id);
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

  function goLiveSlide(itemId: string, slide: Slide, itemPosition?: number) {
    const target = itemPosition
      ? sequence.find((entry) => entry.itemId === itemId && entry.itemSlidePosition === itemPosition)
      : sequence.find((entry) => entry.itemId === itemId && entry.slide.id === slide.id);
    if (target) onGoLive(target);
  }

  function addScriptureToService(verses: SelectedScriptureVerse[]) {
    if (verses.length === 0) return;
    const itemId = selectedItem?.type === "bible" ? selectedItem.id : crypto.randomUUID();
    const slides = verses.map((verse) => ({
      ...createSlide(),
      text: `${verse.text}\n\n${verse.reference} · ${verse.translation}`,
      sectionLabel: verse.reference,
      fontSize: 42,
    }));
    const item: ServiceItem = {
      id: itemId,
      title: verses.length === 1 ? verses[0].reference : `Scripture · ${verses[0].reference} – ${verses[verses.length - 1].reference}`,
      type: "bible",
      slides,
      reference: verses.map((verse) => verse.reference).join(", "),
      bibleVersion: verses[0].translation,
    };
    updateService({ items: [...service.items, item] });
    setSelectedItemId(item.id);
    setSelectedSlideId(slides[0].id);
    setExpandedItems((current) => new Set([...current, item.id]));
    setViewMode("overview");
  }

  function goLiveScripture(verse: SelectedScriptureVerse) {
    const id = crypto.randomUUID();
    const slide: Slide = {
      ...(selectedSlide ?? createSlide()),
      id,
      text: `${verse.text}\n\n${verse.reference} · ${verse.translation}`,
      sectionLabel: verse.reference,
      fontSize: 42,
    };
    onGoLive({
      serviceId: service.id,
      serviceTitle: service.title,
      itemId: `scripture-${id}`,
      itemTitle: verse.reference,
      itemType: "bible",
      itemSlidePosition: 1,
      itemSlideTotal: 1,
      overallPosition: liveSlide?.overallPosition ?? 1,
      overallTotal: Math.max(1, sequence.length),
      slide,
    });
  }

  function saveCurrentAsTemplate(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!templateName.trim() || !selectedSlide) return;
    const saved = saveSlideTemplate({
      id: crypto.randomUUID(),
      name: templateName.trim(),
      content: templateDraft,
      background: selectedSlide.background,
      fontSize: selectedSlide.fontSize,
      textAlign: selectedSlide.textAlign,
      transition: selectedSlide.transition ?? "cut",
      transitionDuration: selectedSlide.transitionDuration ?? 300,
    });
    setSlideTemplates(saved);
    setSelectedTemplateId(saved[0].id);
    setTemplateName("");
  }

  function appendMacroAction() {
    const delayMs = Math.max(0, Math.min(600_000, Math.round((Number(macroActionDelay) || 0) * 1000)));
    setMacroDraftActions((current) => [...current, {
      id: crypto.randomUUID(),
      kind: macroActionKind,
      delayMs,
      ...(macroValueKinds.includes(macroActionKind) ? { value: macroActionValue } : {}),
    }]);
    setMacroActionValue("");
    setMacroActionDelay("0");
  }

  function moveMacroAction(index: number, direction: -1 | 1) {
    const target = index + direction;
    if (target < 0 || target >= macroDraftActions.length) return;
    setMacroDraftActions((current) => {
      const next = [...current];
      [next[index], next[target]] = [next[target], next[index]];
      return next;
    });
  }

  function saveMacroDraft(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!macroName.trim() || macroDraftActions.length === 0) return;
    onSaveMacro({ id: editingMacroId ?? crypto.randomUUID(), name: macroName.trim(), actions: macroDraftActions });
    setMacroName("");
    setMacroDraftActions([]);
    setEditingMacroId(null);
  }

  function applySlideTemplate() {
    if (!selectedItem || !selectedTemplate) return;
    const content = selectedTemplate.content.replace(/{{\s*([\w-]+)\s*}}/g, (_, field: string) => templateValues[field] ?? "");
    const slide: Slide = {
      ...createSlide(),
      text: content,
      background: selectedTemplate.background,
      fontSize: selectedTemplate.fontSize,
      textAlign: selectedTemplate.textAlign,
      transition: selectedTemplate.transition,
      transitionDuration: selectedTemplate.transitionDuration,
    };
    changeItem(selectedItem.id, { slides: [...selectedItem.slides, slide] });
    setSelectedSlideId(slide.id);
    setViewMode("editor");
    setTemplateError("");
  }

  function reorderSlide(itemId: string, targetSlideId: string) {
    const dragged = draggedSlideRef.current;
    draggedSlideRef.current = null;
    if (!dragged || dragged.itemId !== itemId || dragged.slideId === targetSlideId) return;
    const item = service.items.find((entry) => entry.id === itemId);
    if (!item) return;
    const fromIndex = item.slides.findIndex((slide) => slide.id === dragged.slideId);
    const toIndex = item.slides.findIndex((slide) => slide.id === targetSlideId);
    if (fromIndex < 0 || toIndex < 0) return;
    const slides = [...item.slides];
    const [movedSlide] = slides.splice(fromIndex, 1);
    slides.splice(toIndex, 0, movedSlide);
    changeItem(itemId, { slides });
  }

  function reorderItem(sourceItemId: string, targetItemId: string) {
    if (!sourceItemId || sourceItemId === targetItemId) return;
    const fromIndex = service.items.findIndex((item) => item.id === sourceItemId);
    const toIndex = service.items.findIndex((item) => item.id === targetItemId);
    if (fromIndex < 0 || toIndex < 0) return;
    const items = [...service.items];
    const [movedItem] = items.splice(fromIndex, 1);
    items.splice(toIndex, 0, movedItem);
    updateService({ items });
  }

  function addSongItem(item: ServiceItem) {
    const targetId = songTargetItemId;
    const serviceItem = targetId ? { ...item, id: targetId } : item;
    updateService({ items: targetId
      ? service.items.map((entry) => entry.id === targetId ? serviceItem : entry)
      : [...service.items, serviceItem] });
    setSongTargetItemId(null);
    setSelectedItemId(serviceItem.id);
    setSelectedSlideId(serviceItem.slides[0]?.id ?? "");
    setExpandedItems((current) => new Set(current).add(serviceItem.id));
    setViewMode("overview");
  }

  function openSlideContextMenu(event: MouseEvent<HTMLButtonElement>, itemId: string, slideId: string) {
    event.preventDefault();
    setSelectedItemId(itemId);
    setSelectedSlideId(slideId);
    setSlideContextMenu({
      x: Math.min(event.clientX, window.innerWidth - 170),
      y: Math.min(event.clientY, window.innerHeight - 150),
      itemId,
      slideId,
    });
  }

  function runSlideContextAction(action: "live" | "edit" | "duplicate" | "delete") {
    if (!slideContextMenu) return;
    const item = service.items.find((entry) => entry.id === slideContextMenu.itemId);
    const slide = item?.slides.find((entry) => entry.id === slideContextMenu.slideId);
    if (action === "live" && slide) goLiveSlide(slideContextMenu.itemId, slide);
    if (action === "edit") setViewMode("editor");
    if (action === "duplicate" && item && slide) {
      const index = item.slides.findIndex((entry) => entry.id === slide.id);
      const copy = { ...slide, id: crypto.randomUUID() };
      const slides = [...item.slides];
      slides.splice(index + 1, 0, copy);
      changeItem(item.id, { slides });
      setSelectedSlideId(copy.id);
    }
    if (action === "delete") removeSlideFromItem(slideContextMenu.itemId, slideContextMenu.slideId);
    setSlideContextMenu(null);
  }

  function renameItem(event: SyntheticEvent<HTMLFormElement>, itemId: string) {
    event.preventDefault();
    if (editingItemTitle.trim()) changeItem(itemId, { title: editingItemTitle.trim() });
    setEditingItemId("");
  }

  return (
    <main className="app-shell service-editor-shell console-theme-light">
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
            <span className="save-status__check">{saveStatus === "saving" ? "…" : saveStatus === "error" ? "!" : "✓"}</span>
            {saveStatus === "saving" ? "Saving..." : saveStatus === "error" ? "Save failed" : "Saved"}
          </span>
          <label className="display-picker">
            <span>DISPLAY</span>
            <select value={selectedDisplay} onChange={(event) => onDisplayChange(event.currentTarget.value)}>
              <option value="automatic">Automatic</option>
              {displays.map((display) => <option value={display.id} key={display.id}>{display.label}</option>)}
            </select>
          </label>
          <button className="export-service-button" onClick={onExportService} title="Download a complete JSON backup">Export backup</button>
          {isPresenting ? (
            <button className="present-button present-button--stop" onClick={onClosePresentation}>End</button>
          ) : (
            <button className="present-button" onClick={onOpenPresentation}><span className="present-button__icon">▶</span> Present</button>
          )}
          <button className={`stage-output-button${isStageOpen ? " is-open" : ""}`} onClick={isStageOpen ? onCloseStage : onOpenStage}>
            {isStageOpen ? "Close stage" : "Stage display"}
          </button>
        </div>
      </header>

      <div
        className="workspace service-workspace"
        style={{ gridTemplateColumns: `${panelWidths.sidebar}px 6px minmax(320px, 1fr) 6px ${panelWidths.preview}px` }}
        onPointerMove={(event) => {
          if (!draggingPanel.current) return;
          const bounds = event.currentTarget.getBoundingClientRect();
          setPanelWidths((current) => draggingPanel.current === "sidebar"
            ? { ...current, sidebar: Math.max(220, Math.min(420, event.clientX - bounds.left)) }
            : { ...current, preview: Math.max(280, Math.min(480, bounds.right - event.clientX)) });
        }}
        onPointerUp={() => { draggingPanel.current = null; }}
        onPointerCancel={() => { draggingPanel.current = null; }}
      >
        <aside className="slide-sidebar service-sidebar">
          <div className="sidebar-heading service-sidebar-heading">
            <div>
              <span className="eyebrow">SERVICE ORDER</span>
              <h1>{service.items.length} {service.items.length === 1 ? "item" : "items"}</h1>
            </div>
            <span className="slide-count">{sequence.length}</span>
          </div>
          <div className="service-tree" aria-label="Service items and slides">
            {service.items.map((item) => {
              const expanded = expandedItems.has(item.id);
              const isLiveItem = item.id === liveSlide?.itemId;
              const orderedSlides = getItemSequenceSlides(item);
              return (
                <section
                  className={`service-item${isLiveItem ? " service-item--live" : ""}${draggedItemId === item.id ? " is-dragging" : ""}`}
                  key={item.id}
                  onDragOver={(event) => {
                    if (draggedItemId) event.preventDefault();
                  }}
                  onDrop={(event) => {
                    event.preventDefault();
                    const sourceItemId = event.dataTransfer.getData("text/plain") || draggedItemId;
                    reorderItem(sourceItemId, item.id);
                    setDraggedItemId("");
                    setExpandedItems(new Set(service.items.map((entry) => entry.id)));
                  }}
                  onDragEnd={() => {
                    setDraggedItemId("");
                    setExpandedItems(new Set(service.items.map((entry) => entry.id)));
                  }}
                >
                  <div
                    className={`service-item-heading${selectedItemId === item.id ? " is-selected" : ""}`}
                    onDragOver={(event) => {
                      if (draggedItemId) event.preventDefault();
                    }}
                    onDrop={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      const sourceItemId = event.dataTransfer.getData("text/plain") || draggedItemId;
                      reorderItem(sourceItemId, item.id);
                      setDraggedItemId("");
                      setExpandedItems(new Set(service.items.map((entry) => entry.id)));
                    }}
                  >
                    <button
                      className="service-item-drag"
                      draggable
                      title="Drag to reorder service item"
                      aria-label={`Drag ${item.title} to reorder`}
                      onDragStart={(event) => {
                        event.dataTransfer.effectAllowed = "move";
                        event.dataTransfer.setData("text/plain", item.id);
                        setDraggedItemId(item.id);
                        setExpandedItems(new Set());
                      }}
                      onDragEnd={() => {
                        setDraggedItemId("");
                        setExpandedItems(new Set(service.items.map((entry) => entry.id)));
                      }}
                    >⋮</button>
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
                    >{expanded ? "▾" : "▸"}</button>
                    <button
                      className="item-title-button"
                      onClick={() => {
                        setSelectedItemId(item.id);
                        setSelectedSlideId(item.slides[0]?.id ?? "");
                        setViewMode(item.type === "bible" ? "bible" : "overview");
                        if (item.type === "song" && item.slides.length === 0) { setSongTargetItemId(item.id); setSongDialogOpen(true); setUtilityTab("songs"); }
                      }}
                    >
                      <span className="item-type">{item.type.toUpperCase()}</span>
                      <strong>{item.title}</strong>
                    </button>
                    {isLiveItem && <span className="item-live-tag"><i /> LIVE</span>}
                    <button className="item-action-button" aria-label={`Rename ${item.title}`} title="Rename" onClick={() => {
                      setEditingItemId(item.id);
                      setEditingItemTitle(item.title);
                    }}>Rename</button>
                    <button className="item-action-button item-action-button--delete" aria-label={`Delete ${item.title}`} title="Delete" onClick={() => removeItem(item.id)}>Delete</button>
                  </div>
                  <div
                    className="service-item-drop-target"
                    onDragOver={(event) => {
                      if (draggedItemId) event.preventDefault();
                    }}
                    onDrop={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      const sourceItemId = event.dataTransfer.getData("text/plain") || draggedItemId;
                      reorderItem(sourceItemId, item.id);
                      setDraggedItemId("");
                      setExpandedItems(new Set(service.items.map((entry) => entry.id)));
                    }}
                  />
                  {editingItemId === item.id && (
                    <form className="item-rename-form" onSubmit={(event) => renameItem(event, item.id)}>
                      <input autoFocus aria-label="Service item name" value={editingItemTitle} onChange={(event) => setEditingItemTitle(event.currentTarget.value)} />
                      <button type="submit" aria-label="Save name">✓</button>
                      <button type="button" aria-label="Cancel rename" onClick={() => setEditingItemId("")}>×</button>
                    </form>
                  )}
                  {expanded && (
                    <div className="service-item-slides">
                      {orderedSlides.map((slide, slideIndex) => {
                        const slideIsLive = liveSlide?.itemId === item.id && liveSlide.itemSlidePosition === slideIndex + 1;
                        const slideIsNext = sequence[liveSequenceIndex + 1]?.itemId === item.id && sequence[liveSequenceIndex + 1]?.itemSlidePosition === slideIndex + 1;
                        return (
                          <button
                            key={slide.id}
                            className={`service-slide${slide.id === selectedSlideId && item.id === selectedItemId ? " is-preview" : ""}${slideIsLive ? " is-live" : ""}${slideIsNext ? " is-next" : ""}`}
                            draggable={!item.arrangements?.length}
                            onClick={() => { setSelectedItemId(item.id); setSelectedSlideId(slide.id); setViewMode(item.type === "bible" ? "bible" : "overview"); }}
                            onDoubleClick={() => goLiveSlide(item.id, slide, slideIndex + 1)}
                            onContextMenu={(event) => openSlideContextMenu(event, item.id, slide.id)}
                            onDragStart={(event) => {
                              event.dataTransfer.effectAllowed = "move";
                              event.dataTransfer.setData("text/plain", slide.id);
                              draggedSlideRef.current = { itemId: item.id, slideId: slide.id };
                            }}
                            onDragOver={(event) => {
                              if (draggedSlideRef.current?.itemId === item.id) event.preventDefault();
                            }}
                            onDrop={(event) => {
                              event.preventDefault();
                              event.stopPropagation();
                              reorderSlide(item.id, slide.id);
                            }}
                            onDragEnd={() => { draggedSlideRef.current = null; }}
                          >
                            <SlideCanvas slide={slide} screenMode="slide" variant="thumbnail" />
                            <span className="service-slide__index">{String(slideIndex + 1).padStart(2, "0")}</span>
                            <span className="service-slide__text">
                              {slide.sectionLabel && <strong>{slide.sectionLabel} · </strong>}
                              {slide.text.replace(/\s+/g, " ").trim() || "Empty slide"}
                            </span>
                            {slideIsLive && <span className="service-slide__live-dot" title="Currently live" />}
                            {slideIsNext && <span className="service-slide__next-tag">NEXT</span>}
                            {(slide.cueMacroIds?.length ?? 0) > 0 && <span className="slide-cue-badge" title={`${slide.cueMacroIds?.length} automatic cue(s)`}>↯ {slide.cueMacroIds?.length}</span>}
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

        <div
          className="workspace-resize-handle"
          role="separator"
          aria-orientation="vertical"
          aria-label="Resize service playlist panel"
          tabIndex={0}
          onPointerDown={(event) => { draggingPanel.current = "sidebar"; event.currentTarget.setPointerCapture(event.pointerId); }}
          onKeyDown={(event) => {
            if (event.key === "ArrowLeft" || event.key === "ArrowRight") {
              event.preventDefault();
              setPanelWidths((current) => ({ ...current, sidebar: Math.max(220, Math.min(420, current.sidebar + (event.key === "ArrowRight" ? 16 : -16))) }));
            }
          }}
        />

        <section className={`editor-pane service-center-pane${viewMode === "editor" ? " is-slide-editor" : ""}`} aria-label="Production canvas">
          <div className="console-canvas-toolbar"><span className="canvas-toolbar-status">{viewMode === "editor" ? `Editing ${selectedItem?.title ?? "slide"}` : selectedItem?.title ?? "Service overview"}</span></div>
          {viewMode === "bible" && selectedItem?.type === "bible" ? (
            <ScripturePanel onAddToService={addScriptureToService} onGoLive={goLiveScripture} />
          ) : selectedItem?.type === "song" && selectedItem.slides.length === 0 ? (
            <div className="empty-editor-wrap"><div className="empty-editor-copy">Choose a song to add its lyrics to this service item.</div><button className="add-slide-button" onClick={() => { setSongTargetItemId(selectedItem.id); setSongDialogOpen(true); setUtilityTab("songs"); }}>Choose from song library</button></div>
          ) : viewMode === "overview" && selectedItem ? (
            <ServiceItemOverview
              item={selectedItem}
              selectedSlideId={selectedSlideId}
              liveItemPosition={liveSlide?.itemId === selectedItem.id ? liveSlide.itemSlidePosition : null}
              nextItemPosition={sequence[liveSequenceIndex + 1]?.itemId === selectedItem.id ? sequence[liveSequenceIndex + 1].itemSlidePosition : null}
              size={overviewSize}
              onSizeChange={setOverviewSize}
              onSelectSlide={setSelectedSlideId}
              onEditSlide={(slideId) => { setSelectedSlideId(slideId); setViewMode("editor"); }}
              onGoLive={(slide, itemPosition) => goLiveSlide(selectedItem.id, slide, itemPosition)}
              onAddSlide={() => addSlide(selectedItem.id)}
              onItemChange={(update) => changeItem(selectedItem.id, update)}
            />
          ) : viewMode === "overview" ? (
            <div className="empty-editor-wrap">
              <div className="empty-editor-copy">Choose a service item to see its slides.</div>
            </div>
          ) : selectedSlide ? (
            <>
              <div className="pane-heading">
                <div>
                  <span className="eyebrow">SLIDE EDITOR · {selectedItem?.title ?? "NO ITEM"}</span>
                  <h2>{selectedSlide.sectionLabel || `Slide ${selectedItemSlideIndex + 1}`}</h2>
                </div>
                <button className="return-overview-button" onClick={() => setViewMode("overview")}>← Item overview</button>
              </div>
              <div className="slide-object-toolbar" aria-label="Add slide object">
                <span>CANVAS</span>
                <button onClick={() => addSlideObject("text")}>+ Text</button>
                <button onClick={() => addSlideObject("rectangle")}>+ Rectangle</button>
                <button onClick={() => addSlideObject("ellipse")}>+ Circle</button>
                <span className="slide-object-toolbar__hint">Drag an object to move it; use its corner to resize.</span>
              </div>
              <div className="slide-object-canvas-wrap">
                <SlideCanvas slide={selectedSlide} screenMode="slide" variant="stage" interactive selectedObjectId={selectedObjectId} onObjectSelect={setSelectedObjectId} onObjectChange={changeSlideObject} />
              </div>
              {selectedObject && <section className="slide-object-inspector" aria-label="Selected object properties">
                <div className="slide-object-inspector__heading"><strong>{selectedObject.kind[0].toUpperCase() + selectedObject.kind.slice(1)} object</strong><button className="danger-action" onClick={() => {
                  changeSlide({ objects: (selectedSlide.objects ?? []).filter((object) => object.id !== selectedObject.id) });
                  setSelectedObjectId(null);
                }}>Delete object</button></div>
                <div className="slide-object-inspector__layer-actions">
                  <button onClick={() => {
                    const objects = [...(selectedSlide.objects ?? [])];
                    const index = objects.findIndex((object) => object.id === selectedObject.id);
                    if (index >= 0 && index < objects.length - 1) [objects[index], objects[index + 1]] = [objects[index + 1], objects[index]];
                    changeSlide({ objects });
                  }}>Bring forward</button>
                  <button onClick={() => {
                    const objects = [...(selectedSlide.objects ?? [])];
                    const index = objects.findIndex((object) => object.id === selectedObject.id);
                    if (index > 0) [objects[index - 1], objects[index]] = [objects[index], objects[index - 1]];
                    changeSlide({ objects });
                  }}>Send backward</button>
                </div>
                {selectedObject.kind === "text" && <>
                  <label className="slide-object-inspector__text">Text<textarea value={selectedObject.text} onChange={(event) => changeSlideObject(selectedObject.id, { text: event.currentTarget.value })} rows={2} /></label>
                  <label>Text color<input type="color" value={selectedObject.color} onChange={(event) => changeSlideObject(selectedObject.id, { color: event.currentTarget.value })} /></label>
                  <label>Text size<input type="number" min="8" max="240" value={selectedObject.fontSize} onChange={(event) => changeSlideObject(selectedObject.id, { fontSize: Number(event.currentTarget.value) })} /></label>
                  <label>Align<select value={selectedObject.textAlign} onChange={(event) => changeSlideObject(selectedObject.id, { textAlign: event.currentTarget.value as SlideObject["textAlign"] })}><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option></select></label>
                </>}
                {(selectedObject.kind === "rectangle" || selectedObject.kind === "ellipse") && <>
                  <label>Fill<input type="color" value={selectedObject.fill} onChange={(event) => changeSlideObject(selectedObject.id, { fill: event.currentTarget.value })} /></label>
                  <label>Outline<input type="color" value={selectedObject.stroke} onChange={(event) => changeSlideObject(selectedObject.id, { stroke: event.currentTarget.value, strokeWidth: selectedObject.strokeWidth || 2 })} /></label>
                  <label>Outline width<input type="number" min="0" max="24" value={selectedObject.strokeWidth} onChange={(event) => changeSlideObject(selectedObject.id, { strokeWidth: Number(event.currentTarget.value) })} /></label>
                </>}
                <label>X<input type="number" min="0" max={100 - selectedObject.width} value={Math.round(selectedObject.x)} onChange={(event) => changeSlideObject(selectedObject.id, { x: Number(event.currentTarget.value) })} /></label>
                <label>Y<input type="number" min="0" max={100 - selectedObject.height} value={Math.round(selectedObject.y)} onChange={(event) => changeSlideObject(selectedObject.id, { y: Number(event.currentTarget.value) })} /></label>
                <label>Width<input type="number" min="4" max={100 - selectedObject.x} value={Math.round(selectedObject.width)} onChange={(event) => changeSlideObject(selectedObject.id, { width: Number(event.currentTarget.value) })} /></label>
                <label>Height<input type="number" min="4" max={100 - selectedObject.y} value={Math.round(selectedObject.height)} onChange={(event) => changeSlideObject(selectedObject.id, { height: Number(event.currentTarget.value) })} /></label>
                <label>Rotation<input type="number" min="-180" max="180" value={selectedObject.rotation} onChange={(event) => changeSlideObject(selectedObject.id, { rotation: Number(event.currentTarget.value) })} /></label>
                <label className="slide-object-inspector__opacity">Opacity<input type="range" min="0" max="1" step="0.05" value={selectedObject.opacity} onChange={(event) => changeSlideObject(selectedObject.id, { opacity: Number(event.currentTarget.value) })} /></label>
              </section>}
              <textarea
                className="slide-text-input"
                aria-label="Slide text"
                value={selectedSlide.text}
                onChange={(event) => changeSlide({ text: event.currentTarget.value })}
                placeholder="Enter slide text"
                spellCheck
              />
              {selectedItem?.type === "song" && (
                <label className="song-section-editor">
                  <span>SECTION</span>
                  <input
                    value={selectedSlide.sectionLabel ?? ""}
                    onChange={(event) => changeSlide({ sectionLabel: event.currentTarget.value || null })}
                    placeholder="Verse 1, Chorus, Bridge..."
                  />
                </label>
              )}
              <label className="speaker-notes-editor">
                <span>STAGE NOTES · OPERATOR ONLY</span>
                <textarea
                  value={selectedSlide.notes ?? ""}
                  onChange={(event) => changeSlide({ notes: event.currentTarget.value || null })}
                  placeholder="Add speaker cues or notes for the stage display..."
                  rows={2}
                />
              </label>
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
                <button onClick={duplicateSlide}>⧉ Duplicate</button>
                <button className="danger-action" onClick={removeSlide}>× Delete</button>
              </div>
              <div className="slide-transition-editor">
                <label>
                  <span>ON GO LIVE</span>
                  <select value={selectedSlide.transition ?? "cut"} onChange={(event) => changeSlide({ transition: event.currentTarget.value as SlideTransition })}>
                    <option value="cut">Cut</option>
                    <option value="fade">Fade</option>
                    <option value="dissolve">Dissolve</option>
                    <option value="push">Push</option>
                  </select>
                </label>
                <label>
                  <span>DURATION <strong>{((selectedSlide.transitionDuration ?? 300) / 1000).toFixed(1)}s</strong></span>
                  <input type="range" min="100" max="2000" step="100" value={selectedSlide.transitionDuration ?? 300} onChange={(event) => changeSlide({ transitionDuration: Number(event.currentTarget.value) })} disabled={(selectedSlide.transition ?? "cut") === "cut"} />
                </label>
              </div>
              <section className="slide-cue-editor" aria-label="Slide cue actions">
                <div className="slide-cue-editor__heading"><strong>When this slide goes live</strong><span>{(selectedSlide.cueMacroIds ?? []).length} cues</span></div>
                <div className="slide-cue-editor__add">
                  <select value={cueMacroSelection} onChange={(event) => setCueMacroSelection(event.currentTarget.value)} aria-label="Choose a macro to run when this slide goes live">
                    <option value="">Choose a saved macro</option>
                    {macros.map((macro) => <option key={macro.id} value={macro.id}>{macro.name}</option>)}
                  </select>
                  <button disabled={!cueMacroSelection || (selectedSlide.cueMacroIds ?? []).includes(cueMacroSelection)} onClick={() => changeSlide({ cueMacroIds: [...(selectedSlide.cueMacroIds ?? []), cueMacroSelection] })}>Add cue</button>
                </div>
                <div className="slide-cue-editor__list">
                  {(selectedSlide.cueMacroIds ?? []).map((macroId) => {
                    const macro = macros.find((entry) => entry.id === macroId);
                    return <span key={macroId}>{macro?.name ?? "Missing macro"}<button aria-label={`Remove ${macro?.name ?? "macro"} cue`} onClick={() => changeSlide({ cueMacroIds: (selectedSlide.cueMacroIds ?? []).filter((id) => id !== macroId) })}>?</button></span>;
                  })}
                  {macros.length === 0 && <small>Create a macro in the Macros tab first.</small>}
                </div>
              </section>

              <section className="slide-theme-panel" aria-label="Slide themes">
                <div className="slide-theme-panel__heading">
                  <span>THEMES</span>
                  <small>{slideThemes.length} SAVED</small>
                </div>
                <div className="slide-theme-panel__actions">
                  <select value={selectedThemeId} onChange={(event) => setSelectedThemeId(event.currentTarget.value)} aria-label="Select slide theme">
                    <option value="">Choose a saved theme</option>
                    {slideThemes.map((theme) => <option value={theme.id} key={theme.id}>{theme.name}</option>)}
                  </select>
                  <button type="button" disabled={!selectedTheme || !selectedSlide} onClick={() => selectedTheme && applyThemeToSlide(selectedTheme)}>Apply to slide</button>
                  <button type="button" disabled={!selectedTheme || service.items.length === 0} onClick={() => selectedTheme && applyThemeToService(selectedTheme)}>Apply to service</button>
                  <button type="button" disabled={!selectedTheme || !selectedSlide} onClick={() => void updateThemeFromCurrentSlide()}>Update theme</button>
                  <button type="button" className="danger-action" disabled={!selectedTheme} onClick={() => selectedTheme && void onSlideThemeDelete(selectedTheme.id).then(() => setThemeError("")).catch((error: unknown) => setThemeError(error instanceof Error ? error.message : String(error)))}>Delete</button>
                </div>
                <form className="slide-theme-save" onSubmit={(event) => void saveCurrentTheme(event)}>
                  <input maxLength={48} value={newThemeName} onChange={(event) => setNewThemeName(event.currentTarget.value)} placeholder="Save selected slide style as a theme..." aria-label="New theme name" />
                  <button type="submit" disabled={!newThemeName.trim() || !selectedSlide}>Save</button>
                </form>
                {themeError && <p className="media-library-error" role="alert">{themeError}</p>}
              </section>
              <section className="slide-template-panel" aria-label="Slide templates">
                <div className="slide-template-panel__heading"><span>TEMPLATES</span><small>{slideTemplates.length} SAVED</small></div>
                <div className="slide-template-panel__select">
                  <select value={selectedTemplateId} onChange={(event) => setSelectedTemplateId(event.currentTarget.value)} aria-label="Choose a slide template">
                    {slideTemplates.map((template) => <option key={template.id} value={template.id}>{template.name}</option>)}
                  </select>
                  <button type="button" className="danger-action" disabled={!selectedTemplate} onClick={() => {
                    const updated = deleteSlideTemplate(selectedTemplateId);
                    setSlideTemplates(updated);
                    setSelectedTemplateId(updated[0]?.id ?? "");
                  }}>Delete</button>
                </div>
                {selectedTemplate && <>
                  <div className="slide-template-panel__fields">
                    {templateFields.map((field) => <label key={field}><span>{field.replace(/[-_]/g, " ").toUpperCase()}</span><input value={templateValues[field] ?? ""} onChange={(event) => setTemplateValues((current) => ({ ...current, [field]: event.currentTarget.value }))} placeholder={`Enter ${field.replace(/[-_]/g, " ")}`} /></label>)}
                  </div>
                  <button type="button" className="slide-template-panel__apply" disabled={!selectedItem || (templateFields.length > 0 && templateFields.some((field) => !templateValues[field]?.trim()))} onClick={applySlideTemplate}>ADD SLIDE FROM TEMPLATE</button>
                </>}
                <details className="slide-template-panel__editor">
                  <summary>Save a custom template</summary>
                  <form onSubmit={saveCurrentAsTemplate}>
                    <input value={templateName} onChange={(event) => setTemplateName(event.currentTarget.value)} maxLength={48} placeholder="Template name" aria-label="Template name" />
                    <textarea value={templateDraft} onChange={(event) => setTemplateDraft(event.currentTarget.value)} rows={3} aria-label="Template content" />
                    <small>Use placeholders such as {"{{title}}"}, {"{{body}}"}, {"{{reference}}"}, and {"{{translation}}"}.</small>
                    <button type="submit" disabled={!selectedSlide || !templateName.trim()}>SAVE TEMPLATE STYLE</button>
                  </form>
                </details>
                {templateError && <p className="media-library-error" role="alert">{templateError}</p>}
              </section>
              <div className="media-editor-strip">
                <div><span className="eyebrow">BACKGROUND MEDIA</span><strong>{selectedSlide.backgroundMedia ? "Media attached" : "Color background"}</strong></div>
                <label className="media-upload-button">
                  + Upload media
                  <input type="file" accept="image/*,video/*" onChange={uploadSlideMedia} />
                </label>
                {selectedSlide.backgroundMedia && <button className="small-text-button" onClick={() => changeSlide({ backgroundMedia: null })}>Remove</button>}
              </div>
            </>
          ) : (
            <div className="empty-editor-wrap">
              <div className="empty-editor-copy">{selectedItem ? "This item has no slides." : "Choose a service item to edit."}</div>
              {selectedItem && <button className="add-slide-button" onClick={() => addSlide(selectedItem.id)}><span>+</span> Add slide</button>}
            </div>
          )}
        </section>

        <div
          className="workspace-resize-handle"
          role="separator"
          aria-orientation="vertical"
          aria-label="Resize output and media panel"
          tabIndex={0}
          onPointerDown={(event) => { draggingPanel.current = "preview"; event.currentTarget.setPointerCapture(event.pointerId); }}
          onKeyDown={(event) => {
            if (event.key === "ArrowLeft" || event.key === "ArrowRight") {
              event.preventDefault();
              setPanelWidths((current) => ({ ...current, preview: Math.max(280, Math.min(480, current.preview + (event.key === "ArrowLeft" ? 16 : -16))) }));
            }
          }}
        />

        <aside className="preview-pane v03-preview-pane" aria-label="Live and preview slides">
          <div className="pane-heading preview-heading">
            <div><span className="eyebrow">LIVE PREVIEW</span><h2>{liveSlide?.itemTitle ?? "No live slide"}</h2></div>
            <span className={`output-state${isPresenting ? " is-live" : ""}`}><i />{isPresenting ? "LIVE" : "OFFLINE"}</span>
          </div>
          <div className="preview-frame live-output-frame">
            {liveSlide
              ? <SlideCanvas slide={liveSlide.slide} screenMode={screenMode} variant="preview" clearSlide={clearForeground} clearBackground={clearBackground} />
              : <div className="preview-empty">Nothing is live</div>}
            {screenMode === "slide" && liveMedia?.kind === "video" && !clearForeground && <video ref={previewVideoRef} className="presentation-live-media presentation-live-media--preview" src={liveMedia.source} muted playsInline />}
            {liveMessage && !clearForeground && <div className="presentation-message presentation-message--preview">{liveMessage}</div>}
          </div>
          <div className="live-preview-detail">
            <strong>{liveSlide ? `${liveSlide.itemTitle} · ${liveSlide.itemSlidePosition} / ${liveSlide.itemSlideTotal}` : "No live slide"}</strong>
            <span>{liveSlide ? `Overall ${liveSlide.overallPosition} / ${liveSlide.overallTotal}` : "Go live from the service outline"}</span>
          </div>
          <div className="screen-mode-controls screen-mode-controls--preview">
            <button className={screenMode === "black" ? "is-active" : ""} onClick={() => onScreenMode(screenMode === "black" ? "slide" : "black")}>BLACK</button>
            <button className={screenMode === "white" ? "is-active" : ""} onClick={() => onScreenMode(screenMode === "white" ? "slide" : "white")}>WHITE</button>
          </div>
          <div className="clear-layer-controls" aria-label="Clear output">
            <button title="Alt+Backspace" onClick={onClearForeground} className={clearForeground ? "is-cleared" : ""}>CLEAR FOREGROUND</button>
            <button title="Alt+B" onClick={onClearBackground} className={clearBackground ? "is-cleared" : ""}>CLEAR BACKGROUND</button>
            <button title="Alt+X" onClick={onClearAll}>CLEAR ALL</button>
          </div>
          
          
          
          <button className="go-live-button preview-go-live" disabled={!previewSlide || isPreviewLive} onClick={() => previewSlide && onGoLive(previewSlide)}>
            {isPreviewLive ? "LIVE SLIDE" : "GO LIVE"}
          </button>
          <div className="preview-footer">
            <span className="preview-resolution">16:9</span>
            <span className="preview-output"><span className={isPresenting ? "is-live" : ""} />{isPresenting ? "Output connected" : "Output idle"}</span>
          </div>
        </aside>
      </div>

      {utilityTab && <section className="utility-dock-panel" aria-label={`${utilityTab} tools`}>
        <div className="utility-dock-panel__heading"><strong>{utilityTab[0].toUpperCase() + utilityTab.slice(1)}</strong><button onClick={() => setUtilityTab(null)} aria-label="Close tools">Close</button></div>
        {utilityTab === "songs" && <p>Choose a song from your library or create a new one.</p>}
        {utilityTab === "media" && liveMedia?.kind === "video" && <section className="video-playback-control" aria-label="Video playback controls">
          <div className="video-playback-control__heading"><strong>{liveMedia.name}</strong><span>{formatTime(videoProgress.currentTime)} / {formatTime(videoProgress.duration)}</span></div>
          <input className="video-playback-control__seek" type="range" min="0" max={Math.max(1, videoProgress.duration)} step="0.1" value={Math.min(videoProgress.currentTime, videoProgress.duration || 0)} onChange={(event) => onSeekVideo(Number(event.currentTarget.value))} aria-label="Seek video" />
          <div className="video-playback-control__row">
            <button className="video-playback-control__play" onClick={() => onVideoPlaybackChange({ playing: !videoPlayback.playing })}>{videoPlayback.playing ? "Pause" : "Play"}</button>
            <button onClick={() => onSeekVideo(Math.max(0, videoProgress.currentTime - 5))} title="Back five seconds">?5s</button>
            <button onClick={() => onSeekVideo(Math.min(videoProgress.duration, videoProgress.currentTime + 5))} title="Forward five seconds">+5s</button>
            <label>Volume <input type="range" min="0" max="1" step="0.05" value={videoPlayback.volume} onChange={(event) => onVideoPlaybackChange({ volume: Number(event.currentTarget.value) })} /></label>
            <label>Speed <select value={videoPlayback.playbackRate} onChange={(event) => onVideoPlaybackChange({ playbackRate: Number(event.currentTarget.value) })}><option value="0.5">0.5?</option><option value="0.75">0.75?</option><option value="1">1?</option><option value="1.25">1.25?</option><option value="1.5">1.5?</option><option value="2">2?</option></select></label>
            <label className="video-playback-control__check"><input type="checkbox" checked={videoPlayback.loop} onChange={(event) => onVideoPlaybackChange({ loop: event.currentTarget.checked })} /> Loop</label>
            <label>At end <select value={videoPlayback.endBehavior} onChange={(event) => onVideoPlaybackChange({ endBehavior: event.currentTarget.value as VideoPlayback["endBehavior"] })}><option value="stop">Stop</option><option value="advance">Next slide</option></select></label>
          </div>
        </section>}
        {utilityTab === "media" && (<><div
            className="media-library-panel"
            onDragOver={(event) => {
              if (event.dataTransfer.types.includes("Files")) {
                event.preventDefault();
                event.dataTransfer.dropEffect = "copy";
              }
            }}
            onDrop={(event) => {
              if (!event.dataTransfer.files.length) return;
              event.preventDefault();
              Array.from(event.dataTransfer.files).forEach(importMediaFile);
            }}
          >
            <div className="media-library-heading"><span><i /> MEDIA LIBRARY <small>{mediaAssets.length}</small></span><label className="media-upload-button media-upload-button--small">+ Import<input type="file" accept="image/*,video/*,audio/*" multiple onChange={uploadSlideMedia} /></label></div>
            <div className="media-library-search">
              <span aria-hidden="true">⌕</span>
              <input value={mediaSearch} onChange={(event) => setMediaSearch(event.currentTarget.value)} placeholder="Search media..." aria-label="Search media library" />
              {mediaSearch && <button onClick={() => setMediaSearch("")} aria-label="Clear media search">×</button>}
            </div>
            <div className="media-library-filters" role="tablist" aria-label="Media types">
              {(["all", "image", "video", "audio"] as const).map((kind) => (
                <button key={kind} role="tab" aria-selected={mediaKindFilter === kind} className={mediaKindFilter === kind ? "is-active" : ""} onClick={() => setMediaKindFilter(kind)}>
                  {kind === "all" ? "ALL" : kind === "image" ? "IMAGES" : kind === "video" ? "VIDEOS" : "AUDIO"} <small>{kind === "all" ? mediaAssets.length : mediaAssets.filter((asset) => asset.kind === kind).length}</small>
                </button>
              ))}
            </div>
            <div className="media-library-grid">
              {filteredMediaAssets.map((asset) => {
                const isVideo = asset.kind === "video";
                const isAudio = asset.kind === "audio";
                return (
                  <div className="media-library-entry" key={asset.id}>
                    <button
                      className={`media-library-tile${selectedSlide?.backgroundMedia === asset.source ? " is-selected" : ""}`}
                      onClick={() => selectedSlide && !isAudio && changeSlide({ backgroundMedia: asset.source })}
                      disabled={!selectedSlide || isAudio}
                      title={isAudio ? "Audio asset" : selectedSlide ? `Apply ${asset.name} as this slide's background` : "Select a slide to apply this background"}
                    >
                      {isVideo ? <video src={asset.source} muted preload="metadata" aria-hidden="true" /> : isAudio ? <span className="media-library-audio-icon">♫</span> : <img src={asset.source} alt="" />}
                      <span>{asset.name}</span>
                      <small>{asset.kind.toUpperCase()} · {formatFileSize(asset.size)}</small>
                    </button>
                    {isVideo && <button className="media-library-go-live" onClick={() => onGoLiveMedia(asset)} title={`Play ${asset.name} on audience output`}>▶ Go Live</button>}
                    {isAudio && <button className="media-library-go-live" onClick={() => onGoLiveAudio(asset)} title={`Play ${asset.name} through this computer`}>♫ Play Audio</button>}
                    {!isAudio && <button className="media-library-go-live" onClick={() => addMediaObject(asset)} disabled={!selectedSlide} title="Place this media on the selected slide">Place on slide</button>}
                    <button className="media-library-delete" onClick={() => void onMediaDelete(asset.id).catch((error: unknown) => setMediaError(error instanceof Error ? error.message : String(error)))} title={`Remove ${asset.name} from the library`} aria-label={`Remove ${asset.name} from the library`}>×</button>
                  </div>
                );
              })}
              {filteredMediaAssets.length === 0 && <span className="media-library-empty">{mediaAssets.length === 0 ? "Import images, video, and audio to build your shared media bin." : "No media matches this view."}</span>}
            </div>
            {liveAudio && <div className="media-live-audio"><span>♫ NOW PLAYING · {liveAudio.name}</span><audio key={liveAudio.id} controls autoPlay src={liveAudio.source} aria-label={`Playback controls for ${liveAudio.name}`} /><button onClick={onClearAudio}>Stop audio</button></div>}
            {mediaError && <p className="media-library-error" role="alert">{mediaError}</p>}
          </div></>)}
        {utilityTab === "messages" && (<><form className="live-message-control" onSubmit={(event) => {
            event.preventDefault();
            const form = event.currentTarget;
            const value = new FormData(form).get("live-message");
            if (typeof value === "string" && value.trim()) {
              onTriggerMessage(value);
              form.reset();
            }
          }}>
            <label htmlFor="live-message-input">MESSAGE OVERLAY</label>
            <div>
              <input id="live-message-input" name="live-message" maxLength={180} placeholder="Type a message for the room..." />
              <button type="submit">SEND</button>
            </div>
          </form></>)}
        {utilityTab === "timers" && (<><form className="stage-timer-control" onSubmit={(event) => {
            event.preventDefault();
            const duration = Math.max(1, Math.min(999, Number(timerMinutes) || 1));
            onStartTimer(timerLabel.trim() || "Timer", duration * 60);
          }}>
            <div className="stage-timer-control__heading"><label htmlFor="stage-timer-name">STAGE COUNTDOWN</label>{stageTimer && <button type="button" onClick={onClearTimer}>Clear</button>}</div>
            <div className="stage-timer-control__fields">
              <input id="stage-timer-name" maxLength={32} value={timerLabel} onChange={(event) => setTimerLabel(event.currentTarget.value)} aria-label="Timer label" />
              <input type="number" min="1" max="999" value={timerMinutes} onChange={(event) => setTimerMinutes(event.currentTarget.value)} aria-label="Timer duration in minutes" />
              <span>MIN</span>
              <button type="submit">{stageTimer ? "RESTART" : "START"}</button>
            </div>
            {stageTimer && <p>Running on the stage display: <strong>{stageTimer.label}</strong></p>}
          </form></>)}
        {utilityTab === "macros" && (<><section className="macro-control">
            <summary>MACROS <span>{macros.length} SAVED{isMacroRunning ? " · RUNNING" : ""}</span></summary>
            {isMacroRunning && <button className="macro-control__stop" onClick={onStopMacro}>STOP MACRO</button>}
            <form className="macro-control__builder" onSubmit={saveMacroDraft}>
              <input value={macroName} onChange={(event) => setMacroName(event.currentTarget.value)} maxLength={40} placeholder="Macro name" aria-label="Macro name" />
              <div className="macro-control__add-row">
                <select value={macroActionKind} onChange={(event) => { setMacroActionKind(event.currentTarget.value as MacroActionKind); setMacroActionValue(""); }} aria-label="Macro action">
                  {macroActions.map((action) => <option key={action.value} value={action.value}>{action.label}</option>)}
                </select>
                {macroActionKind === "send-message" && <input value={macroActionValue} onChange={(event) => setMacroActionValue(event.currentTarget.value)} placeholder="Message text" aria-label="Macro message text" />}
                {macroActionKind === "play-video" && <select value={macroActionValue} onChange={(event) => setMacroActionValue(event.currentTarget.value)} aria-label="Video to play"><option value="">Choose video</option>{mediaAssets.filter((asset) => asset.kind === "video").map((asset) => <option key={asset.id} value={asset.id}>{asset.name}</option>)}</select>}
                {macroActionKind === "play-audio" && <select value={macroActionValue} onChange={(event) => setMacroActionValue(event.currentTarget.value)} aria-label="Audio to play"><option value="">Choose audio</option>{mediaAssets.filter((asset) => asset.kind === "audio").map((asset) => <option key={asset.id} value={asset.id}>{asset.name}</option>)}</select>}
                <label>WAIT <input type="number" min="0" max="600" step="0.1" value={macroActionDelay} onChange={(event) => setMacroActionDelay(event.currentTarget.value)} aria-label="Delay before macro action in seconds" /> s</label>
                <button type="button" onClick={appendMacroAction} disabled={macroValueKinds.includes(macroActionKind) && !macroActionValue}>ADD STEP</button>
              </div>
              {macroDraftActions.length > 0 && <ol className="macro-control__draft">
                {macroDraftActions.map((action, index) => <li key={action.id}>
                  <span>{macroActions.find((entry) => entry.value === action.kind)?.label}{action.value ? ` · ${action.kind === "play-video" || action.kind === "play-audio" ? mediaAssets.find((asset) => asset.id === action.value)?.name ?? "Missing media" : action.value}` : ""}{action.delayMs > 0 ? ` · wait ${(action.delayMs / 1000).toFixed(1)}s` : ""}</span>
                  <button type="button" onClick={() => moveMacroAction(index, -1)} disabled={index === 0} aria-label="Move action up">↑</button>
                  <button type="button" onClick={() => moveMacroAction(index, 1)} disabled={index === macroDraftActions.length - 1} aria-label="Move action down">↓</button>
                  <button type="button" onClick={() => setMacroDraftActions((current) => current.filter((entry) => entry.id !== action.id))} aria-label="Remove action">×</button>
                </li>)}
              </ol>}
              <button className="macro-control__save" type="submit" disabled={!macroName.trim() || macroDraftActions.length === 0}>{editingMacroId ? "UPDATE MACRO" : "SAVE MACRO"}</button>
            </form>
            <div className="macro-control__saved">
              {macros.map((macro) => <div className="macro-control__saved-item" key={macro.id}>
                <span><strong>{macro.name}</strong><small>{macro.actions.length} actions</small></span>
                <button onClick={() => onRunMacro(macro)} disabled={isMacroRunning} title={`Run ${macro.name}`}>▶ RUN</button>
                <button onClick={() => { setEditingMacroId(macro.id); setMacroName(macro.name); setMacroDraftActions(macro.actions); }} disabled={isMacroRunning} title={`Edit ${macro.name}`}>EDIT</button>
                <button onClick={() => { onDeleteMacro(macro.id); if (editingMacroId === macro.id) { setEditingMacroId(null); setMacroName(""); setMacroDraftActions([]); } }} disabled={isMacroRunning} aria-label={`Delete macro ${macro.name}`} title="Delete macro">×</button>
              </div>)}
              {macros.length === 0 && <span className="utility-empty">Add actions above to build a one-click sequence.</span>}
            </div>
          </section></>)}
      </section>}
      <footer className="bottom-bar service-bottom-bar">
        <nav className="utility-dock-tabs" aria-label="Workspace tools">
          {(["songs", "media", "messages", "timers", "macros"] as const).map((tab) => <button key={tab} className={utilityTab === tab ? "is-active" : ""} onClick={() => {
            if (tab === "songs") { setSongTargetItemId(selectedItem?.type === "song" && selectedItem.slides.length === 0 ? selectedItem.id : null); setSongDialogOpen(true); setUtilityTab("songs"); }
            else setUtilityTab((current) => current === tab ? null : tab);
          }}>{tab === "songs" ? "Song library" : tab[0].toUpperCase() + tab.slice(1)}</button>)}
        </nav>
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
      </footer>
      {recentlyRemovedItem && <div className="undo-toast" role="status">
        <span>Removed “{recentlyRemovedItem.item.title}”</span>
        <button onClick={undoRemoveItem}>Undo</button>
      </div>}
      {slideContextMenu && (
        <div
          className="slide-context-menu"
          style={{ left: slideContextMenu.x, top: slideContextMenu.y }}
          tabIndex={-1}
        >
          <button onClick={() => runSlideContextAction("live")}>▶ Go Live</button>
          <button onClick={() => runSlideContextAction("edit")}>✎ Edit</button>
          <button onClick={() => runSlideContextAction("duplicate")}>⧉ Duplicate</button>
          <button className="danger-action" onClick={() => runSlideContextAction("delete")}>× Delete</button>
        </div>
      )}
      {songDialogOpen && (
        <SongLibraryDialog
          songs={songs}
          onClose={() => setSongDialogOpen(false)}
          onSave={onSongSave}
          onDelete={onSongDelete}
          onAddToService={addSongItem}
          docked
        />
      )}
    </main>
  );
}

