export type SlideType = "text";
export type TextAlignment = "left" | "center" | "right";
export type ScreenMode = "slide" | "black" | "white";
export type SlideTransition = "cut" | "fade" | "dissolve" | "push";
export type ServiceItemType = "welcome" | "song" | "bible" | "sermon" | "announcement" | "other";
export type SlideObjectKind = "text" | "rectangle" | "ellipse" | "image" | "video";

export interface SlideObject {
  id: string;
  kind: SlideObjectKind;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  opacity: number;
  fill: string;
  stroke: string;
  strokeWidth: number;
  color: string;
  text: string;
  fontSize: number;
  textAlign: TextAlignment;
  source?: string;
}

export interface Slide {
  id: string;
  type: SlideType;
  text: string;
  background: string;
  textAlign: TextAlignment;
  fontSize: number;
  sectionLabel?: string | null;
  notes?: string | null;
  backgroundMedia?: string | null;
  transition?: SlideTransition;
  transitionDuration?: number;
  themeId?: string | null;
  objects?: SlideObject[];
  cueMacroIds?: string[];
}

export interface Song {
  id: string;
  title: string;
  lyrics: string;
  reflowMode?: "markers" | "blank-lines" | "four-lines";
}

export type MediaKind = "image" | "video" | "audio";

export interface MediaAsset {
  id: string;
  name: string;
  kind: MediaKind;
  source: string;
  size: number;
  addedAt: string;
}

export interface SlideTheme {
  id: string;
  name: string;
  background: string;
  fontSize: number;
  textAlign: TextAlignment;
  transition: SlideTransition;
  transitionDuration: number;
}

export interface SlideTemplate {
  id: string;
  name: string;
  content: string;
  background: string;
  fontSize: number;
  textAlign: TextAlignment;
  transition: SlideTransition;
  transitionDuration: number;
}

export type MacroActionKind = "next-slide" | "previous-slide" | "black-screen" | "white-screen" | "show-slide" | "clear-foreground" | "clear-background" | "clear-all" | "send-message" | "play-video" | "play-audio";

export interface MacroAction {
  id: string;
  kind: MacroActionKind;
  delayMs: number;
  value?: string;
}

export interface ShowMacro {
  id: string;
  name: string;
  actions: MacroAction[];
}

export interface SlideGroup {
  id: string;
  name: string;
  color: string;
  slideIds: string[];
}

export interface Arrangement {
  id: string;
  name: string;
  groupIds: string[];
}

export interface ServiceItem {
  id: string;
  title: string;
  type: ServiceItemType;
  slides: Slide[];
  groups?: SlideGroup[];
  arrangements?: Arrangement[];
  activeArrangementId?: string | null;
  reference?: string;
  bibleVersion?: string;
}

export interface Service {
  id: string;
  title: string;
  date: string;
  folder?: string;
  items: ServiceItem[];
}

export interface ServiceSummary {
  id: string;
  title: string;
  date: string;
  folder?: string;
}

export interface LiveSlide {
  serviceId: string;
  serviceTitle: string;
  itemId: string;
  itemTitle: string;
  itemType: ServiceItemType;
  itemSlidePosition: number;
  itemSlideTotal: number;
  overallPosition: number;
  overallTotal: number;
  slide: Slide;
}

export interface PresentationUpdate {
  liveSlide: LiveSlide | null;
  liveMedia: MediaAsset | null;
  videoPlayback: VideoPlayback;
  screenMode: ScreenMode;
  clearForeground: boolean;
  clearBackground: boolean;
  message: string | null;
}

export interface VideoPlayback {
  playing: boolean;
  volume: number;
  playbackRate: number;
  loop: boolean;
  seekTo: number;
  seekRequest: number;
  endBehavior: "stop" | "advance";
}

export interface VideoProgress {
  currentTime: number;
  duration: number;
}

export interface StagePresentationUpdate {
  current: LiveSlide | null;
  next: LiveSlide | null;
  message: string | null;
  timer: { label: string; endsAt: number } | null;
  audienceDisplayId: string | null;
}
