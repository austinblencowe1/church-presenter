/** Shared presentation models — single source of truth for editor + audience */

export type ItemKind = 'Song' | 'Slide' | 'PowerPoint' | 'PDF' | 'Video' | string;

export interface TextStyle {
  fs: number;
  ff: 'sans' | 'serif';
  tc: 'none' | 'upper' | 'title';
  al: 'left' | 'center' | 'right';
  va: 'top' | 'middle' | 'bottom';
  b: 0 | 1;
  sh: 0 | 1;
  color: string;
}

export interface CanvasElement {
  id: string;
  kind: 'text' | 'image';
  tx?: string;
  src?: string;
  mediaId?: string;
  x: number;
  y: number;
  w: number;
  h: number;
  rot?: number;
  z?: number;
  f?: Partial<TextStyle>;
}

export interface SlideImage {
  full?: boolean;
  src?: string;
  mediaId?: string;
  x?: number;
  y?: number;
  w?: number;
}

export interface SlidePart {
  id: string;
  n: string;
  tx: string;
  im?: SlideImage[];
  elements?: CanvasElement[];
  bgMedia?: MediaRef;
  videoMedia?: MediaRef & { settings?: { autoplay?: boolean; loop?: boolean } };
}

export interface MediaRef {
  id?: string;
  kind?: 'image' | 'video';
  name?: string;
  path?: string;
  source?: string;
  src?: string;
}

export interface ServiceItem {
  t: string;
  k: ItemKind;
  parts: SlidePart[];
  ord: string[];
  bg: number;
  dim: number;
  f: TextStyle;
  bgMedia?: MediaRef;
  sourceFile?: string;
  presentationId?: string;
}

export interface TeamContact {
  name: string;
  email: string;
}

export interface ServiceTeam {
  leader: TeamContact;
  band: TeamContact[];
  sound: TeamContact;
}

export interface Service {
  name: string;
  date?: string;
  tag?: string;
  slides: ServiceItem[];
  team: ServiceTeam;
}

export interface ClearMode {
  mode: '' | 'black' | 'white' | 'logo';
}

export const defaultTextStyle = (): TextStyle => ({
  fs: 5.2,
  ff: 'sans',
  tc: 'none',
  al: 'center',
  va: 'middle',
  b: 1,
  sh: 0,
  color: '#ffffff',
});

export const emptyContact = (): TeamContact => ({ name: '', email: '' });

export const defaultTeam = (): ServiceTeam => ({
  leader: emptyContact(),
  band: [emptyContact(), emptyContact(), emptyContact()],
  sound: emptyContact(),
});

export function uid(): string {
  return 'id-' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

export function clone<T>(x: T): T {
  return JSON.parse(JSON.stringify(x));
}

/** Ordered parts for an item */
export function slideList(it: ServiceItem): SlidePart[] {
  return it.ord
    .map((id) => it.parts.find((p) => p.id === id))
    .filter((p): p is SlidePart => !!p);
}

export function isSongItem(it: ServiceItem | null | undefined): boolean {
  return !!it && it.k === 'Song';
}
