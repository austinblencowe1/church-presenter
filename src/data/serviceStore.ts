/**
 * React state for services + navigation. Replace / merge with existing serviceStore if present.
 */
import { create } from 'zustand';
import {
  type Service,
  type ServiceItem,
  type SlidePart,
  type TextStyle,
  clone,
  defaultTeam,
  defaultTextStyle,
  slideList,
  uid,
} from '../types/presentation';

const STORAGE_KEY = 'service-slides-state';

function mkSlideItem(title: string, lines: string[], bg = 0, fs = 6.5): ServiceItem {
  const parts: SlidePart[] = lines.map((tx, i) => ({
    id: uid(),
    n: 'Slide ' + (i + 1),
    tx,
  }));
  return {
    t: title,
    k: 'Slide',
    parts,
    ord: parts.map((p) => p.id),
    bg,
    dim: 30,
    f: { ...defaultTextStyle(), fs },
  };
}

function defaultServices(): Service[] {
  return [
    {
      name: 'Sunday 10am service',
      tag: 'AM',
      date: new Date().toISOString().slice(0, 10),
      slides: [
        mkSlideItem('Welcome', ['Welcome to church\nWe are glad you are here'], 0, 6.5),
        mkSlideItem('Sermon', ['Sermon Title\nSpeaker Name'], 3, 6.5),
      ],
      team: defaultTeam(),
    },
  ];
}

export interface PresenterState {
  services: Service[];
  activeService: number;
  /** slides of active service */
  slides: ServiceItem[];
  sel: number;
  aj: number;
  live: boolean;
  clear: '' | 'black' | 'white' | 'logo';
  theme: 'system' | 'light' | 'dark';
  selectedDisplay: string;
  mediaURLs: Record<string, string>;
  homeMode: boolean;

  // derived helpers
  currentItem: () => ServiceItem | null;
  currentPart: () => SlidePart | null;

  load: () => void;
  save: () => void;
  setHomeMode: (v: boolean) => void;
  openService: (index: number) => void;
  selectItem: (index: number) => void;
  selectSlide: (itemIndex: number, partIndex: number) => void;
  go: (delta: number) => 'ok' | 'black' | 'end';
  setClear: (mode: '' | 'black' | 'white' | 'logo') => void;
  toggleClear: (mode: 'black' | 'white' | 'logo') => void;
  setLive: (v: boolean) => void;
  setDisplay: (id: string) => void;
  updateItemStyle: (patch: Partial<TextStyle>) => void;
  addBlankSlide: (title?: string) => void;
}

function persist(get: () => PresenterState) {
  const s = get();
  try {
    const payload = {
      services: s.services.map((svc, i) =>
        i === s.activeService ? { ...svc, slides: s.slides } : svc
      ),
      activeService: s.activeService,
      theme: s.theme,
      display: s.selectedDisplay,
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  } catch {
    /* ignore */
  }
}

export const usePresenterStore = create<PresenterState>((set, get) => ({
  services: defaultServices(),
  activeService: 0,
  slides: defaultServices()[0].slides,
  sel: 0,
  aj: 0,
  live: false,
  clear: '',
  theme: 'system',
  selectedDisplay: '',
  mediaURLs: {},
  homeMode: true,

  currentItem: () => {
    const { slides, sel } = get();
    return slides[sel] || null;
  },
  currentPart: () => {
    const it = get().currentItem();
    if (!it) return null;
    const list = slideList(it);
    return list[get().aj] || null;
  },

  load: () => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const saved = JSON.parse(raw);
      if (!saved?.services?.length) return;
      const activeService = Math.min(saved.activeService || 0, saved.services.length - 1);
      const slides = clone(saved.services[activeService].slides);
      set({
        services: saved.services,
        activeService,
        slides,
        sel: 0,
        aj: 0,
        theme: saved.theme || 'system',
        selectedDisplay: saved.display || '',
      });
    } catch {
      /* ignore */
    }
  },

  save: () => persist(get),

  setHomeMode: (v) => set({ homeMode: v }),

  openService: (index) => {
    const { services } = get();
    if (index < 0 || index >= services.length) return;
    const slides = clone(services[index].slides);
    set({
      activeService: index,
      slides,
      sel: 0,
      aj: 0,
      homeMode: false,
      clear: '',
    });
    persist(get);
  },

  selectItem: (index) => {
    const { slides } = get();
    if (index < 0 || index >= slides.length) return;
    set({ sel: index, aj: 0, clear: '' });
  },

  selectSlide: (itemIndex, partIndex) => {
    const { slides } = get();
    const it = slides[itemIndex];
    if (!it) return;
    const list = slideList(it);
    const aj = Math.max(0, Math.min(list.length - 1, partIndex));
    set({ sel: itemIndex, aj, clear: '' });
  },

  go: (delta) => {
    const { slides, sel, aj, live, clear } = get();
    const flat: { itemIndex: number; partIndex: number }[] = [];
    slides.forEach((it, itemIndex) => {
      slideList(it).forEach((_, partIndex) => flat.push({ itemIndex, partIndex }));
    });
    if (!flat.length) return 'ok';
    const k = flat.findIndex((x) => x.itemIndex === sel && x.partIndex === aj);
    if (k < 0) return 'ok';
    const target = k + delta;

    if (delta > 0 && target >= flat.length) {
      if (live) {
        if (clear === 'black') return 'end';
        set({ clear: 'black' });
        return 'black';
      }
      return 'ok';
    }
    if (delta < 0 && live && clear === 'black' && k === flat.length - 1) {
      set({ clear: '' });
      return 'ok';
    }
    if (target < 0 || target >= flat.length) return 'ok';
    const n = flat[target];
    set({ sel: n.itemIndex, aj: n.partIndex, clear: '' });
    return 'ok';
  },

  setClear: (mode) => set({ clear: mode }),
  toggleClear: (mode) => set({ clear: get().clear === mode ? '' : mode }),
  setLive: (v) => set({ live: v, clear: v ? get().clear : '' }),
  setDisplay: (id) => {
    set({ selectedDisplay: id });
    persist(get);
  },

  updateItemStyle: (patch) => {
    const { slides, sel } = get();
    const next = clone(slides);
    const it = next[sel];
    if (!it) return;
    it.f = { ...it.f, ...patch };
    set({ slides: next });
    persist(get);
  },

  addBlankSlide: (title = 'New Slide') => {
    const { slides, sel } = get();
    const part: SlidePart = { id: uid(), n: 'Slide 1', tx: '' };
    const item: ServiceItem = {
      t: title,
      k: 'Slide',
      parts: [part],
      ord: [part.id],
      bg: 0,
      dim: 30,
      f: defaultTextStyle(),
    };
    const next = clone(slides);
    const insertAt = Math.min(sel + 1, next.length);
    next.splice(insertAt, 0, item);
    set({ slides: next, sel: insertAt, aj: 0, clear: '' });
    persist(get);
  },
}));
