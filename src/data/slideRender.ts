/**
 * Pure slide → HTML. Used by editor previews AND audience output.
 * Keep this free of React and DOM APIs so both windows stay in sync.
 */
import type { CanvasElement, ServiceItem, SlidePart, TextStyle } from '../types/presentation';
import { defaultTextStyle, isSongItem } from '../types/presentation';

export interface BgDef {
  n: string;
  c: string;
  v?: boolean;
}

export const BACKGROUNDS: BgDef[] = [
  { n: 'Black', c: '#000000' },
  { n: 'Navy', c: '#0b3b6b' },
  { n: 'Teal', c: '#0d4f4f' },
  { n: 'Wave', c: '#0b3b6b', v: true },
  { n: 'Charcoal', c: '#1a1a1a' },
  { n: 'Deep purple', c: '#2a1540' },
];

const VA: Record<string, string> = {
  top: 'flex-start',
  middle: 'center',
  bottom: 'flex-end',
};

function esc(s: string): string {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function styleForText(f: Partial<TextStyle>): string {
  const full = { ...defaultTextStyle(), ...f };
  const fontFamily =
    full.ff === 'serif'
      ? "'Lora',Georgia,'Times New Roman',serif"
      : "'Figtree','Segoe UI',system-ui,sans-serif";
  const textTransform =
    full.tc === 'upper' ? 'uppercase' : full.tc === 'title' ? 'capitalize' : 'none';
  const justifyContent =
    full.al === 'left' ? 'flex-start' : full.al === 'right' ? 'flex-end' : 'center';
  const alignItems = VA[full.va] || 'center';
  const shadow = full.sh ? 'text-shadow:0 .1em .35em rgba(0,0,0,0.85)' : '';
  return [
    `font-size:${full.fs}cqw`,
    `font-family:${fontFamily}`,
    `text-transform:${textTransform}`,
    `text-align:${full.al}`,
    `justify-content:${justifyContent}`,
    `align-items:${alignItems}`,
    `font-weight:${full.b ? 700 : 400}`,
    `color:${full.color || '#ffffff'}`,
    shadow,
  ]
    .filter(Boolean)
    .join(';');
}

/** Ensure non-song parts have elements[] migrated from legacy tx / im */
export function ensureElements(part: SlidePart, item: ServiceItem): CanvasElement[] | null {
  if (isSongItem(item)) return null;
  if (!Array.isArray(part.elements)) {
    part.elements = [];
    if ((part.tx || '').trim()) {
      part.elements.push({
        id: 'el-' + Math.random().toString(36).slice(2, 9),
        kind: 'text',
        tx: part.tx,
        x: 8,
        y: 18,
        w: 84,
        h: 50,
        rot: 0,
        z: 1,
        f: { ...item.f },
      });
    }
    (part.im || []).forEach((m, i) => {
      if (m.full) return;
      part.elements!.push({
        id: 'el-' + Math.random().toString(36).slice(2, 9),
        kind: 'image',
        src: m.src,
        mediaId: m.mediaId,
        x: m.x ?? 50,
        y: m.y ?? 50,
        w: m.w ?? 30,
        h: m.w ?? 30,
        rot: 0,
        z: 2 + i,
      });
    });
  }
  return part.elements;
}

function renderElement(
  el: CanvasElement,
  opts: { editable?: boolean; selectedId?: string | null; mediaURLs?: Record<string, string> }
): string {
  const rot = el.rot || 0;
  const base = `left:${el.x}%;top:${el.y}%;width:${el.w}%;height:${el.h}%;transform:rotate(${rot}deg);z-index:${el.z || 1}`;
  const selected = opts.editable && opts.selectedId === el.id ? ' selected' : '';
  const pe = opts.editable ? 'auto' : 'none';
  const handles =
    opts.editable && opts.selectedId === el.id
      ? `<span class="handle nw" data-h="nw"></span><span class="handle ne" data-h="ne"></span><span class="handle sw" data-h="sw"></span><span class="handle se" data-h="se"></span><span class="handle rot" data-h="rot"></span>`
      : '';

  if (el.kind === 'text') {
    return `<div class="el text-el${selected}" data-el="${esc(el.id)}" style="${base};${styleForText(el.f || {})};pointer-events:${pe}">${handles}<span class="el-tx">${esc(el.tx || '')}</span></div>`;
  }
  const src = el.src || (el.mediaId && opts.mediaURLs?.[el.mediaId]) || '';
  return `<div class="el img-el${selected}" data-el="${esc(el.id)}" style="${base};pointer-events:${pe}">${handles}<img src="${esc(src)}" alt="" draggable="false"></div>`;
}

export interface RenderSlideOptions {
  /** Editor canvas: allow selection chrome */
  editable?: boolean;
  selectedElementId?: string | null;
  /** Live presentation / audience: no chrome */
  stage?: boolean;
  mediaURLs?: Record<string, string>;
}

/**
 * Returns inner HTML for a single `.sl` slide (no outer chrome).
 */
export function renderSlideHtml(
  item: ServiceItem,
  part: SlidePart,
  opts: RenderSlideOptions = {}
): string {
  const bg = BACKGROUNDS[item.bg] || BACKGROUNDS[0];
  const f = item.f || defaultTextStyle();
  const media = part.bgMedia || item.bgMedia;
  const video = media?.kind === 'video';
  const mediaURLs = opts.mediaURLs || {};
  const src =
    (media?.id && mediaURLs[media.id]) || media?.path || media?.source || media?.src || '';

  let background = '';
  if (media && src) {
    if (video) {
      const vkey = media.id || src;
      background = opts.stage
        ? `<video class="bg-video" data-vkey="${esc(vkey)}" src="${esc(src)}" muted loop playsinline autoplay></video>`
        : `<video src="${esc(src)}" muted playsinline preload="metadata"></video>`;
    } else {
      background = `<img src="${esc(src)}" alt="">`;
    }
  }

  const els = ensureElements(part, item);
  let bodyHtml = '';

  if (els && els.length) {
    const sorted = [...els].sort((a, b) => (a.z || 0) - (b.z || 0));
    bodyHtml = sorted
      .map((el) =>
        renderElement(el, {
          editable: !!opts.editable && !opts.stage,
          selectedId: opts.selectedElementId,
          mediaURLs,
        })
      )
      .join('');
  } else {
    // Song / classic single text layer
    bodyHtml = `<div class="tx" style="${styleForText(f)};pointer-events:${opts.editable ? 'auto' : 'none'}">
      <span style="font-family:inherit;text-transform:inherit;text-align:${f.al}">${esc(part.tx)}</span>
    </div>`;
  }

  const fullImgs = (part.im || [])
    .filter((m) => m.full)
    .map(
      (m) =>
        `<img class="stage-image" alt="" src="${esc(m.src || (m.mediaId && mediaURLs[m.mediaId]) || '')}" draggable="false" style="position:absolute;inset:0;width:100%;height:100%;object-fit:contain;pointer-events:none">`
    )
    .join('');

  let slideVideo = '';
  if (part.videoMedia?.id && mediaURLs[part.videoMedia.id]) {
    const vs = part.videoMedia.settings || {};
    slideVideo = `<video class="slide-video${opts.stage ? '' : ' thumbnail-video'}" src="${esc(mediaURLs[part.videoMedia.id])}" ${opts.stage && vs.autoplay ? 'autoplay' : ''} ${opts.stage && vs.loop ? 'loop' : ''} playsinline ${opts.stage ? 'controls' : ''} preload="metadata"></video>`;
  }

  const wave = !media && bg.v ? 'wave' : '';
  const bgStyle = !media && !bg.v ? `background:${bg.c}` : '';

  return `<div class="sl">
    <div class="bgl ${wave}" style="${bgStyle}">${background}</div>
    <div class="dim" style="opacity:${(item.dim || 0) / 100}"></div>
    ${fullImgs}
    ${slideVideo}
    ${bodyHtml}
  </div>`;
}

/** Full stage payload for audience (wrapped in .cv) */
export function buildAudienceStageHtml(
  item: ServiceItem,
  part: SlidePart,
  mediaURLs?: Record<string, string>
): string {
  return `<div class="cv">${renderSlideHtml(item, part, { stage: true, mediaURLs })}</div>`;
}

/** Minimal CSS the audience window needs for freeform + classic slides */
export const AUDIENCE_SLIDE_CSS = `
.sl{position:relative;width:100%;height:100%;overflow:hidden;background:#000;aspect-ratio:16/9;container-type:inline-size}
.bgl,.dim,.tx,.ov{position:absolute;inset:0}
.dim{background:#000;pointer-events:none}
.bgl>img,.bgl>video,.bg-video{width:100%;height:100%;object-fit:cover;display:block}
.tx{display:flex;padding:6% 8%;color:#fff;white-space:pre-line;line-height:1.28;box-sizing:border-box}
.wave{background:linear-gradient(120deg,#0b3b6b,#1b7fa8,#0b3b6b,#143d5f);background-size:300% 300%}
.el{position:absolute;box-sizing:border-box;transform-origin:center center;pointer-events:none;user-select:none}
.el.text-el{display:flex;padding:2%;color:#fff;white-space:pre-wrap;line-height:1.28;overflow:hidden}
.el.img-el img{width:100%;height:100%;object-fit:fill;pointer-events:none;display:block}
.el-tx{width:100%;display:block}
.handle{display:none}
.slide-video{position:absolute;inset:0;width:100%;height:100%;object-fit:contain;background:#000;z-index:4}
`;
