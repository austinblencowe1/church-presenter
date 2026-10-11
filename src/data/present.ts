/**
 * Tauri present / audience bridge.
 * Opens WebviewWindow "audience", pushes stage HTML via BroadcastChannel + localStorage.
 */
import type { ServiceItem, SlidePart } from '../types/presentation';
import { buildAudienceStageHtml } from './slideRender';

export const AUDIENCE_LABEL = 'audience';
export const AUDIENCE_BC = 'church-presenter-audience';

let audienceWin: { setFocus?: () => Promise<void>; setFullscreen?: (v: boolean) => Promise<void>; close?: () => Promise<void> } | null =
  null;
let audienceBc: BroadcastChannel | null = null;

function getBc(): BroadcastChannel | null {
  if (audienceBc) return audienceBc;
  try {
    audienceBc = new BroadcastChannel(AUDIENCE_BC);
  } catch {
    audienceBc = null;
  }
  return audienceBc;
}

export async function closeAudienceWindow(): Promise<void> {
  try {
    if (audienceWin?.close) await audienceWin.close();
  } catch {
    /* ignore */
  }
  audienceWin = null;
  try {
    const { WebviewWindow } = await import('@tauri-apps/api/webviewWindow');
    const existing = await WebviewWindow.getByLabel(AUDIENCE_LABEL);
    if (existing) await existing.close();
  } catch {
    /* ignore */
  }
}

function audienceUrl(): string {
  try {
    const u = new URL(window.location.href);
    u.searchParams.set('audience', '1');
    // Prefer dedicated audience.html if present next to index
    if (!u.pathname.endsWith('audience.html')) {
      // In Vite dev: http://localhost:1420/?audience=1
      return u.toString();
    }
    return u.toString();
  } catch {
    return '/?audience=1';
  }
}

export async function openAudienceOnDisplay(displayId?: string): Promise<void> {
  await closeAudienceWindow();

  let WebviewWindow: typeof import('@tauri-apps/api/webviewWindow').WebviewWindow | null = null;
  try {
    const mod = await import('@tauri-apps/api/webviewWindow');
    WebviewWindow = mod.WebviewWindow;
  } catch {
    throw new Error('WebviewWindow API not available. Are you running inside Tauri?');
  }

  let monitors: Awaited<ReturnType<typeof import('@tauri-apps/api/window').availableMonitors>> = [];
  let current: Awaited<ReturnType<typeof import('@tauri-apps/api/window').currentMonitor>> = null;
  try {
    const winApi = await import('@tauri-apps/api/window');
    monitors = await winApi.availableMonitors();
    current = await winApi.currentMonitor();
  } catch {
    /* ignore */
  }

  let target = null as (typeof monitors)[0] | null;
  if (displayId && displayId !== 'default') {
    target =
      monitors.find((m) => String(m.name) === String(displayId)) ||
      monitors.find((_, i) => String(i) === String(displayId)) ||
      null;
  }
  if (!target && monitors.length > 1 && current) {
    target =
      monitors.find((m) => m.name !== current!.name) ||
      monitors.find((m) => {
        try {
          return m.position.x !== current!.position.x || m.position.y !== current!.position.y;
        } catch {
          return false;
        }
      }) ||
      null;
  }
  if (!target) target = monitors[monitors.length > 1 ? 1 : 0] || null;

  const opts: Record<string, unknown> = {
    url: audienceUrl(),
    title: 'Audience',
    decorations: false,
    resizable: false,
    fullscreen: true,
    focus: true,
    visible: true,
    skipTaskbar: true,
  };

  if (target) {
    try {
      const pos =
        typeof target.position?.toLogical === 'function'
          ? target.position.toLogical(target.scaleFactor)
          : {
              x: (target.position?.x || 0) / (target.scaleFactor || 1),
              y: (target.position?.y || 0) / (target.scaleFactor || 1),
            };
      const size =
        typeof target.size?.toLogical === 'function'
          ? target.size.toLogical(target.scaleFactor)
          : {
              width: (target.size?.width || 1920) / (target.scaleFactor || 1),
              height: (target.size?.height || 1080) / (target.scaleFactor || 1),
            };
      opts.x = pos.x;
      opts.y = pos.y;
      opts.width = size.width;
      opts.height = size.height;
    } catch {
      /* ignore geometry */
    }
  }

  try {
    const existing = await WebviewWindow.getByLabel(AUDIENCE_LABEL);
    if (existing) await existing.close();
  } catch {
    /* ignore */
  }

  const audience = new WebviewWindow(AUDIENCE_LABEL, opts as ConstructorParameters<typeof WebviewWindow>[1]);
  await new Promise<void>((resolve, reject) => {
    let settled = false;
    const done = (fn: (v?: unknown) => void, arg?: unknown) => {
      if (settled) return;
      settled = true;
      fn(arg);
    };
    audience.once('tauri://created', () => done(resolve));
    audience.once('tauri://error', (e) =>
      done(
        reject,
        new Error(
          (typeof e?.payload === 'string' ? e.payload : null) ||
            'Failed to create audience window. Check capabilities include windows: ["main","audience"].'
        )
      )
    );
    setTimeout(() => done(resolve), 2500);
  });

  audienceWin = audience;

  const focusAudience = async () => {
    try {
      if (typeof audience.setFocus === 'function') await audience.setFocus();
    } catch {
      /* ignore */
    }
    try {
      if (typeof audience.setFullscreen === 'function') await audience.setFullscreen(true);
    } catch {
      /* ignore */
    }
  };

  await new Promise((r) => setTimeout(r, 150));
  await focusAudience();
  for (const ms of [400, 900, 1600]) {
    setTimeout(() => {
      focusAudience();
    }, ms);
  }
}

export function pushAudienceStage(
  item: ServiceItem | null,
  part: SlidePart | null,
  mediaURLs?: Record<string, string>
): void {
  if (!item || !part) return;
  const html = buildAudienceStageHtml(item, part, mediaURLs);
  try {
    localStorage.setItem('audience-stage-html', html);
    localStorage.setItem('audience-stage-ping', String(Date.now()));
  } catch {
    /* ignore */
  }
  try {
    const bc = getBc();
    if (bc) bc.postMessage({ type: 'stage', html });
  } catch {
    /* ignore */
  }
}

export function pushAudienceClear(mode: '' | 'black' | 'white' | 'logo'): void {
  try {
    localStorage.setItem('audience-clr', mode);
    localStorage.setItem('audience-clr-ping', String(Date.now()));
  } catch {
    /* ignore */
  }
  try {
    const bc = getBc();
    if (bc) bc.postMessage({ type: 'clear', mode });
  } catch {
    /* ignore */
  }
}

export type AudienceControlHandler = (msg: {
  action?: string;
  key?: string;
  t?: number;
}) => void;

let lastControlT = 0;

export function subscribeAudienceControls(handler: AudienceControlHandler): () => void {
  const bc = getBc();
  const onBc = (ev: MessageEvent) => {
    const d = ev?.data;
    if (!d || d.type !== 'control') return;
    const t = d.t || 0;
    if (t && t - lastControlT < 150) return;
    if (t) lastControlT = t;
    handler(d);
  };
  if (bc) bc.addEventListener('message', onBc);

  const poll = window.setInterval(() => {
    try {
      const raw = localStorage.getItem('audience-control');
      if (!raw) return;
      const msg = JSON.parse(raw);
      if (!msg?.t) return;
      if (msg.t - lastControlT < 150) return;
      lastControlT = msg.t;
      handler(msg);
    } catch {
      /* ignore */
    }
  }, 60);

  return () => {
    if (bc) bc.removeEventListener('message', onBc);
    clearInterval(poll);
  };
}
