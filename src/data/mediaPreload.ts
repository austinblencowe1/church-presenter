import type { MediaRef, ServiceItem } from '../types/presentation';

export type MediaAsset = {
  id: string;
  url: string;
  kind: 'image' | 'video';
};

function mediaUrl(media: MediaRef | undefined, mediaURLs: Record<string, string>): string | null {
  if (!media) return null;
  if (media.id && mediaURLs[media.id]) return mediaURLs[media.id];
  return media.path || media.source || media.src || null;
}

export function collectMediaAssets(
  slides: ServiceItem[],
  mediaURLs: Record<string, string>
): MediaAsset[] {
  const assets: MediaAsset[] = [];
  const seen = new Set<string>();
  const add = (id: string | undefined, url: string | null, kind: 'image' | 'video') => {
    if (!url || seen.has(url)) return;
    seen.add(url);
    assets.push({ id: id || url, url, kind });
  };

  for (const item of slides) {
    const bg = mediaUrl(item.bgMedia, mediaURLs);
    if (bg) add(item.bgMedia?.id, bg, item.bgMedia?.kind === 'video' ? 'video' : 'image');

    for (const part of item.parts) {
      const bm = mediaUrl(part.bgMedia, mediaURLs);
      if (bm) add(part.bgMedia?.id, bm, part.bgMedia?.kind === 'video' ? 'video' : 'image');

      if (part.videoMedia) {
        const v = mediaUrl(part.videoMedia, mediaURLs);
        if (v) add(part.videoMedia?.id, v, 'video');
      }

      for (const im of part.im || []) {
        const u = im.src || (im.mediaId && mediaURLs[im.mediaId]) || '';
        if (u) add(im.mediaId, u, 'image');
      }

      for (const el of part.elements || []) {
        if (el.kind === 'image') {
          const u = el.src || (el.mediaId && mediaURLs[el.mediaId]) || '';
          if (u) add(el.mediaId, u, 'image');
        }
      }
    }
  }
  return assets;
}

function loadImage(url: string): Promise<void> {
  return new Promise((resolve) => {
    const img = new Image();
    const done = () => {
      img.removeEventListener('load', done);
      img.removeEventListener('error', done);
      resolve();
    };
    img.addEventListener('load', done, { once: true });
    img.addEventListener('error', done, { once: true });
    img.src = url;
    setTimeout(done, 8000);
  });
}

function loadVideo(url: string): Promise<void> {
  return new Promise((resolve) => {
    const v = document.createElement('video');
    v.preload = 'auto';
    v.muted = true;
    v.playsInline = true;
    v.src = url;
    const done = () => {
      v.removeEventListener('loadeddata', done);
      v.removeEventListener('error', done);
      resolve();
    };
    v.addEventListener('loadeddata', done, { once: true });
    v.addEventListener('error', done, { once: true });
    v.load();
    setTimeout(done, 8000);
  });
}

export function preloadAsset(asset: MediaAsset): Promise<void> {
  return asset.kind === 'video' ? loadVideo(asset.url) : loadImage(asset.url);
}

export async function preloadAllMedia(
  slides: ServiceItem[],
  mediaURLs: Record<string, string>,
  onProgress?: (done: number, total: number) => void
): Promise<void> {
  const assets = collectMediaAssets(slides, mediaURLs);
  const total = assets.length;
  let done = 0;
  if (onProgress) onProgress(0, total);
  await Promise.all(
    assets.map(async (asset) => {
      await preloadAsset(asset).catch(() => {});
      done += 1;
      if (onProgress) onProgress(done, total);
    })
  );
}
