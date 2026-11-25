import { useCallback, useEffect, useState } from 'react';
import { usePresenterStore } from './data/serviceStore';
import {
  closeAudienceWindow,
  openAudienceOnDisplay,
  pushAudienceClear,
  pushAudienceStage,
  subscribeAudienceControls,
} from './data/present';
import { collectMediaAssets, preloadAllMedia } from './data/mediaPreload';
import { TopBar } from './components/TopBar';
import { HomeServices } from './components/HomeServices';
import { EditorLayout } from './components/EditorLayout';
import { PresenterView } from './components/PresenterView';
import './App.css';

function formatServiceTitle(
  svc: { name?: string; tag?: string } | undefined
): string {
  if (!svc) return 'Service';
  const tag = (svc.tag || '').trim();
  const name = (svc.name || '').trim();
  if (tag && name) return `${tag} - ${name}`;
  return tag || name || 'Service';
}

export default function App() {
  const load = usePresenterStore((s) => s.load);
  const homeMode = usePresenterStore((s) => s.homeMode);
  const setHomeMode = usePresenterStore((s) => s.setHomeMode);
  const live = usePresenterStore((s) => s.live);
  const setLive = usePresenterStore((s) => s.setLive);
  const services = usePresenterStore((s) => s.services);
  const activeService = usePresenterStore((s) => s.activeService);
  const selectedDisplay = usePresenterStore((s) => s.selectedDisplay);
  const go = usePresenterStore((s) => s.go);
  const toggleClear = usePresenterStore((s) => s.toggleClear);
  const setClear = usePresenterStore((s) => s.setClear);
  const clear = usePresenterStore((s) => s.clear);
  const currentItem = usePresenterStore((s) => s.currentItem);
  const currentPart = usePresenterStore((s) => s.currentPart);
  const mediaURLs = usePresenterStore((s) => s.mediaURLs);
  const preparing = usePresenterStore((s) => s.preparing);
  const prepProgress = usePresenterStore((s) => s.prepProgress);
  const setPreparing = usePresenterStore((s) => s.setPreparing);
  const setPrepProgress = usePresenterStore((s) => s.setPrepProgress);

  const [presenting, setPresenting] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    load();
  }, [load]);

  const syncAudience = useCallback(() => {
    const item = currentItem();
    const part = currentPart();
    pushAudienceStage(item, part, mediaURLs);
    pushAudienceClear(clear);
  }, [currentItem, currentPart, mediaURLs, clear]);

  useEffect(() => {
    if (!live) return;
    syncAudience();
  }, [live, syncAudience, clear]);

  useEffect(() => {
    if (!live) return;
    return subscribeAudienceControls((msg) => {
      const action = msg.action;
      if (action === 'next') {
        const r = go(1);
        if (r === 'end') void endPresent();
        return;
      }
      if (action === 'prev') {
        go(-1);
        return;
      }
      if (action === 'end') {
        void endPresent();
        return;
      }
      if (action === 'toggle-clear' && msg.key) {
        toggleClear(msg.key as 'black' | 'white' | 'logo');
      }
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [live, go, toggleClear]);

  // Keyboard when main window focused (presenter mode)
  useEffect(() => {
    if (!live) return;
    const onKey = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement | null;
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) return;
      const k = e.key.toLowerCase();
      if (e.key === 'Escape') {
        e.preventDefault();
        void endPresent();
        return;
      }
      if (k === 'b') {
        e.preventDefault();
        toggleClear('black');
        return;
      }
      if (k === 'w') {
        e.preventDefault();
        toggleClear('white');
        return;
      }
      if (k === 'l') {
        e.preventDefault();
        toggleClear('logo');
        return;
      }
      if (k === 'arrowright' || k === 'arrowdown' || e.key === ' ') {
        e.preventDefault();
        const r = go(1);
        if (r === 'end') void endPresent();
        return;
      }
      if (k === 'arrowleft' || k === 'arrowup') {
        e.preventDefault();
        go(-1);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [live, go, toggleClear]);

  async function startPresent() {
    setError('');
    setPreparing(true);
    setPrepProgress(0);
    try {
      const { slides } = usePresenterStore.getState();
      const start = Date.now();
      await preloadAllMedia(slides, mediaURLs, (done, total) => {
        setPrepProgress(total ? Math.round((done / total) * 100) : 100);
      });
      const elapsed = Date.now() - start;
      const minMs = 1500;
      if (elapsed < minMs) await new Promise((r) => setTimeout(r, minMs - elapsed));
      setPrepProgress(100);
      setPreparing(false);

      setLive(true);
      setPresenting(true);
      setClear('');
      try {
        await openAudienceOnDisplay(selectedDisplay);
        syncAudience();
      } catch (err) {
        console.error(err);
        setError(err instanceof Error ? err.message : 'Could not open audience display');
        // Keep presenter UI even if audience fails
      }
    } catch (err) {
      setPreparing(false);
      setError(err instanceof Error ? err.message : 'Failed to prepare presentation');
    }
  }

  async function endPresent() {
    setLive(false);
    setPresenting(false);
    setClear('');
    pushAudienceClear('');
    try {
      localStorage.setItem('audience-stage-html', '');
    } catch {
      /* ignore */
    }
    await closeAudienceWindow();
  }

  const title = formatServiceTitle(services[activeService]);

  if (presenting && live) {
    return <PresenterView onEnd={() => void endPresent()} />;
  }

  return (
    <div className={homeMode ? 'home-mode app-root' : 'app-root'}>
      {preparing && (
        <div className="prep-overlay">
          <div className="prep-card">
            <div className="prep-title">Preparing presentation…</div>
            <div className="prep-bar">
              <div className="prep-fill" style={{ width: `${prepProgress}%` }} />
            </div>
            <div className="prep-pct">{prepProgress}% loaded</div>
          </div>
        </div>
      )}
      {!homeMode && (
        <TopBar
          serviceTitle={title}
          live={live}
          onServices={() => setHomeMode(true)}
          onPresent={() => void startPresent()}
          onSettings={() => setError('Settings panel — port in Phase 2')}
        />
      )}
      {homeMode ? <HomeServices /> : <EditorLayout />}
      {error && (
        <div className="toast on" role="status">
          {error}
        </div>
      )}
    </div>
  );
}
