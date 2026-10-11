import { useEffect, useRef, useState } from 'react';
import { AUDIENCE_BC } from '../data/present';
import { AUDIENCE_SLIDE_CSS } from '../data/slideRender';
import './AudiencePage.css';

/**
 * Full-screen audience output. Listens to BroadcastChannel + localStorage
 * for stage HTML and clear mode. Forwards keyboard to the main window.
 */
export function AudiencePage() {
  const [html, setHtml] = useState('');
  const [clear, setClear] = useState('');
  const [logo, setLogo] = useState('');
  const lastHtml = useRef('');
  const lastClr = useRef('');

  useEffect(() => {
    try {
      setLogo(localStorage.getItem('church-presenter-logo') || '');
    } catch {
      /* ignore */
    }

    const applyHtml = (h: string) => {
      if (!h || h === lastHtml.current) return;
      lastHtml.current = h;
      setHtml(h);
    };
    const applyClear = (mode: string) => {
      const m = mode || '';
      if (m === lastClr.current) return;
      lastClr.current = m;
      setClear(m);
    };

    try {
      applyHtml(localStorage.getItem('audience-stage-html') || '');
      applyClear(localStorage.getItem('audience-clr') || '');
    } catch {
      /* ignore */
    }

    let bc: BroadcastChannel | null = null;
    try {
      bc = new BroadcastChannel(AUDIENCE_BC);
      bc.onmessage = (e) => {
        const d = e.data;
        if (!d) return;
        if (d.type === 'stage') applyHtml(d.html || '');
        if (d.type === 'clear') applyClear(d.mode || '');
      };
      bc.postMessage({ type: 'ready' });
    } catch {
      /* ignore */
    }

    const onStorage = (e: StorageEvent) => {
      if (e.key === 'audience-stage-html' && e.newValue) applyHtml(e.newValue);
      if (e.key === 'audience-clr') applyClear(e.newValue || '');
    };
    window.addEventListener('storage', onStorage);

    const poll = window.setInterval(() => {
      try {
        applyHtml(localStorage.getItem('audience-stage-html') || '');
        applyClear(localStorage.getItem('audience-clr') || '');
      } catch {
        /* ignore */
      }
    }, 300);

    const sendControl = (action: string, key = '') => {
      const payload = { type: 'control', action, key, t: Date.now() };
      try {
        if (bc) bc.postMessage(payload);
      } catch {
        /* ignore */
      }
      try {
        localStorage.setItem('audience-control', JSON.stringify(payload));
      } catch {
        /* ignore */
      }
    };

    const onKey = (e: KeyboardEvent) => {
      const k = e.key;
      const lower = k.length === 1 ? k.toLowerCase() : k;
      if (k === 'Escape') {
        e.preventDefault();
        sendControl('end');
        return;
      }
      if (lower === 'b') {
        e.preventDefault();
        sendControl('toggle-clear', 'black');
        return;
      }
      if (lower === 'w') {
        e.preventDefault();
        sendControl('toggle-clear', 'white');
        return;
      }
      if (lower === 'l') {
        e.preventDefault();
        sendControl('toggle-clear', 'logo');
        return;
      }
      if (k === 'ArrowRight' || k === 'ArrowDown' || k === ' ' || k === 'PageDown') {
        e.preventDefault();
        sendControl('next');
        return;
      }
      if (k === 'ArrowLeft' || k === 'ArrowUp' || k === 'PageUp') {
        e.preventDefault();
        sendControl('prev');
        return;
      }
    };
    document.addEventListener('keydown', onKey, true);
    window.addEventListener('keydown', onKey, true);

    const onClick = (e: MouseEvent) => {
      if (e.button !== 0) return;
      sendControl('next');
    };
    document.addEventListener('click', onClick);

    try {
      document.body.focus();
    } catch {
      /* ignore */
    }

    return () => {
      window.removeEventListener('storage', onStorage);
      clearInterval(poll);
      document.removeEventListener('keydown', onKey, true);
      window.removeEventListener('keydown', onKey, true);
      document.removeEventListener('click', onClick);
      try {
        bc?.close();
      } catch {
        /* ignore */
      }
    };
  }, []);

  return (
    <div className="audience-root">
      <style>{AUDIENCE_SLIDE_CSS}</style>
      <div
        id="stage"
        className="audience-stage"
        dangerouslySetInnerHTML={{
          __html: html || '<div class="wait">Waiting for slides…</div>',
        }}
      />
      <div id="clear" className={clear ? `on ${clear}` : ''} aria-hidden={!clear}>
        {clear === 'logo' && logo ? <img src={logo} alt="Logo" /> : null}
      </div>
    </div>
  );
}
