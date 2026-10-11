import React from 'react';
import ReactDOM from 'react-dom/client';
import { LegacyBridge } from './LegacyBridge';
import { AudiencePage } from './pages/AudiencePage';

const isAudience =
  typeof window !== 'undefined' &&
  (new URLSearchParams(window.location.search).get('audience') === '1' ||
    window.location.pathname.includes('audience'));

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    {isAudience ? <AudiencePage /> : <LegacyBridge />}
  </React.StrictMode>
);