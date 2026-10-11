/**
 * Temporary host: loads the full legacy app so nothing is missing
 * while real React panels are ported.
 *
 * Setup:
 * 1. Copy index.legacy.html → public/index.legacy.html
 * 2. Copy audience.legacy.html → public/audience.legacy.html
 * 3. In main.tsx, render <LegacyBridge /> instead of <App /> until parity.
 *
 * Present in the legacy app should open audience.legacy.html
 * (or ?audience=1 on the legacy file) as before.
 */
export function LegacyBridge() {
  return (
    <iframe
      title="Church Presenter"
      src="/index.legacy.html"
      style={{
        position: 'fixed',
        inset: 0,
        width: '100%',
        height: '100%',
        border: 0,
        margin: 0,
        padding: 0,
      }}
      allow="fullscreen; display-capture; clipboard-read; clipboard-write"
    />
  );
}
