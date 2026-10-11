interface Props {
  serviceTitle: string;
  live: boolean;
  onServices: () => void;
  onPresent: () => void;
  onSettings: () => void;
  onLibrary?: () => void;
  onMedia?: () => void;
  onImport?: () => void;
}

export function TopBar({
  serviceTitle,
  live,
  onServices,
  onPresent,
  onSettings,
  onLibrary,
  onMedia,
  onImport,
}: Props) {
  return (
    <header className="top">
      <b className="brand">Service Slides</b>
      <button type="button" onClick={onServices}>
        Services
      </button>
      <button type="button" className="service-title" title="Service">
        <span className="svc">{serviceTitle}</span>
      </button>
      <span className="sp" />
      {onLibrary && (
        <button type="button" onClick={onLibrary}>
          Song Library
        </button>
      )}
      {onMedia && (
        <button type="button" onClick={onMedia}>
          Media Library
        </button>
      )}
      {onImport && (
        <button type="button" onClick={onImport}>
          Import slides
        </button>
      )}
      <button
        type="button"
        className="icon-btn"
        title="Settings"
        aria-label="Settings"
        onClick={onSettings}
      >
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <circle cx="12" cy="12" r="3" />
          <path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42" />
        </svg>
      </button>
      <button type="button" className="pri" onClick={onPresent}>
        {live ? 'End presentation' : 'Present'}
      </button>
    </header>
  );
}
