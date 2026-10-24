# Church Presenter

Church Presenter is a desktop app for preparing church services and presenting slides to an audience display on Windows and macOS. It is built with Tauri 2, Rust, and a Vite frontend. The current app entry point is `index.html` with `src/presenter.js`; React components in `src/` are part of the ongoing frontend migration.

## Features

- Create and organize services, songs, slides, and reusable slide templates.
- Edit slide text, backgrounds, media, and presentation layouts.
- Import PowerPoint and PDF slides into a service, and export or import `.cpres` service backups.
- Present to a separate audience display, with presenter controls and keyboard shortcuts.
- Connect Google Drive for service and template sync.

## Requirements

- Windows 10 or later, or a supported macOS version
- Node.js and npm
- Rust toolchain
- Windows: MSVC build tools and Microsoft Edge WebView2 Runtime
- macOS: Xcode Command Line Tools

## Development

Install frontend dependencies:

```powershell
npm install
```

Run the desktop app in development mode:

```powershell
npm run tauri dev
```

Run the frontend in a browser for UI development:

```powershell
npm run dev
```

Build the frontend:

```powershell
npm run build
```

Build the Windows desktop installer:

```powershell
npm run tauri build
```

The NSIS installer is configured for a per-user installation. Tauri development and packaging require the Rust and Windows desktop build prerequisites above.

Build the macOS app and DMG on a Mac:

```bash
npm install
npm run tauri -- build --bundles app,dmg --config src-tauri/tauri.macos.conf.json
```

The macOS build uses the system Keychain for Google Drive tokens and AppleScript to control Microsoft PowerPoint. Builds for distribution outside the Mac App Store need Apple Developer ID signing and notarization to avoid Gatekeeper warnings.

## Data and configuration

The app stores its SQLite database and imported background media in its per-user application data directory. Google Drive sign-in uses the Desktop OAuth client ID bundled with the app; refresh tokens are protected with Windows DPAPI or the macOS Keychain. A local `src-tauri/gdrive-client-secret.txt` can optionally provide a client secret at build time and is excluded from Git.

## Project layout

- `index.html`, `audience.html` — presenter and audience display pages
- `src/presenter.js` — current presenter interface and behavior
- `src-tauri/` — Tauri desktop shell, Rust commands, SQLite persistence, and Windows packaging configuration
- `src/components/`, `src/pages/`, and `src/data/` — React UI and supporting modules used in the frontend migration

