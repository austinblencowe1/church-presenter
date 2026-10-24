mod database;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use tauri::{Emitter, Manager};

// The OAuth client ID is public configuration. Embed the checked-in value so
// packaged installs can seed the per-user config on first launch.
const BUNDLED_GDRIVE_CLIENT_ID: &str = include_str!("../../gdrive-client-id.txt");
const BUNDLED_GDRIVE_CLIENT_SECRET: Option<&str> =
    option_env!("CHURCH_PRESENTER_GDRIVE_CLIENT_SECRET");

struct PendingCpresFiles(Mutex<Vec<String>>);

#[tauri::command]
fn take_pending_cpres_files(state: tauri::State<'_, PendingCpresFiles>) -> Vec<String> {
    state.0.lock().map(|mut paths| std::mem::take(&mut *paths)).unwrap_or_default()
}

#[tauri::command]
fn read_cpres_file(path: String) -> Result<String, String> {
    let bytes = std::fs::read(&path).map_err(|error| error.to_string())?;
    Ok(base64::Engine::encode(&base64::engine::general_purpose::STANDARD, bytes))
}

#[tauri::command]
fn load_tavily_key(app: tauri::AppHandle) -> Result<String, String> {
    let path = app
        .path()
        .app_data_dir()
        .map_err(|error| error.to_string())?
        .join("tavily.key");
    match std::fs::read_to_string(path) {
        Ok(value) => Ok(value.trim().to_string()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(String::new()),
        Err(error) => Err(error.to_string()),
    }
}

#[tauri::command]
fn save_tavily_key(app: tauri::AppHandle, api_key: String) -> Result<(), String> {
    let path = app
        .path()
        .app_data_dir()
        .map_err(|error| error.to_string())?;
    std::fs::create_dir_all(&path).map_err(|error| error.to_string())?;
    std::fs::write(path.join("tavily.key"), api_key.trim()).map_err(|error| error.to_string())
}

#[tauri::command(rename_all = "camelCase")]
fn store_background_media(
    app: tauri::AppHandle,
    database: tauri::State<'_, database::Database>,
    id: String,
    name: String,
    kind: String,
    extension: String,
    data_base64: String,
) -> Result<database::MediaAsset, String> {
    use base64::Engine;
    let bytes = base64::engine::general_purpose::STANDARD
        .decode(data_base64)
        .map_err(|_| "Could not read this media file.".to_string())?;
    if bytes.len() > 150 * 1024 * 1024 {
        return Err("Background media must be 150 MB or smaller.".into());
    }
    let extension = extension.trim_start_matches('.').to_ascii_lowercase();
    let valid_extension = match kind.as_str() {
        "image" => matches!(extension.as_str(), "png" | "jpg" | "jpeg" | "webp" | "gif"),
        "video" => matches!(extension.as_str(), "mp4" | "webm" | "mov"),
        _ => false,
    };
    if !valid_extension
        || id.is_empty()
        || !id
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '-')
    {
        return Err("That background media file type is not supported.".into());
    }
    let directory = app
        .path()
        .app_data_dir()
        .map_err(|error| error.to_string())?
        .join("background-media");
    std::fs::create_dir_all(&directory).map_err(|error| error.to_string())?;
    let path = directory.join(format!("{id}.{extension}"));
    std::fs::write(&path, &bytes).map_err(|error| error.to_string())?;
    let asset = database::MediaAsset {
        id,
        name,
        kind,
        source: path.to_string_lossy().into_owned(),
        size: bytes.len() as u64,
        added_at: chrono_like_timestamp(),
    };
    if let Err(error) = database::save_media(asset.clone(), database) {
        let _ = std::fs::remove_file(path);
        return Err(error);
    }
    Ok(asset)
}

fn chrono_like_timestamp() -> String {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|duration| duration.as_secs().to_string())
        .unwrap_or_default()
}

#[tauri::command(rename_all = "camelCase")]
fn delete_background_media(
    app: tauri::AppHandle,
    database: tauri::State<'_, database::Database>,
    media_id: String,
) -> Result<(), String> {
    let asset = database::list_media(database.clone())?
        .into_iter()
        .find(|asset| asset.id == media_id);
    database::delete_media(media_id, database)?;
    if let Some(asset) = asset {
        let directory = app
            .path()
            .app_data_dir()
            .map_err(|error| error.to_string())?
            .join("background-media");
        let path = std::path::PathBuf::from(asset.source);
        if path.parent() == Some(directory.as_path()) {
            let _ = std::fs::remove_file(path);
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// PowerPoint hand-off (Church Presenter → Microsoft PowerPoint slideshow)
// ---------------------------------------------------------------------------

#[tauri::command(rename_all = "camelCase")]
fn store_presentation_file(
    app: tauri::AppHandle,
    id: String,
    file_name: String,
    extension: String,
    data_base64: String,
) -> Result<String, String> {
    use base64::Engine;
    let bytes = base64::engine::general_purpose::STANDARD
        .decode(data_base64)
        .map_err(|_| "Could not read the presentation file.".to_string())?;
    if bytes.len() > 200 * 1024 * 1024 {
        return Err("Presentation file must be 200 MB or smaller.".into());
    }
    let id_ok = !id.is_empty()
        && id
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '-' || c == '_');
    if !id_ok {
        return Err("Invalid presentation id.".into());
    }
    let extension = extension.trim_start_matches('.').to_ascii_lowercase();
    if !matches!(extension.as_str(), "pptx" | "ppt" | "pdf") {
        return Err("Only .pptx, .ppt, and .pdf are supported for hand-off.".into());
    }
    let directory = app
        .path()
        .app_data_dir()
        .map_err(|e| e.to_string())?
        .join("imported-presentations");
    std::fs::create_dir_all(&directory).map_err(|e| e.to_string())?;
    let stem = file_name
        .rsplit(['/', '\\'])
        .next()
        .unwrap_or("presentation")
        .chars()
        .map(|c| {
            if c.is_ascii_alphanumeric() || matches!(c, '-' | '_' | '.') {
                c
            } else {
                '_'
            }
        })
        .collect::<String>();
    let file = if stem.to_ascii_lowercase().ends_with(&format!(".{extension}")) {
        stem
    } else {
        format!("{stem}.{extension}")
    };
    let path = directory.join(format!("{id}_{file}"));
    std::fs::write(&path, &bytes).map_err(|e| e.to_string())?;
    Ok(path.to_string_lossy().into_owned())
}

#[cfg(windows)]
fn powerpoint_exe_candidates() -> &'static [&'static str] {
    &[
        r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\root\Office15\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office15\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\root\Office14\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office14\POWERPNT.EXE",
        "POWERPNT.EXE",
        "powerpnt.exe",
    ]
}

/// Foolproof PowerPoint slideshow launch (Windows).
#[tauri::command(rename_all = "camelCase")]
fn launch_ppt_slideshow(path: String) -> Result<String, String> {
    let path = path.trim().to_string();
    if path.is_empty() {
        return Err("No file path provided.".into());
    }
    let path_buf = std::path::PathBuf::from(&path);
    if !path_buf.is_file() {
        return Err(format!("File not found: {path}"));
    }

    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        use std::process::Command;
        const CREATE_NO_WINDOW: u32 = 0x08000000;

        for exe in powerpoint_exe_candidates() {
            let mut cmd = Command::new(exe);
            cmd.arg("/S").arg(&path);
            cmd.creation_flags(CREATE_NO_WINDOW);
            if cmd.spawn().is_ok() {
                return Ok(format!("started {exe} /S"));
            }
        }

        let mut cmd = Command::new("cmd.exe");
        cmd.args(["/C", "start", "", "POWERPNT.EXE", "/S", &path]);
        cmd.creation_flags(CREATE_NO_WINDOW);
        if cmd.status().map(|s| s.success()).unwrap_or(false) {
            return Ok("started via cmd POWERPNT.EXE /S".into());
        }

        open::that(&path).map_err(|e| format!("open failed: {e}"))?;
        Ok("opened with default association".into())
    }

    #[cfg(target_os = "macos")]
    {
        open_path_with_powerpoint(&path, true)?;
        Ok("started Microsoft PowerPoint slideshow".into())
    }

    #[cfg(not(any(windows, target_os = "macos")))]
    {
        open::that(&path).map_err(|e| e.to_string())?;
        Ok("opened".into())
    }
}

#[tauri::command(rename_all = "camelCase")]
fn prepare_presentation(path: String) -> Result<(), String> {
    open_path_with_powerpoint(path.trim(), false)
}

#[tauri::command(rename_all = "camelCase")]
fn open_presentation_slideshow(path: String) -> Result<(), String> {
    open_path_with_powerpoint(path.trim(), true)
}

#[tauri::command(rename_all = "camelCase")]
fn open_presentation_file(path: String) -> Result<(), String> {
    open_path_with_powerpoint(path.trim(), false)
}

fn open_path_with_powerpoint(path: &str, slideshow: bool) -> Result<(), String> {
    if path.is_empty() {
        return Err("No file path provided.".into());
    }
    let path_buf = std::path::PathBuf::from(path);
    if !path_buf.exists() {
        return Err(format!("File not found: {path}"));
    }

    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        use std::process::Command;
        const CREATE_NO_WINDOW: u32 = 0x08000000;

        for exe in powerpoint_exe_candidates() {
            let mut cmd = Command::new(exe);
            if slideshow {
                cmd.arg("/S");
            }
            cmd.arg(path);
            cmd.creation_flags(CREATE_NO_WINDOW);
            if cmd.spawn().is_ok() {
                return Ok(());
            }
        }

        let mut cmd = Command::new("cmd");
        cmd.arg("/C").arg("start").arg("");
        if slideshow {
            cmd.arg("POWERPNT.EXE").arg("/S").arg(path);
        } else {
            cmd.arg(path);
        }
        cmd.creation_flags(CREATE_NO_WINDOW);
        if cmd.spawn().is_ok() {
            return Ok(());
        }

        open::that(path).map_err(|e| format!("Could not open presentation: {e}"))?;
        Ok(())
    }

    #[cfg(target_os = "macos")]
    {
        let script_path = path.replace('\\', "\\\\").replace('"', "\\\"");
        let script = if slideshow {
            format!(
                "tell application \"Microsoft PowerPoint\"\nactivate\nopen (POSIX file \"{script_path}\")\ndelay 1\nrun slide show slide show settings of active presentation\nend tell"
            )
        } else {
            format!(
                "tell application \"Microsoft PowerPoint\"\nactivate\nopen (POSIX file \"{script_path}\")\nend tell"
            )
        };
        match run_osascript(&script) {
            Ok(_) => Ok(()),
            Err(error) if !slideshow => {
                open::that(path).map_err(|open_error| format!("{error}; could not open presentation: {open_error}"))
            }
            Err(error) => Err(format!("Could not start the PowerPoint slideshow: {error}")),
        }
    }

    #[cfg(not(any(windows, target_os = "macos")))]
    {
        open::that(path).map_err(|e| e.to_string())?;
        Ok(())
    }
}

#[cfg(target_os = "macos")]
fn run_osascript(script: &str) -> Result<String, String> {
    let output = std::process::Command::new("osascript")
        .arg("-e")
        .arg(script)
        .output()
        .map_err(|e| format!("Could not run AppleScript: {e}"))?;
    if output.status.success() {
        Ok(String::from_utf8_lossy(&output.stdout).trim().to_string())
    } else {
        Err(String::from_utf8_lossy(&output.stderr).trim().to_string())
    }
}

struct CloseGuard(AtomicBool);

#[tauri::command]
fn close_guard_ready(state: tauri::State<'_, CloseGuard>) {
    state.0.store(true, Ordering::SeqCst);
}

#[tauri::command]
fn quit_app(app: tauri::AppHandle) {
    app.exit(0);
}

#[cfg(windows)]
fn run_powershell(script: &str) -> Result<String, String> {
    use std::os::windows::process::CommandExt;
    let output = std::process::Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ])
        .creation_flags(0x08000000)
        .output()
        .map_err(|e| e.to_string())?;
    Ok(String::from_utf8_lossy(&output.stdout).trim().to_string())
}

#[cfg(windows)]
const PPT_WAIT_SCRIPT: &str = r#"
$ErrorActionPreference = 'SilentlyContinue'
function Get-Show {
  try {
    $a = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    if ($a -and $a.SlideShowWindows.Count -gt 0) { return $a.SlideShowWindows.Item(1) }
  } catch {}
  return $null
}
$deadline = (Get-Date).AddSeconds(__SECS__)
$w = $null
while ((Get-Date) -lt $deadline) {
  $w = Get-Show
  if ($w) { break }
  Start-Sleep -Milliseconds 400
}
if (-not $w) { Write-Output 'timeout'; exit 0 }
$miss = 0
while ($true) {
  $w = Get-Show
  if (-not $w) {
    $miss++
    if ($miss -ge 3) { break }
  } else {
    $miss = 0
    try {
      if ($w.View.State -eq 5) { try { $w.View.Exit() } catch {}; break }
    } catch {}
  }
  Start-Sleep -Milliseconds 300
}
Write-Output 'ended'
"#;

#[cfg(windows)]
const PPT_STOP_SCRIPT: &str = r#"
$ErrorActionPreference = 'SilentlyContinue'
try {
  $a = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
  for ($i = 0; $i -lt 5 -and $a.SlideShowWindows.Count -gt 0; $i++) {
    $a.SlideShowWindows.Item(1).View.Exit()
    Start-Sleep -Milliseconds 150
  }
} catch {}
Write-Output 'stopped'
"#;

#[tauri::command(rename_all = "camelCase")]
async fn wait_ppt_slideshow_end(start_timeout_secs: Option<u64>) -> Result<String, String> {
    #[cfg(windows)]
    {
        let secs = start_timeout_secs.unwrap_or(40);
        let script = PPT_WAIT_SCRIPT.replace("__SECS__", &secs.to_string());
        return tauri::async_runtime::spawn_blocking(move || run_powershell(&script))
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(target_os = "macos")]
    {
        let checks = start_timeout_secs.unwrap_or(40).saturating_mul(5);
        let script = format!(r#"
tell application "Microsoft PowerPoint"
  set showStarted to false
  repeat {checks} times
    if (count of slide show windows) > 0 then
      set showStarted to true
      exit repeat
    end if
    delay 0.2
  end repeat
  if showStarted is false then return "timeout"
  repeat while (count of slide show windows) > 0
    delay 0.3
  end repeat
  return "ended"
end tell
"#);
        return tauri::async_runtime::spawn_blocking(move || run_osascript(&script))
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(not(any(windows, target_os = "macos")))]
    {
        let _ = start_timeout_secs;
        Err("PowerPoint slideshow tracking is not implemented for this platform.".into())
    }
}

#[tauri::command]
async fn stop_ppt_slideshow() -> Result<String, String> {
    #[cfg(windows)]
    {
        return tauri::async_runtime::spawn_blocking(|| run_powershell(PPT_STOP_SCRIPT))
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(target_os = "macos")]
    {
        let script = r#"
tell application "Microsoft PowerPoint"
  if (count of slide show windows) > 0 then
    exit slide show slide show view of slide show window 1 of active presentation
  end if
  return "stopped"
end tell
"#;
        return tauri::async_runtime::spawn_blocking(move || run_osascript(script))
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(not(any(windows, target_os = "macos")))]
    {
        Ok("noop".into())
    }
}

/// Bring the running PowerPoint slideshow window to the foreground (Windows).
#[tauri::command]
async fn focus_ppt_slideshow() -> Result<String, String> {
    #[cfg(windows)]
    {
        const SCRIPT: &str = r#"
$ErrorActionPreference = 'SilentlyContinue'
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class PptFocus {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
}
"@
function Get-Show {
  try {
    $a = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    if ($a -and $a.SlideShowWindows.Count -gt 0) { return $a.SlideShowWindows.Item(1) }
  } catch {}
  return $null
}
$deadline = (Get-Date).AddSeconds(12)
$w = $null
while ((Get-Date) -lt $deadline) {
  $w = Get-Show
  if ($w) { break }
  Start-Sleep -Milliseconds 250
}
if (-not $w) { Write-Output 'no-show'; exit 0 }
try {
  $hwnd = [IntPtr]$w.HWND
  if ([PptFocus]::IsIconic($hwnd)) { [void][PptFocus]::ShowWindow($hwnd, 9) }
  [void][PptFocus]::SetForegroundWindow($hwnd)
  Write-Output 'focused'
} catch {
  Write-Output 'focus-failed'
}
"#;
        return tauri::async_runtime::spawn_blocking(move || run_powershell(SCRIPT))
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(target_os = "macos")]
    {
        let script = r#"
tell application "Microsoft PowerPoint"
  activate
  set frontmost to true
  return "focused"
end tell
"#;
        return tauri::async_runtime::spawn_blocking(move || run_osascript(script))
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(not(any(windows, target_os = "macos")))]
    {
        Ok("noop".into())
    }
}

// ─── Google Drive (OAuth PKCE + native secure token storage + sync) ─────────
//
// Put these in the app data folder (app.path().app_data_dir(), e.g.
// %APPDATA%\<identifier>\):
//   gdrive-client-id.txt      one line: your Desktop OAuth Client ID
//   gdrive-client-secret.txt  one line: your Desktop OAuth Client secret
// (or env vars CHURCH_PRESENTER_GDRIVE_CLIENT_ID / _CLIENT_SECRET)
//
// Scope: drive (read + write) so the app can upload/sync into a shared folder.
// Tokens: Windows DPAPI on Windows and the macOS Keychain on macOS.

const GDRIVE_SCOPE: &str = "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/userinfo.email openid";
const GDRIVE_AUTH_URL: &str = "https://accounts.google.com/o/oauth2/v2/auth";
const GDRIVE_TOKEN_URL: &str = "https://oauth2.googleapis.com/token";
const GDRIVE_REVOKE_URL: &str = "https://oauth2.googleapis.com/revoke";

#[derive(serde::Serialize, serde::Deserialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
struct GdriveStatus {
    ok: bool,
    connected: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    email: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    folder_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    folder_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    scope_note: Option<String>,
}

fn gdrive_data_dir(app: &tauri::AppHandle) -> Result<std::path::PathBuf, String> {
    let dir = app.path().app_data_dir().map_err(|e| e.to_string())?;
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir)
}

fn read_client_id(app: &tauri::AppHandle) -> Result<String, String> {
    if let Ok(v) = std::env::var("CHURCH_PRESENTER_GDRIVE_CLIENT_ID") {
        let t = v.trim().to_string();
        if !t.is_empty() {
            return Ok(t);
        }
    }
    let path = gdrive_data_dir(app)?.join("gdrive-client-id.txt");
    let raw = std::fs::read_to_string(&path).map_err(|_| {
        format!(
            "Missing Google Client ID. Create a Desktop OAuth client in Google Cloud, then save the Client ID as a single line in:\n{}",
            path.display()
        )
    })?;
    let id = raw
        .lines()
        .map(str::trim)
        .find(|l| !l.is_empty() && !l.starts_with('#'))
        .unwrap_or("")
        .to_string();
    if id.is_empty() {
        return Err(format!("gdrive-client-id.txt is empty ({})", path.display()));
    }
    Ok(id)
}

fn read_client_secret(app: &tauri::AppHandle) -> Option<String> {
    if let Ok(v) = std::env::var("CHURCH_PRESENTER_GDRIVE_CLIENT_SECRET") {
        let t = v.trim().to_string();
        if !t.is_empty() {
            return Some(t);
        }
    }
    let path = gdrive_data_dir(app).ok()?.join("gdrive-client-secret.txt");
    let raw = std::fs::read_to_string(path).ok()?;
    raw.lines()
        .map(str::trim)
        .find(|l| !l.is_empty() && !l.starts_with('#'))
        .map(|s| s.to_string())
}

fn folder_config_path(app: &tauri::AppHandle) -> Result<std::path::PathBuf, String> {
    Ok(gdrive_data_dir(app)?.join("gdrive-folder.json"))
}

fn load_folder_config(app: &tauri::AppHandle) -> (Option<String>, Option<String>) {
    let Ok(path) = folder_config_path(app) else {
        return (None, None);
    };
    let Ok(raw) = std::fs::read_to_string(path) else {
        return (None, None);
    };
    let v: serde_json::Value = serde_json::from_str(&raw).unwrap_or(serde_json::json!({}));
    (
        v.get("folderId").and_then(|x| x.as_str()).map(|s| s.to_string()),
        v.get("folderName").and_then(|x| x.as_str()).map(|s| s.to_string()),
    )
}

fn save_folder_config(app: &tauri::AppHandle, folder_id: &str, folder_name: &str) -> Result<(), String> {
    let path = folder_config_path(app)?;
    let body = serde_json::json!({ "folderId": folder_id, "folderName": folder_name });
    std::fs::write(path, body.to_string()).map_err(|e| e.to_string())
}

#[cfg(windows)]
fn dpapi_protect(plain: &str) -> Result<String, String> {
    // Protect UTF-8 bytes with Windows DPAPI (CurrentUser) via PowerShell
    let b64_in = {
        const T: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
        let data = plain.as_bytes();
        let mut out = String::new();
        let mut i = 0;
        while i < data.len() {
            let b0 = data[i] as u32;
            let b1 = if i + 1 < data.len() { data[i + 1] as u32 } else { 0 };
            let b2 = if i + 2 < data.len() { data[i + 2] as u32 } else { 0 };
            let triple = (b0 << 16) | (b1 << 8) | b2;
            out.push(T[((triple >> 18) & 63) as usize] as char);
            out.push(T[((triple >> 12) & 63) as usize] as char);
            if i + 1 < data.len() {
                out.push(T[((triple >> 6) & 63) as usize] as char);
            } else {
                out.push('=');
            }
            if i + 2 < data.len() {
                out.push(T[(triple & 63) as usize] as char);
            } else {
                out.push('=');
            }
            i += 3;
        }
        out
    };
    let script = format!(
        r#"$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Security
$raw = [Convert]::FromBase64String('{b64}')
$prot = [System.Security.Cryptography.ProtectedData]::Protect($raw, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
[Convert]::ToBase64String($prot)
"#,
        b64 = b64_in.replace('\'', "''")
    );
    run_powershell(&script)
}

#[cfg(windows)]
fn dpapi_unprotect(b64_prot: &str) -> Result<String, String> {
    let script = format!(
        r#"$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Security
$prot = [Convert]::FromBase64String('{b64}')
$raw = [System.Security.Cryptography.ProtectedData]::Unprotect($prot, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
[System.Text.Encoding]::UTF8.GetString($raw)
"#,
        b64 = b64_prot.trim().replace('\'', "''")
    );
    run_powershell(&script)
}

#[cfg(not(any(windows, target_os = "macos")))]
fn dpapi_protect(_plain: &str) -> Result<String, String> {
    Err("Google Drive token storage is not implemented for this platform.".into())
}
#[cfg(not(any(windows, target_os = "macos")))]
fn dpapi_unprotect(_b64: &str) -> Result<String, String> {
    Err("Google Drive token storage is not implemented for this platform.".into())
}

#[cfg(windows)]
fn token_store_path(app: &tauri::AppHandle) -> Result<std::path::PathBuf, String> {
    Ok(gdrive_data_dir(app)?.join("gdrive-refresh.dpapi"))
}

#[cfg(target_os = "macos")]
fn token_keyring_entry() -> Result<keyring::Entry, String> {
    keyring::Entry::new("com.mbc.churchpresenter", "google-refresh-token")
        .map_err(|e| format!("Could not access macOS Keychain: {e}"))
}

fn save_refresh_token(app: &tauri::AppHandle, refresh: &str) -> Result<(), String> {
    #[cfg(windows)]
    {
    let protected = dpapi_protect(refresh)?;
    std::fs::write(token_store_path(app)?, protected.trim()).map_err(|e| e.to_string())
    }
    #[cfg(target_os = "macos")]
    {
        let _ = app;
        token_keyring_entry()?.set_password(refresh)
            .map_err(|e| format!("Could not save Google sign-in to macOS Keychain: {e}"))
    }
    #[cfg(not(any(windows, target_os = "macos")))]
    {
        let _ = (app, refresh);
        Err("Google Drive token storage is not implemented for this platform.".into())
    }
}

fn load_refresh_token(app: &tauri::AppHandle) -> Option<String> {
    #[cfg(windows)]
    {
    let path = token_store_path(app).ok()?;
    let raw = std::fs::read_to_string(path).ok()?;
    let t = raw.trim();
    if t.is_empty() {
        return None;
    }
    dpapi_unprotect(t).ok().map(|s| s.trim().to_string())
    }
    #[cfg(target_os = "macos")]
    {
        let _ = app;
        token_keyring_entry().ok()?.get_password().ok()
    }
    #[cfg(not(any(windows, target_os = "macos")))]
    {
        let _ = app;
        None
    }
}

fn clear_refresh_token(app: &tauri::AppHandle) {
    #[cfg(windows)]
    {
    if let Ok(path) = token_store_path(app) {
        let _ = std::fs::remove_file(path);
    }
    }
    #[cfg(target_os = "macos")]
    {
        let _ = app;
        if let Ok(entry) = token_keyring_entry() {
            let _ = entry.delete_credential();
        }
    }
}

fn random_url_safe(len: usize) -> String {
    use rand::distributions::Alphanumeric;
    use rand::Rng;
    rand::thread_rng()
        .sample_iter(&Alphanumeric)
        .take(len)
        .map(char::from)
        .collect()
}

fn sha256_base64url(input: &str) -> String {
    use base64::engine::general_purpose::URL_SAFE_NO_PAD;
    use base64::Engine;
    use sha2::{Digest, Sha256};
    URL_SAFE_NO_PAD.encode(Sha256::digest(input.as_bytes()))
}

fn http_json(method: &str, url: &str, body: Option<&str>, auth_bearer: Option<&str>) -> Result<String, String> {
    let client = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(60))
        .build()
        .map_err(|e| e.to_string())?;
    let method = reqwest::Method::from_bytes(method.as_bytes())
        .map_err(|e| format!("Invalid HTTP method: {e}"))?;
    let mut request = client.request(method, url);
    if let Some(token) = auth_bearer {
        request = request.bearer_auth(token);
    }
    if let Some(body) = body {
        request = request
            .header(reqwest::header::CONTENT_TYPE, "application/x-www-form-urlencoded")
            .body(body.to_string());
    }
    let response = request.send().map_err(|e| format!("HTTP request failed: {e}"))?;
    let status = response.status();
    let content = response.text().map_err(|e| format!("Could not read HTTP response: {e}"))?;
    if !status.is_success() {
        return Err(format!("HTTP_ERROR:{} | {}", status, content));
    }
    Ok(content)
}

fn urlencoding_encode(s: &str) -> String {
    let mut out = String::new();
    for b in s.as_bytes() {
        match *b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' => out.push(*b as char),
            _ => out.push_str(&format!("%{b:02X}")),
        }
    }
    out
}

fn exchange_code_for_tokens(
    client_id: &str,
    client_secret: Option<&str>,
    code: &str,
    redirect_uri: &str,
    code_verifier: &str,
) -> Result<serde_json::Value, String> {
    let mut body = format!(
        "code={}&client_id={}&redirect_uri={}&grant_type=authorization_code&code_verifier={}",
        urlencoding_encode(code),
        urlencoding_encode(client_id),
        urlencoding_encode(redirect_uri),
        urlencoding_encode(code_verifier)
    );
    if let Some(s) = client_secret {
        body.push_str(&format!("&client_secret={}", urlencoding_encode(s)));
    }
    let raw = http_json("POST", GDRIVE_TOKEN_URL, Some(&body), None)?;
    serde_json::from_str(&raw).map_err(|e| format!("token parse: {e} | {raw}"))
}

fn refresh_access_token(client_id: &str, client_secret: Option<&str>, refresh: &str) -> Result<serde_json::Value, String> {
    let mut body = format!(
        "client_id={}&refresh_token={}&grant_type=refresh_token",
        urlencoding_encode(client_id),
        urlencoding_encode(refresh)
    );
    if let Some(s) = client_secret {
        body.push_str(&format!("&client_secret={}", urlencoding_encode(s)));
    }
    let raw = http_json("POST", GDRIVE_TOKEN_URL, Some(&body), None)?;
    serde_json::from_str(&raw).map_err(|e| format!("refresh parse: {e} | {raw}"))
}

fn get_valid_access_token(app: &tauri::AppHandle) -> Result<String, String> {
    let client_id = read_client_id(app)?;
    let refresh = load_refresh_token(app).ok_or_else(|| "Not connected to Google. Click Connect Google.".to_string())?;
    let secret = read_client_secret(app);
    let tok = refresh_access_token(&client_id, secret.as_deref(), &refresh)?;
    let access = tok
        .get("access_token")
        .and_then(|v| v.as_str())
        .ok_or_else(|| format!("No access_token in refresh response: {tok}"))?
        .to_string();
    if let Some(new_r) = tok.get("refresh_token").and_then(|v| v.as_str()) {
        let _ = save_refresh_token(app, new_r);
    }
    Ok(access)
}

fn fetch_user_email(access: &str) -> Result<String, String> {
    let raw = http_json("GET", "https://www.googleapis.com/oauth2/v2/userinfo", None, Some(access))?;
    let v: serde_json::Value = serde_json::from_str(&raw).map_err(|e| e.to_string())?;
    Ok(v.get("email").and_then(|e| e.as_str()).unwrap_or("").to_string())
}

/// Run local loopback OAuth (PKCE) and persist refresh token via DPAPI.
fn run_pkce_login(app: &tauri::AppHandle) -> Result<GdriveStatus, String> {
    let client_id = read_client_id(app)?;
    let listener = std::net::TcpListener::bind("127.0.0.1:0").map_err(|e| e.to_string())?;
    let port = listener.local_addr().map_err(|e| e.to_string())?.port();
    let redirect_uri = format!("http://127.0.0.1:{port}/oauth2callback");
    let state = random_url_safe(32);
    let verifier = random_url_safe(64);
    let challenge = sha256_base64url(&verifier);
    if challenge.is_empty() {
        return Err("Failed to compute PKCE challenge".into());
    }
    let auth_url = format!(
        "{auth}?client_id={cid}&redirect_uri={redir}&response_type=code&scope={scope}&state={state}&code_challenge={cc}&code_challenge_method=S256&access_type=offline&prompt=consent",
        auth = GDRIVE_AUTH_URL,
        cid = urlencoding_encode(&client_id),
        redir = urlencoding_encode(&redirect_uri),
        scope = urlencoding_encode(GDRIVE_SCOPE),
        state = urlencoding_encode(&state),
        cc = urlencoding_encode(&challenge)
    );

    // Open system browser (the `open` crate handles & in URLs correctly)
    open::that(&auth_url).map_err(|e| format!("Could not open browser: {e}"))?;

    // Wait for the browser callback, polling so we can time out instead of hanging forever
    listener.set_nonblocking(true).map_err(|e| e.to_string())?;
    let started = std::time::Instant::now();
    let (mut stream, _) = loop {
        match listener.accept() {
            Ok(pair) => break pair,
            Err(e) if e.kind() == std::io::ErrorKind::WouldBlock => {
                if started.elapsed() > std::time::Duration::from_secs(180) {
                    return Err("Sign-in timed out. Try Connect Google again.".into());
                }
                std::thread::sleep(std::time::Duration::from_millis(200));
            }
            Err(e) => return Err(format!("OAuth callback wait failed: {e}")),
        }
    };
    let _ = stream.set_nonblocking(false);
    use std::io::{Read, Write};
    let mut buf = [0u8; 8192];
    let n = stream.read(&mut buf).map_err(|e| e.to_string())?;
    let req = String::from_utf8_lossy(&buf[..n]);
    let first_line = req.lines().next().unwrap_or("");
    // GET /oauth2callback?code=...&state=... HTTP/1.1
    let q = first_line.split_whitespace().nth(1).unwrap_or("");
    let qs = q.split('?').nth(1).unwrap_or("");
    let mut code = None;
    let mut got_state = None;
    let mut err = None;
    for pair in qs.split('&') {
        let mut it = pair.splitn(2, '=');
        let k = it.next().unwrap_or("");
        let v = it.next().unwrap_or("");
        let dec = |s: &str| {
            let mut o = String::new();
            let b = s.as_bytes();
            let mut i = 0;
            while i < b.len() {
                if b[i] == b'%' && i + 2 < b.len() {
                    let h = u8::from_str_radix(std::str::from_utf8(&b[i + 1..i + 3]).unwrap_or("00"), 16).unwrap_or(0);
                    o.push(h as char);
                    i += 3;
                } else if b[i] == b'+' {
                    o.push(' ');
                    i += 1;
                } else {
                    o.push(b[i] as char);
                    i += 1;
                }
            }
            o
        };
        match k {
            "code" => code = Some(dec(v)),
            "state" => got_state = Some(dec(v)),
            "error" => err = Some(dec(v)),
            _ => {}
        }
    }
    let html_ok = b"HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nConnection: close\r\n\r\n<html><body style=\"font-family:system-ui;padding:2rem\"><h2>Connected</h2><p>You can close this window and return to Church Presenter.</p></body></html>";
    let html_err = b"HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nConnection: close\r\n\r\n<html><body style=\"font-family:system-ui;padding:2rem\"><h2>Sign-in problem</h2><p>Return to Church Presenter and try again.</p></body></html>";
    if err.is_some() || code.is_none() || got_state.as_deref() != Some(state.as_str()) {
        let _ = stream.write_all(html_err);
        return Err(err.unwrap_or_else(|| "OAuth state mismatch or missing code".into()));
    }
    let _ = stream.write_all(html_ok);
    let secret = read_client_secret(app);
    let tokens = exchange_code_for_tokens(&client_id, secret.as_deref(), code.as_ref().unwrap(), &redirect_uri, &verifier)?;
    let refresh = tokens
        .get("refresh_token")
        .and_then(|v| v.as_str())
        .ok_or_else(|| format!("Google did not return a refresh token. Response: {tokens}"))?;
    save_refresh_token(app, refresh)?;
    let access = tokens.get("access_token").and_then(|v| v.as_str()).unwrap_or("");
    let email = if access.is_empty() {
        String::new()
    } else {
        fetch_user_email(access).unwrap_or_default()
    };
    let (fid, fname) = load_folder_config(app);
    Ok(GdriveStatus {
        ok: true,
        connected: true,
        email: if email.is_empty() { None } else { Some(email) },
        folder_id: fid,
        folder_name: fname,
        error: None,
        scope_note: Some("drive (read + write)".into()),
    })
}

// ── Binary-safe Drive transfer helpers ───────────────────────────────────────
fn http_download(url: &str, bearer: &str, out: &std::path::Path) -> Result<(), String> {
    let response = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(120))
        .build()
        .map_err(|e| e.to_string())?
        .get(url)
        .bearer_auth(bearer)
        .send()
        .map_err(|e| format!("Drive download failed: {e}"))?;
    let status = response.status();
    let bytes = response.bytes().map_err(|e| e.to_string())?;
    if !status.is_success() {
        return Err(format!("HTTP_ERROR:{} | {}", status, String::from_utf8_lossy(&bytes)));
    }
    std::fs::write(out, bytes).map_err(|e| e.to_string())
}

fn http_upload(method: &str, url: &str, bearer: &str, body_file: &std::path::Path, content_type: &str) -> Result<String, String> {
    let method = reqwest::Method::from_bytes(method.as_bytes())
        .map_err(|e| format!("Invalid HTTP method: {e}"))?;
    let bytes = std::fs::read(body_file).map_err(|e| e.to_string())?;
    let response = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(120))
        .build()
        .map_err(|e| e.to_string())?
        .request(method, url)
        .bearer_auth(bearer)
        .header(reqwest::header::CONTENT_TYPE, content_type)
        .body(bytes)
        .send()
        .map_err(|e| format!("Drive upload failed: {e}"))?;
    let status = response.status();
    let content = response.text().map_err(|e| e.to_string())?;
    if !status.is_success() {
        return Err(format!("HTTP_ERROR:{} | {}", status, content));
    }
    Ok(content)
}

fn gdrive_upload_resumable(
    app: &tauri::AppHandle,
    path: &std::path::Path,
    name: &str,
    mime: &str,
    folder_id: Option<&str>,
    existing_file_id: Option<&str>,
    cpres_id: Option<&str>,
) -> Result<serde_json::Value, String> {
    use reqwest::header::{CONTENT_LENGTH, CONTENT_TYPE, LOCATION};
    use std::io::{Read, Seek, SeekFrom};

    let file_size = std::fs::metadata(path).map_err(|e| e.to_string())?.len();
    if file_size > 100 * 1024 * 1024 {
        return Err("Upload must be 100 MB or smaller.".into());
    }
    let access = get_valid_access_token(app)?;
    let existing = existing_file_id.map(str::trim).filter(|s| !s.is_empty());
    let mut metadata = serde_json::json!({"name": name});
    if let Some(id) = cpres_id.map(str::trim).filter(|s| !s.is_empty()) {
        metadata["appProperties"] = serde_json::json!({"churchPresenterServiceId": id});
    }
    if existing.is_none() {
        if let Some(folder) = folder_id.map(str::trim).filter(|s| !s.is_empty()) {
            metadata["parents"] = serde_json::json!([folder]);
        }
    }
    let endpoint = match existing {
        Some(id) => format!("https://www.googleapis.com/upload/drive/v3/files/{}?uploadType=resumable&supportsAllDrives=true&fields=id,name,modifiedTime", urlencoding_encode(id)),
        None => "https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true&fields=id,name,modifiedTime".to_string(),
    };
    let client = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(120))
        .build().map_err(|e| e.to_string())?;
    let method = if existing.is_some() { reqwest::Method::PATCH } else { reqwest::Method::POST };
    let session = client.request(method, endpoint)
        .bearer_auth(&access)
        .header(CONTENT_TYPE, "application/json; charset=UTF-8")
        .header("X-Upload-Content-Type", mime)
        .header("X-Upload-Content-Length", file_size.to_string())
        .json(&metadata)
        .send().map_err(|e| format!("Drive upload session failed: {e}"))?;
    let status = session.status();
    if !status.is_success() {
        return Err(format!("HTTP_ERROR:{} | {}", status, session.text().unwrap_or_default()));
    }
    let session_url = session.headers().get(LOCATION)
        .and_then(|v| v.to_str().ok()).ok_or("Drive did not return a resumable upload URL.")?.to_string();

    const CHUNK_SIZE: usize = 8 * 1024 * 1024; // Drive requires non-final chunks to be multiples of 256 KiB.
    let _ = app.emit("gdrive-upload-progress", serde_json::json!({"bytesSent": 0, "totalBytes": file_size}));
    let mut file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut offset = 0u64;
    while offset < file_size {
        let count = std::cmp::min(CHUNK_SIZE as u64, file_size - offset) as usize;
        file.seek(SeekFrom::Start(offset)).map_err(|e| e.to_string())?;
        let mut chunk = vec![0; count];
        file.read_exact(&mut chunk).map_err(|e| e.to_string())?;
        let end = offset + count as u64 - 1;
        let mut attempt = 0;
        loop {
            let response = client.put(&session_url)
                .header(CONTENT_LENGTH, count.to_string())
                .header("Content-Range", format!("bytes {offset}-{end}/{file_size}"))
                .body(chunk.clone())
                .send();
            match response {
                Ok(response) if response.status().as_u16() == 308 => {
                    let received = response.headers().get("Range").and_then(|v| v.to_str().ok())
                        .and_then(|v| v.rsplit('-').next()).and_then(|v| v.parse::<u64>().ok());
                    let Some(received) = received else {
                        return Err("Drive did not confirm receipt of the upload chunk.".into());
                    };
                    offset = received + 1;
                    let _ = app.emit("gdrive-upload-progress", serde_json::json!({"bytesSent": offset, "totalBytes": file_size}));
                    break;
                }
                Ok(response) if response.status().is_success() => {
                    let raw = response.text().map_err(|e| e.to_string())?;
                    let value: serde_json::Value = serde_json::from_str(&raw).map_err(|e| format!("Drive upload response parse: {e} | {raw}"))?;
                    let _ = app.emit("gdrive-upload-progress", serde_json::json!({"bytesSent": file_size, "totalBytes": file_size}));
                    return Ok(serde_json::json!({
                        "ok": true,
                        "id": value.get("id").and_then(|x| x.as_str()).unwrap_or(""),
                        "name": value.get("name").and_then(|x| x.as_str()).unwrap_or(""),
                        "modifiedTime": value.get("modifiedTime").and_then(|x| x.as_str()).unwrap_or("")
                    }));
                }
                Ok(response) => {
                    let code = response.status();
                    let body = response.text().unwrap_or_default();
                    if attempt < 3 && (code.is_server_error() || code.as_u16() == 429) {
                        std::thread::sleep(std::time::Duration::from_secs(1 << attempt));
                        attempt += 1;
                        continue;
                    }
                    return Err(format!("HTTP_ERROR:{} | {}", code, body));
                }
                Err(error) if attempt < 3 => {
                    std::thread::sleep(std::time::Duration::from_secs(1 << attempt));
                    attempt += 1;
                    let _ = error;
                }
                Err(error) => return Err(format!("Drive upload failed: {error}")),
            }
        }
    }
    Err("Drive completed an upload without returning the uploaded file.".into())
}

fn gdrive_tmp_dir(app: &tauri::AppHandle) -> Result<std::path::PathBuf, String> {
    let d = gdrive_data_dir(app)?.join("gdrive-tmp");
    std::fs::create_dir_all(&d).map_err(|e| e.to_string())?;
    Ok(d)
}

#[derive(serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct CpresPackageMedia {
    id: String,
    temp_name: String,
    name: String,
    mime_type: String,
}

#[tauri::command(rename_all = "camelCase")]
async fn gdrive_sync_service_media(
    app: tauri::AppHandle,
    folder_id: String,
    media_files: Vec<CpresPackageMedia>,
) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || {
        use sha2::{Digest, Sha256};
        let folder = folder_id.trim().to_string();
        if folder.is_empty() { return Err("Google Drive media folder is unavailable.".into()); }
        let temp_dir = gdrive_tmp_dir(&app)?;
        let listing = gdrive_list_sync(&app, Some(folder.clone()))?;
        let existing = listing.get("files").and_then(|x| x.as_array()).cloned().unwrap_or_default();
        let mut synced = Vec::with_capacity(media_files.len());
        for media in media_files {
            if !media.temp_name.starts_with("cpres-upload-") || !media.temp_name.ends_with(".bin")
                || !media.temp_name.chars().all(|c| c.is_ascii_alphanumeric() || c == '-' || c == '.') {
                return Err("Invalid temporary media file name.".into());
            }
            let path = temp_dir.join(&media.temp_name);
            let mut source = std::fs::File::open(&path).map_err(|e| e.to_string())?;
            let mut hasher = Sha256::new();
            let mut buf = [0u8; 1024 * 1024];
            loop { let n = std::io::Read::read(&mut source, &mut buf).map_err(|e| e.to_string())?; if n == 0 { break; } hasher.update(&buf[..n]); }
            let hash = format!("{:x}", hasher.finalize());
            let ext = std::path::Path::new(&media.name).extension().and_then(|x| x.to_str()).filter(|x| x.chars().all(|c| c.is_ascii_alphanumeric())).unwrap_or("bin");
            let remote_name = format!("media-{hash}.{ext}");
            let found = existing.iter().find(|f| f.get("name").and_then(|x| x.as_str()) == Some(remote_name.as_str()));
            let (drive_id, modified) = if let Some(file) = found {
                (file.get("id").and_then(|x| x.as_str()).unwrap_or("").to_string(), file.get("modifiedTime").and_then(|x| x.as_str()).unwrap_or("").to_string())
            } else {
                let uploaded = gdrive_upload_resumable(&app, &path, &remote_name, &media.mime_type, Some(&folder), None, None)?;
                (uploaded.get("id").and_then(|x| x.as_str()).unwrap_or("").to_string(), uploaded.get("modifiedTime").and_then(|x| x.as_str()).unwrap_or("").to_string())
            };
            synced.push(serde_json::json!({"id":media.id,"driveFileId":drive_id,"contentHash":hash,"modifiedTime":modified}));
        }
        Ok(serde_json::json!({"ok":true,"files":synced}))
    }).await.map_err(|e| e.to_string())?
}

#[tauri::command(rename_all = "camelCase")]
async fn gdrive_upload_service_package(
    app: tauri::AppHandle,
    name: String,
    service_json: String,
    manifest_json: String,
    media_files: Vec<CpresPackageMedia>,
    folder_id: Option<String>,
    existing_file_id: Option<String>,
    cpres_id: Option<String>,
) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || {
        use std::io::Write;
        use zip::write::SimpleFileOptions;

        let temp_dir = gdrive_tmp_dir(&app)?;
        let mut staged = Vec::with_capacity(media_files.len());
        let mut total_size = service_json.len() as u64 + manifest_json.len() as u64;
        for media in &media_files {
            if !media.temp_name.starts_with("cpres-upload-")
                || !media.temp_name.ends_with(".bin")
                || !media.temp_name.chars().all(|c| c.is_ascii_alphanumeric() || c == '-' || c == '.')
            {
                return Err("Invalid temporary media file name.".into());
            }
            let path = temp_dir.join(&media.temp_name);
            total_size = total_size.saturating_add(std::fs::metadata(&path).map_err(|e| e.to_string())?.len());
            staged.push((media, path));
        }
        if total_size > 100 * 1024 * 1024 {
            return Err("Upload must be 100 MB or smaller.".into());
        }
        let _media_cleanup: Vec<GdriveTempFile> = staged.iter().map(|(_, path)| GdriveTempFile(path.clone())).collect();
        let package = GdriveTempFile::create(&app, "service-package")?;
        let file = std::fs::File::create(package.path()).map_err(|e| e.to_string())?;
        let mut archive = zip::ZipWriter::new(file);
        let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        archive.start_file("service.json", options).map_err(|e| e.to_string())?;
        archive.write_all(service_json.as_bytes()).map_err(|e| e.to_string())?;
        for (media, path) in &staged {
            archive.start_file(format!("media/{}", urlencoding_encode(&media.id)), options).map_err(|e| e.to_string())?;
            let mut source = std::fs::File::open(path).map_err(|e| e.to_string())?;
            std::io::copy(&mut source, &mut archive).map_err(|e| e.to_string())?;
        }
        archive.start_file("media-manifest.json", options).map_err(|e| e.to_string())?;
        archive.write_all(manifest_json.as_bytes()).map_err(|e| e.to_string())?;
        archive.finish().map_err(|e| e.to_string())?;
        drop(_media_cleanup);
        gdrive_upload_resumable(
            &app, package.path(), &name, "application/vnd.church-presenter.service+zip", folder_id.as_deref(),
            existing_file_id.as_deref(), cpres_id.as_deref(),
        )
    }).await.map_err(|e| e.to_string())?
}

struct GdriveTempFile(std::path::PathBuf);

impl GdriveTempFile {
    fn create(app: &tauri::AppHandle, prefix: &str) -> Result<Self, String> {
        use std::io::ErrorKind;
        let dir = gdrive_tmp_dir(app)?;
        for _ in 0..8 {
            let path = dir.join(format!(
                "{prefix}-{}-{}.bin",
                std::process::id(),
                random_url_safe(24)
            ));
            match std::fs::OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(&path)
            {
                Ok(_) => return Ok(Self(path)),
                Err(err) if err.kind() == ErrorKind::AlreadyExists => continue,
                Err(err) => return Err(err.to_string()),
            }
        }
        Err("Could not allocate a unique Google Drive transfer file.".into())
    }

    fn path(&self) -> &std::path::Path {
        &self.0
    }
}

impl Drop for GdriveTempFile {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}

// ── Status / connect / disconnect ──
fn gdrive_auth_status_sync(app: &tauri::AppHandle) -> GdriveStatus {
    let (fid, fname) = load_folder_config(app);
    match load_refresh_token(app) {
        Some(_) => {
            let email = get_valid_access_token(app).and_then(|a| fetch_user_email(&a)).ok();
            GdriveStatus {
                ok: true,
                connected: true,
                email,
                folder_id: fid,
                folder_name: fname,
                error: None,
                scope_note: Some("drive (read + write)".into()),
            }
        }
        None => GdriveStatus {
            ok: true,
            connected: false,
            email: None,
            folder_id: fid,
            folder_name: fname,
            error: None,
            scope_note: Some("Configure Client ID in app data gdrive-client-id.txt, then Connect Google.".into()),
        },
    }
}

#[tauri::command]
async fn gdrive_auth_status(app: tauri::AppHandle) -> GdriveStatus {
    tauri::async_runtime::spawn_blocking(move || gdrive_auth_status_sync(&app))
        .await
        .unwrap_or_default()
}

#[tauri::command]
async fn gdrive_connect(app: tauri::AppHandle) -> Result<GdriveStatus, String> {
    tauri::async_runtime::spawn_blocking(move || run_pkce_login(&app))
        .await
        .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn gdrive_disconnect(app: tauri::AppHandle) -> Result<GdriveStatus, String> {
    tauri::async_runtime::spawn_blocking(move || -> Result<GdriveStatus, String> {
        if let Some(refresh) = load_refresh_token(&app) {
            let body = format!("token={}", urlencoding_encode(&refresh));
            let _ = http_json("POST", GDRIVE_REVOKE_URL, Some(&body), None);
        }
        clear_refresh_token(&app);
        let (fid, fname) = load_folder_config(&app);
        Ok(GdriveStatus {
            ok: true,
            connected: false,
            email: None,
            folder_id: fid,
            folder_name: fname,
            error: None,
            scope_note: None,
        })
    })
    .await
    .map_err(|e| e.to_string())?
}

// ── Folder + listing ──
fn gdrive_pick_folder_sync(
    app: &tauri::AppHandle,
    folder_id: Option<String>,
    folder_name: Option<String>,
) -> Result<serde_json::Value, String> {
    let id = folder_id
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
        .ok_or_else(|| "Provide the Google Drive folder ID.".to_string())?;
    let access = get_valid_access_token(app)?;
    let url = format!(
        "https://www.googleapis.com/drive/v3/files/{}?fields=id,name,mimeType&supportsAllDrives=true",
        urlencoding_encode(&id)
    );
    let raw = http_json("GET", &url, None, Some(&access))?;
    let meta: serde_json::Value = serde_json::from_str(&raw).map_err(|e| format!("folder meta: {e} | {raw}"))?;
    if meta.get("error").is_some() {
        return Err(format!("Cannot access folder: {raw}"));
    }
    let name = folder_name
        .filter(|s| !s.trim().is_empty())
        .unwrap_or_else(|| {
            meta.get("name")
                .and_then(|n| n.as_str())
                .unwrap_or("Presentation folder")
                .to_string()
        });
    save_folder_config(app, &id, &name)?;
    Ok(serde_json::json!({ "ok": true, "folderId": id, "folderName": name }))
}

#[tauri::command]
async fn gdrive_pick_folder(
    app: tauri::AppHandle,
    folder_id: Option<String>,
    folder_name: Option<String>,
) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || gdrive_pick_folder_sync(&app, folder_id, folder_name))
        .await
        .map_err(|e| e.to_string())?
}

fn gdrive_list_sync(app: &tauri::AppHandle, folder_id: Option<String>) -> Result<serde_json::Value, String> {
    let access = get_valid_access_token(app)?;
    let (cfg_id, _) = load_folder_config(app);
    let fid = folder_id
        .filter(|s| !s.trim().is_empty())
        .or(cfg_id)
        .ok_or_else(|| "No folder configured. Set the presentation folder first.".to_string())?;
    let mut files = Vec::new();
    let mut page_token: Option<String> = None;
    loop {
        let q = urlencoding_encode(&format!("'{fid}' in parents and trashed=false"));
        let mut url = format!(
            "https://www.googleapis.com/drive/v3/files?q={q}&pageSize=100&orderBy=modifiedTime%20desc&fields=nextPageToken,files(id,name,mimeType,modifiedTime,size,thumbnailLink,appProperties)&supportsAllDrives=true&includeItemsFromAllDrives=true&corpora=allDrives"
        );
        if let Some(pt) = &page_token {
            url.push_str(&format!("&pageToken={}", urlencoding_encode(pt)));
        }
        let raw = http_json("GET", &url, None, Some(&access))?;
        let v: serde_json::Value = serde_json::from_str(&raw).map_err(|e| format!("list parse: {e} | {raw}"))?;
        if let Some(err) = v.get("error") {
            return Err(format!("Drive error: {err}"));
        }
        if let Some(arr) = v.get("files").and_then(|f| f.as_array()) {
            for f in arr {
                files.push(serde_json::json!({
                    "id": f.get("id").and_then(|x| x.as_str()).unwrap_or(""),
                    "name": f.get("name").and_then(|x| x.as_str()).unwrap_or("Untitled"),
                    "mimeType": f.get("mimeType").and_then(|x| x.as_str()).unwrap_or(""),
                    "modifiedTime": f.get("modifiedTime").and_then(|x| x.as_str()).unwrap_or(""),
                    "size": f.get("size").and_then(|x| x.as_str()).unwrap_or(""),
                    "thumbnailLink": f.get("thumbnailLink").and_then(|x| x.as_str()).unwrap_or(""),
                    "cpresId": f.get("appProperties").and_then(|x| x.get("churchPresenterServiceId")).and_then(|x| x.as_str()).unwrap_or(""),
                }));
            }
        }
        page_token = v.get("nextPageToken").and_then(|t| t.as_str()).map(|s| s.to_string());
        if page_token.is_none() || files.len() > 2000 {
            break;
        }
    }
    Ok(serde_json::json!({ "ok": true, "files": files, "folderId": fid }))
}

#[tauri::command]
async fn gdrive_list_presentations(app: tauri::AppHandle, folder_id: Option<String>) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || gdrive_list_sync(&app, folder_id))
        .await
        .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn gdrive_download_thumbnail(app: tauri::AppHandle, file_id: String) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || -> Result<serde_json::Value, String> {
        use base64::Engine;
        let access = get_valid_access_token(&app)?;
        let id = urlencoding_encode(&file_id);
        let url = format!("https://www.googleapis.com/drive/v3/files/{id}?fields=thumbnailLink&supportsAllDrives=true");
        let raw = http_json("GET", &url, None, Some(&access))?;
        let value: serde_json::Value = serde_json::from_str(&raw).map_err(|e| format!("thumbnail metadata parse: {e}"))?;
        let link = value.get("thumbnailLink").and_then(|x| x.as_str()).filter(|s| !s.is_empty()).ok_or_else(|| "Drive did not provide a thumbnail for this file.".to_string())?;
        let safe_id: String = file_id.chars().filter(|c| c.is_ascii_alphanumeric() || *c == '-' || *c == '_').collect();
        let path = gdrive_tmp_dir(&app)?.join(format!("thumbnail-{safe_id}.bin"));
        let _ = std::fs::remove_file(&path);
        http_download(link, &access, &path)?;
        let bytes = std::fs::read(&path).map_err(|e| e.to_string())?;
        let _ = std::fs::remove_file(&path);
        Ok(serde_json::json!({"ok":true,"dataBase64":base64::engine::general_purpose::STANDARD.encode(bytes)}))
    }).await.map_err(|e| e.to_string())?
}

// ── Sync: file info, download, upload ──
#[tauri::command]
async fn gdrive_file_info(app: tauri::AppHandle, file_id: String) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || -> Result<serde_json::Value, String> {
        let access = get_valid_access_token(&app)?;
        let url = format!(
            "https://www.googleapis.com/drive/v3/files/{}?fields=id,name,modifiedTime,trashed&supportsAllDrives=true",
            urlencoding_encode(&file_id)
        );
        let raw = http_json("GET", &url, None, Some(&access))?;
        let v: serde_json::Value = serde_json::from_str(&raw).map_err(|e| format!("info parse: {e} | {raw}"))?;
        if v.get("error").is_some() {
            return Err(format!("Drive error: {raw}"));
        }
        Ok(serde_json::json!({
            "ok": true,
            "id": v.get("id").and_then(|x| x.as_str()).unwrap_or(""),
            "name": v.get("name").and_then(|x| x.as_str()).unwrap_or(""),
            "modifiedTime": v.get("modifiedTime").and_then(|x| x.as_str()).unwrap_or(""),
            "trashed": v.get("trashed").and_then(|x| x.as_bool()).unwrap_or(false),
        }))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn gdrive_download_file(
    app: tauri::AppHandle,
    file_id: String,
    mime_type: Option<String>,
) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || -> Result<serde_json::Value, String> {
        use base64::Engine;
        let access = get_valid_access_token(&app)?;
        let id = urlencoding_encode(&file_id);
        let mime = mime_type.unwrap_or_default();
        let url = if mime == "application/vnd.google-apps.presentation" {
            format!(
                "https://www.googleapis.com/drive/v3/files/{id}/export?mimeType={}",
                urlencoding_encode("application/vnd.openxmlformats-officedocument.presentationml.presentation")
            )
        } else {
            format!("https://www.googleapis.com/drive/v3/files/{id}?alt=media&supportsAllDrives=true")
        };
        let tmp = GdriveTempFile::create(&app, "download")?;
        http_download(&url, &access, tmp.path())?;
        let bytes = std::fs::read(tmp.path()).map_err(|e| e.to_string())?;
        if bytes.len() > 200 * 1024 * 1024 {
            return Err("File is larger than 200 MB.".into());
        }
        Ok(serde_json::json!({
            "ok": true,
            "size": bytes.len(),
            "dataBase64": base64::engine::general_purpose::STANDARD.encode(&bytes),
        }))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn gdrive_upload_file(
    app: tauri::AppHandle,
    name: String,
    data_base64: String,
    mime_type: Option<String>,
    folder_id: Option<String>,
    existing_file_id: Option<String>,
    cpres_id: Option<String>,
) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || -> Result<serde_json::Value, String> {
        use base64::Engine;
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(data_base64)
            .map_err(|_| "Could not read upload data.".to_string())?;
        if bytes.len() > 100 * 1024 * 1024 {
            return Err("Upload must be 100 MB or smaller.".into());
        }
        let access = get_valid_access_token(&app)?;
        let boundary = "cpres_boundary_7f3a9c21";
        let existing = existing_file_id.map(|s| s.trim().to_string()).filter(|s| !s.is_empty());
        let mut metadata = serde_json::json!({ "name": name });
        if let Some(id) = cpres_id.map(|s| s.trim().to_string()).filter(|s| !s.is_empty()) {
            metadata["appProperties"] = serde_json::json!({"churchPresenterServiceId": id});
        }
        if existing.is_none() {
            if let Some(f) = folder_id.map(|s| s.trim().to_string()).filter(|s| !s.is_empty()) {
                metadata["parents"] = serde_json::json!([f]);
            }
        }
        let mime = mime_type.unwrap_or_else(|| "application/octet-stream".into());
        let mut body: Vec<u8> = Vec::new();
        body.extend_from_slice(
            format!(
                "--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{metadata}\r\n--{boundary}\r\nContent-Type: {mime}\r\n\r\n"
            )
            .as_bytes(),
        );
        body.extend_from_slice(&bytes);
        body.extend_from_slice(format!("\r\n--{boundary}--\r\n").as_bytes());
        let tmp = GdriveTempFile::create(&app, "upload")?;
        std::fs::write(tmp.path(), &body).map_err(|e| e.to_string())?;
        let (method, url) = match &existing {
            Some(id) => (
                "PATCH",
                format!(
                    "https://www.googleapis.com/upload/drive/v3/files/{}?uploadType=multipart&supportsAllDrives=true&fields=id,name,modifiedTime",
                    urlencoding_encode(id)
                ),
            ),
            None => (
                "POST",
                "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,name,modifiedTime".to_string(),
            ),
        };
        let res = http_upload(method, &url, &access, tmp.path(), &format!("multipart/related; boundary={boundary}"));
        let raw = res?;
        let v: serde_json::Value = serde_json::from_str(&raw).map_err(|e| format!("upload parse: {e} | {raw}"))?;
        if v.get("error").is_some() {
            return Err(format!("Drive error: {raw}"));
        }
        Ok(serde_json::json!({
            "ok": true,
            "id": v.get("id").and_then(|x| x.as_str()).unwrap_or(""),
            "name": v.get("name").and_then(|x| x.as_str()).unwrap_or(""),
            "modifiedTime": v.get("modifiedTime").and_then(|x| x.as_str()).unwrap_or(""),
        }))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn gdrive_ensure_subfolder(app: tauri::AppHandle, parent_id: String, name: String) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || -> Result<serde_json::Value, String> {
        let access = get_valid_access_token(&app)?;
        let parent = parent_id.trim().to_string();
        if parent.is_empty() {
            return Err("No parent folder id.".into());
        }
        let q = urlencoding_encode(&format!(
            "name='{}' and '{}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            name.replace('\'', "\\'"),
            parent
        ));
        let url = format!(
            "https://www.googleapis.com/drive/v3/files?q={q}&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true&corpora=allDrives"
        );
        let raw = http_json("GET", &url, None, Some(&access))?;
        let v: serde_json::Value = serde_json::from_str(&raw).map_err(|e| format!("folder parse: {e} | {raw}"))?;
        if v.get("error").is_some() {
            return Err(format!("Drive error: {raw}"));
        }
        if let Some(id) = v["files"][0]["id"].as_str() {
            return Ok(serde_json::json!({ "ok": true, "id": id }));
        }
        let body = serde_json::json!({
            "name": name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent]
        })
        .to_string();
        let tmp = gdrive_tmp_dir(&app)?.join("folder.json");
        std::fs::write(&tmp, body.as_bytes()).map_err(|e| e.to_string())?;
        let url = "https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&fields=id".to_string();
        let res = http_upload("POST", &url, &access, &tmp, "application/json");
        let _ = std::fs::remove_file(&tmp);
        let raw = res?;
        let v: serde_json::Value = serde_json::from_str(&raw).map_err(|e| format!("folder create parse: {e} | {raw}"))?;
        if v.get("error").is_some() {
            return Err(format!("Drive error: {raw}"));
        }
        Ok(serde_json::json!({ "ok": true, "id": v.get("id").and_then(|x| x.as_str()).unwrap_or("") }))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn gdrive_delete_file(app: tauri::AppHandle, file_id: String) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || -> Result<serde_json::Value, String> {
        let id = file_id.trim().to_string();
        if id.is_empty() {
            return Err("No file id provided.".into());
        }
        let access = get_valid_access_token(&app)?;
        let enc = urlencoding_encode(&id);

        // Ask Drive what this account is allowed to do with the file
        let info_url = format!(
            "https://www.googleapis.com/drive/v3/files/{enc}?fields=id,name,ownedByMe,owners(emailAddress),capabilities(canDelete,canTrash)&supportsAllDrives=true"
        );
        let raw = http_json("GET", &info_url, None, Some(&access))?;
        let info: serde_json::Value =
            serde_json::from_str(&raw).map_err(|e| format!("info parse: {e} | {raw}"))?;
        if info.get("error").is_some() {
            return Err(format!("Drive error: {raw}"));
        }
        let can_delete = info["capabilities"]["canDelete"].as_bool().unwrap_or(false);
        let can_trash = info["capabilities"]["canTrash"].as_bool().unwrap_or(false);

        if can_delete {
            let url = format!("https://www.googleapis.com/drive/v3/files/{enc}?supportsAllDrives=true");
            let out = http_json("DELETE", &url, None, Some(&access))?;
            if out.contains("\"error\"") {
                return Err(format!("Drive error: {out}"));
            }
            return Ok(serde_json::json!({ "ok": true, "method": "deleted" }));
        }

        if can_trash {
            let url = format!("https://www.googleapis.com/drive/v3/files/{enc}?supportsAllDrives=true&fields=id,trashed");
            let tmp = gdrive_tmp_dir(&app)?.join("trash.json");
            std::fs::write(&tmp, br#"{"trashed":true}"#).map_err(|e| e.to_string())?;
            let res = http_upload("PATCH", &url, &access, &tmp, "application/json");
            let _ = std::fs::remove_file(&tmp);
            let out = res?;
            if out.contains("\"error\"") {
                return Err(format!("Drive error: {out}"));
            }
            return Ok(serde_json::json!({ "ok": true, "method": "trashed" }));
        }

        let owner = info["owners"][0]["emailAddress"].as_str().unwrap_or("another account");
        let me = fetch_user_email(&access).unwrap_or_default();
        Err(format!(
            "The Google account signed in to this app ({me}) is not allowed to delete this file. It is owned by {owner}. Sign in with that account, or delete the file in Drive."
        ))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn gdrive_trash_file(app: tauri::AppHandle, file_id: String) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || -> Result<serde_json::Value, String> {
        let id = file_id.trim().to_string();
        if id.is_empty() {
            return Err("No file id provided.".into());
        }
        let access = get_valid_access_token(&app)?;
        let url = format!(
            "https://www.googleapis.com/drive/v3/files/{}?supportsAllDrives=true&fields=id,trashed",
            urlencoding_encode(&id)
        );
        let tmp = gdrive_tmp_dir(&app)?.join("trash.json");
        std::fs::write(&tmp, br#"{"trashed":true}"#).map_err(|e| e.to_string())?;
        let res = http_upload("PATCH", &url, &access, &tmp, "application/json");
        let _ = std::fs::remove_file(&tmp);
        let raw = res?;
        let v: serde_json::Value =
            serde_json::from_str(&raw).map_err(|e| format!("trash parse: {e} | {raw}"))?;
        if v.get("error").is_some() {
            return Err(format!("Drive error: {raw}"));
        }
        Ok(serde_json::json!({
            "ok": true,
            "id": v.get("id").and_then(|x| x.as_str()).unwrap_or(""),
            "trashed": v.get("trashed").and_then(|x| x.as_bool()).unwrap_or(true),
        }))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_deep_link::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init())
        .plugin(tauri_plugin_single_instance::init(|app, args, _cwd| {
            let paths: Vec<String> = args.iter().filter(|arg| arg.to_lowercase().ends_with(".cpres")).cloned().collect();
            if !paths.is_empty() {
                if let Some(state) = app.try_state::<PendingCpresFiles>() {
                    if let Ok(mut pending) = state.0.lock() { pending.extend(paths.clone()); }
                }
                let _ = app.emit("open-cpres-files", paths);
            }
        }))
        .manage(CloseGuard(AtomicBool::new(false)))
        .on_window_event(|window, event| {
            if let tauri::WindowEvent::CloseRequested { api, .. } = event {
                if window.label() == "main" {
                    let ready = window.state::<CloseGuard>().0.load(Ordering::SeqCst);
                    if ready {
                        api.prevent_close();
                        let _ = window.emit("confirm-close", ());
                    }
                }
            }
        })
        .setup(|app| {
            let app_data_dir = app.path().app_data_dir()?;
            std::fs::create_dir_all(&app_data_dir)?;
            std::fs::create_dir_all(app_data_dir.join("gdrive-tmp"))?;
            let client_id_path = app_data_dir.join("gdrive-client-id.txt");
            if !client_id_path.exists() && !BUNDLED_GDRIVE_CLIENT_ID.trim().is_empty() {
                std::fs::write(&client_id_path, BUNDLED_GDRIVE_CLIENT_ID.trim())?;
            }
            if let Some(secret) = BUNDLED_GDRIVE_CLIENT_SECRET {
                let secret_path = app_data_dir.join("gdrive-client-secret.txt");
                if !secret_path.exists() && !secret.trim().is_empty() {
                    std::fs::write(&secret_path, secret.trim())?;
                }
            }
            let database = database::Database::open(app_data_dir.join("church-presenter.sqlite"))?;
            app.manage(database);
            let startup_paths = std::env::args().skip(1).filter(|arg| arg.to_lowercase().ends_with(".cpres")).collect();
            app.manage(PendingCpresFiles(Mutex::new(startup_paths)));
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            take_pending_cpres_files,
            read_cpres_file,
            quit_app,
            close_guard_ready,
            wait_ppt_slideshow_end,
            stop_ppt_slideshow,
            focus_ppt_slideshow,
            load_tavily_key,
            save_tavily_key,
            store_background_media,
            delete_background_media,
            store_presentation_file,
            launch_ppt_slideshow,
            prepare_presentation,
            open_presentation_file,
            open_presentation_slideshow,
            database::list_services,
            database::load_service,
            database::save_service,
            database::delete_service,
            database::list_songs,
            database::save_song,
            database::delete_song,
            database::list_media,
            database::save_media,
            database::delete_media,
            database::list_slide_themes,
            database::save_slide_theme,
            database::delete_slide_theme,
            gdrive_auth_status,
            gdrive_connect,
            gdrive_disconnect,
            gdrive_pick_folder,
            gdrive_list_presentations,
            gdrive_download_thumbnail,
            gdrive_file_info,
            gdrive_download_file,
            gdrive_sync_service_media,
            gdrive_upload_service_package,
            gdrive_upload_file,
            gdrive_trash_file,
            gdrive_delete_file,
            gdrive_ensure_subfolder
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
