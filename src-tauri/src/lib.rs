mod database;
use tauri::Manager;

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

#[tauri::command]
fn setup_ollama() -> Result<(), String> {
    #[cfg(windows)]
    {
        let script = include_str!("../scripts/install-ollama.ps1");
        let mut command = std::process::Command::new("powershell.exe");
        command.args([
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ]);
        use std::os::windows::process::CommandExt;
        command.creation_flags(0x08000000);
        let output = command.output().map_err(|error| error.to_string())?;
        if output.status.success() {
            Ok(())
        } else {
            let details = String::from_utf8_lossy(&output.stderr).trim().to_string();
            Err(if details.is_empty() {
                format!("Ollama setup exited with status {}", output.status)
            } else {
                details
            })
        }
    }
    #[cfg(not(windows))]
    {
        Err("The guided Ollama installer is currently available on Windows.".into())
    }
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
    if !valid_extension || id.is_empty() || !id.chars().all(|c| c.is_ascii_alphanumeric() || c == '-') {
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

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .setup(|app| {
            let app_data_dir = app.path().app_data_dir()?;
            std::fs::create_dir_all(&app_data_dir)?;
            let database = database::Database::open(app_data_dir.join("church-presenter.sqlite"))?;
            app.manage(database);
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            load_tavily_key,
            save_tavily_key,
            setup_ollama,
            store_background_media,
            delete_background_media,
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
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
