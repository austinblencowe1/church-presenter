mod database;
use tauri::Manager;

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
            database::list_services,
            database::load_service,
            database::save_service,
            database::delete_service,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
