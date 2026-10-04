use std::path::Path;
use std::sync::{Mutex, MutexGuard};

use rusqlite::{params, Connection, Transaction};
use serde::{Deserialize, Serialize};
use tauri::State;

pub struct Database {
    connection: Mutex<Connection>,
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Service {
    pub id: String,
    pub title: String,
    pub date: String,
    #[serde(default)]
    pub folder: String,
    pub items: Vec<ServiceItem>,
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ServiceItem {
    pub id: String,
    pub title: String,
    #[serde(rename = "type")]
    pub item_type: ServiceItemType,
    pub slides: Vec<Slide>,
    #[serde(default)]
    pub groups: Vec<SlideGroup>,
    #[serde(default)]
    pub arrangements: Vec<Arrangement>,
    #[serde(default)]
    pub active_arrangement_id: Option<String>,
    #[serde(default)]
    pub reference: String,
    #[serde(default = "default_bible_version")]
    pub bible_version: String,
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideGroup {
    pub id: String,
    pub name: String,
    pub color: String,
    pub slide_ids: Vec<String>,
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Arrangement {
    pub id: String,
    pub name: String,
    pub group_ids: Vec<String>,
}

fn default_bible_version() -> String {
    "WEB".to_string()
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum ServiceItemType {
    Welcome,
    Song,
    Bible,
    Sermon,
    Announcement,
    Other,
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Slide {
    pub id: String,
    #[serde(rename = "type")]
    pub slide_type: String,
    pub text: String,
    pub background: String,
    pub text_align: String,
    pub font_size: u32,
    pub section_label: Option<String>,
    pub notes: Option<String>,
    pub background_media: Option<String>,
    #[serde(default = "default_transition")]
    pub transition: String,
    #[serde(default = "default_transition_duration")]
    pub transition_duration: u32,
    pub theme_id: Option<String>,
    #[serde(default)]
    pub objects: Vec<SlideObject>,
    #[serde(default)]
    pub cue_macro_ids: Vec<String>,
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideObject {
    pub id: String,
    pub kind: String,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub rotation: f64,
    pub opacity: f64,
    pub fill: String,
    pub stroke: String,
    pub stroke_width: f64,
    pub color: String,
    pub text: String,
    pub font_size: f64,
    pub text_align: String,
    pub source: Option<String>,
}

fn default_transition() -> String {
    "cut".to_string()
}

fn default_transition_duration() -> u32 {
    300
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideTheme {
    pub id: String,
    pub name: String,
    pub background: String,
    pub font_size: u32,
    pub text_align: String,
    pub transition: String,
    pub transition_duration: u32,
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Song {
    pub id: String,
    pub title: String,
    pub lyrics: String,
    #[serde(default = "default_reflow_mode")]
    pub reflow_mode: String,
}

fn default_reflow_mode() -> String {
    "markers".to_string()
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MediaAsset {
    pub id: String,
    pub name: String,
    pub kind: String,
    pub source: String,
    pub size: u64,
    pub added_at: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ServiceSummary {
    pub id: String,
    pub title: String,
    pub date: String,
    pub folder: String,
}

impl Database {
    pub fn open(path: impl AsRef<Path>) -> rusqlite::Result<Self> {
        let connection = Connection::open(path)?;
        connection.pragma_update(None, "foreign_keys", "ON")?;
        migrate(&connection)?;
        Ok(Self {
            connection: Mutex::new(connection),
        })
    }

    fn lock(&self) -> Result<MutexGuard<'_, Connection>, String> {
        self.connection
            .lock()
            .map_err(|_| "Database lock was poisoned".to_string())
    }
}

fn migrate(connection: &Connection) -> rusqlite::Result<()> {
    let mut version: u32 = connection.pragma_query_value(None, "user_version", |row| row.get(0))?;
    if version < 1 {
        connection.execute_batch(
            "BEGIN;
             CREATE TABLE services (
                 id TEXT PRIMARY KEY NOT NULL,
                 title TEXT NOT NULL,
                 date TEXT NOT NULL,
                 created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now')),
                 updated_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
             );
             CREATE TABLE service_items (
                 id TEXT PRIMARY KEY NOT NULL,
                 service_id TEXT NOT NULL REFERENCES services(id) ON DELETE CASCADE,
                 title TEXT NOT NULL,
                 item_type TEXT NOT NULL,
                 position INTEGER NOT NULL
             );
             CREATE INDEX service_items_order ON service_items(service_id, position);
             CREATE TABLE slides (
                 id TEXT PRIMARY KEY NOT NULL,
                 service_item_id TEXT NOT NULL REFERENCES service_items(id) ON DELETE CASCADE,
                 slide_type TEXT NOT NULL,
                 text TEXT NOT NULL,
                 background TEXT NOT NULL,
                 text_align TEXT NOT NULL,
                 font_size INTEGER NOT NULL,
                 position INTEGER NOT NULL
             );
             CREATE INDEX slides_order ON slides(service_item_id, position);
             PRAGMA user_version = 1;
             COMMIT;",
        )?;
        version = 1;
    }
    if version < 2 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE slides ADD COLUMN section_label TEXT;
             CREATE TABLE songs (
                 id TEXT PRIMARY KEY NOT NULL,
                 title TEXT NOT NULL,
                 lyrics TEXT NOT NULL,
                 created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
             );
             CREATE INDEX songs_title ON songs(title COLLATE NOCASE);
             PRAGMA user_version = 2;
             COMMIT;",
        )?;
        version = 2;
    }
    if version < 3 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE service_items ADD COLUMN reference TEXT NOT NULL DEFAULT '';
             ALTER TABLE service_items ADD COLUMN bible_version TEXT NOT NULL DEFAULT 'WEB';
             PRAGMA user_version = 3;
             COMMIT;",
        )?;
        version = 3;
    }
    if version < 4 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE slides ADD COLUMN background_media TEXT;
             PRAGMA user_version = 4;
             COMMIT;",
        )?;
        version = 4;
    }
    if version < 5 {
        connection.execute_batch(
            "BEGIN;
             CREATE TABLE media_assets (
                 id TEXT PRIMARY KEY NOT NULL,
                 name TEXT NOT NULL,
                 kind TEXT NOT NULL,
                 source TEXT NOT NULL,
                 size INTEGER NOT NULL,
                 added_at TEXT NOT NULL
             );
             CREATE INDEX media_assets_name ON media_assets(name COLLATE NOCASE);
             PRAGMA user_version = 5;
             COMMIT;",
        )?;
        version = 5;
    }
    if version < 6 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE services ADD COLUMN folder TEXT NOT NULL DEFAULT '';
             PRAGMA user_version = 6;
             COMMIT;",
        )?;
        version = 6;
    }
    if version < 7 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE songs ADD COLUMN reflow_mode TEXT NOT NULL DEFAULT 'markers';
             PRAGMA user_version = 7;
             COMMIT;",
        )?;
        version = 7;
    }
    if version < 8 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE slides ADD COLUMN notes TEXT;
             PRAGMA user_version = 8;
             COMMIT;",
        )?;
        version = 8;
    }
    if version < 9 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE slides ADD COLUMN transition TEXT NOT NULL DEFAULT 'cut';
             ALTER TABLE slides ADD COLUMN transition_duration INTEGER NOT NULL DEFAULT 300;
             PRAGMA user_version = 9;
             COMMIT;",
        )?;
        version = 9;
    }
    if version < 10 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE slides ADD COLUMN theme_id TEXT;
             CREATE TABLE slide_themes (
                 id TEXT PRIMARY KEY NOT NULL,
                 name TEXT NOT NULL,
                 background TEXT NOT NULL,
                 font_size INTEGER NOT NULL,
                 text_align TEXT NOT NULL,
                 transition TEXT NOT NULL,
                 transition_duration INTEGER NOT NULL,
                 updated_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
             );
             CREATE INDEX slide_themes_name ON slide_themes(name COLLATE NOCASE);
             PRAGMA user_version = 10;
             COMMIT;",
        )?;
    }
    if version < 11 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE service_items ADD COLUMN groups_json TEXT NOT NULL DEFAULT '[]';
             ALTER TABLE service_items ADD COLUMN arrangements_json TEXT NOT NULL DEFAULT '[]';
             ALTER TABLE service_items ADD COLUMN active_arrangement_id TEXT;
             PRAGMA user_version = 11;
             COMMIT;",
        )?;
    }
    if version < 12 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE slides ADD COLUMN objects_json TEXT NOT NULL DEFAULT '[]';
             PRAGMA user_version = 12;
             COMMIT;",
        )?;
        version = 12;
    }
    if version < 13 {
        connection.execute_batch(
            "BEGIN;
             ALTER TABLE slides ADD COLUMN cue_macro_ids_json TEXT NOT NULL DEFAULT '[]';
             PRAGMA user_version = 13;
             COMMIT;",
        )?;
    }
    Ok(())
}

fn item_type_name(item_type: &ServiceItemType) -> &'static str {
    match item_type {
        ServiceItemType::Welcome => "welcome",
        ServiceItemType::Song => "song",
        ServiceItemType::Bible => "bible",
        ServiceItemType::Sermon => "sermon",
        ServiceItemType::Announcement => "announcement",
        ServiceItemType::Other => "other",
    }
}

fn parse_item_type(value: &str) -> rusqlite::Result<ServiceItemType> {
    match value {
        "welcome" => Ok(ServiceItemType::Welcome),
        "song" => Ok(ServiceItemType::Song),
        "bible" => Ok(ServiceItemType::Bible),
        "sermon" => Ok(ServiceItemType::Sermon),
        "announcement" => Ok(ServiceItemType::Announcement),
        "other" => Ok(ServiceItemType::Other),
        _ => Err(rusqlite::Error::InvalidQuery),
    }
}

fn write_service(transaction: &Transaction<'_>, service: &Service) -> rusqlite::Result<()> {
    transaction.execute(
        "INSERT INTO services (id, title, date, folder) VALUES (?1, ?2, ?3, ?4)
         ON CONFLICT(id) DO UPDATE SET title = excluded.title, date = excluded.date, folder = excluded.folder,
         updated_at = strftime('%Y-%m-%dT%H:%M:%fZ', 'now')",
        params![service.id, service.title, service.date, service.folder],
    )?;
    transaction.execute(
        "DELETE FROM service_items WHERE service_id = ?1",
        [&service.id],
    )?;

    for (item_position, item) in service.items.iter().enumerate() {
        transaction.execute(
            "INSERT INTO service_items
             (id, service_id, title, item_type, position, reference, bible_version, groups_json, arrangements_json, active_arrangement_id)
             VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10)",
            params![
                item.id,
                service.id,
                item.title,
                item_type_name(&item.item_type),
                item_position as i64,
                item.reference,
                item.bible_version,
                serde_json::to_string(&item.groups).unwrap_or_else(|_| "[]".to_string()),
                serde_json::to_string(&item.arrangements).unwrap_or_else(|_| "[]".to_string()),
                item.active_arrangement_id
            ],
        )?;

        for (slide_position, slide) in item.slides.iter().enumerate() {
            transaction.execute(
                "INSERT INTO slides
                 (id, service_item_id, slide_type, text, background, text_align, font_size, position, section_label, notes, background_media, transition, transition_duration, theme_id, objects_json, cue_macro_ids_json)
                 VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16)",
                params![
                    slide.id,
                    item.id,
                    slide.slide_type,
                    slide.text,
                    slide.background,
                    slide.text_align,
                    slide.font_size,
                    slide_position as i64,
                    slide.section_label,
                    slide.notes,
                    slide.background_media,
                    slide.transition,
                    slide.transition_duration,
                    slide.theme_id,
                    serde_json::to_string(&slide.objects).unwrap_or_else(|_| "[]".to_string()),
                    serde_json::to_string(&slide.cue_macro_ids).unwrap_or_else(|_| "[]".to_string())
                ],
            )?;
        }
    }
    Ok(())
}

#[tauri::command]
pub fn list_services(database: State<'_, Database>) -> Result<Vec<ServiceSummary>, String> {
    let connection = database.lock()?;
    let mut statement = connection
        .prepare("SELECT id, title, date, folder FROM services ORDER BY date DESC, updated_at DESC")
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map([], |row| {
            Ok(ServiceSummary {
                id: row.get(0)?,
                title: row.get(1)?,
                date: row.get(2)?,
                folder: row.get(3)?,
            })
        })
        .map_err(|error| error.to_string())?;
    rows.collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|error| error.to_string())
}

#[tauri::command(rename_all = "camelCase")]
pub fn load_service(
    service_id: String,
    database: State<'_, Database>,
) -> Result<Service, String> {
    let connection = database.lock()?;
    let (title, date, folder): (String, String, String) = connection
        .query_row(
            "SELECT title, date, folder FROM services WHERE id = ?1",
            [&service_id],
            |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?)),
        )
        .map_err(|error| error.to_string())?;

    let mut item_statement = connection
        .prepare("SELECT id, title, item_type, reference, bible_version, groups_json, arrangements_json, active_arrangement_id FROM service_items WHERE service_id = ?1 ORDER BY position")
        .map_err(|error| error.to_string())?;
    let item_rows = item_statement
        .query_map([&service_id], |row| {
            Ok((
                row.get::<_, String>(0)?,
                row.get::<_, String>(1)?,
                row.get::<_, String>(2)?,
                row.get::<_, String>(3)?,
                row.get::<_, String>(4)?,
                row.get::<_, String>(5)?,
                row.get::<_, String>(6)?,
                row.get::<_, Option<String>>(7)?,
            ))
        })
        .map_err(|error| error.to_string())?;
    let item_data = item_rows
        .collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|error| error.to_string())?;
    drop(item_statement);

    let mut items = Vec::with_capacity(item_data.len());
    for (id, item_title, item_type, reference, bible_version, groups_json, arrangements_json, active_arrangement_id) in item_data {
        let mut slide_statement = connection
            .prepare(
                "SELECT id, slide_type, text, background, text_align, font_size, section_label, notes, background_media, transition, transition_duration, theme_id, objects_json, cue_macro_ids_json
                 FROM slides WHERE service_item_id = ?1 ORDER BY position",
            )
            .map_err(|error| error.to_string())?;
        let slide_rows = slide_statement
            .query_map([&id], |row| {
                Ok(Slide {
                    id: row.get(0)?,
                    slide_type: row.get(1)?,
                    text: row.get(2)?,
                    background: row.get(3)?,
                    text_align: row.get(4)?,
                    font_size: row.get(5)?,
                    section_label: row.get(6)?,
                    notes: row.get(7)?,
                    background_media: row.get(8)?,
                    transition: row.get(9)?,
                    transition_duration: row.get(10)?,
                    theme_id: row.get(11)?,
                    objects: serde_json::from_str(&row.get::<_, String>(12)?).unwrap_or_default(),
                    cue_macro_ids: serde_json::from_str(&row.get::<_, String>(13)?).unwrap_or_default(),
                })
            })
            .map_err(|error| error.to_string())?;
        let slides = slide_rows
            .collect::<rusqlite::Result<Vec<_>>>()
            .map_err(|error| error.to_string())?;
        items.push(ServiceItem {
            id,
            title: item_title,
            item_type: parse_item_type(&item_type).map_err(|error| error.to_string())?,
            slides,
            groups: serde_json::from_str(&groups_json).unwrap_or_default(),
            arrangements: serde_json::from_str(&arrangements_json).unwrap_or_default(),
            active_arrangement_id,
            reference,
            bible_version,
        });
    }

    Ok(Service {
        id: service_id,
        title,
        date,
        folder,
        items,
    })
}

#[tauri::command]
pub fn save_service(service: Service, database: State<'_, Database>) -> Result<(), String> {
    let mut connection = database.lock()?;
    let transaction = connection.transaction().map_err(|error| error.to_string())?;
    write_service(&transaction, &service).map_err(|error| error.to_string())?;
    transaction.commit().map_err(|error| error.to_string())
}

#[tauri::command(rename_all = "camelCase")]
pub fn delete_service(service_id: String, database: State<'_, Database>) -> Result<(), String> {
    database
        .lock()?
        .execute("DELETE FROM services WHERE id = ?1", [&service_id])
        .map(|_| ())
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn list_songs(database: State<'_, Database>) -> Result<Vec<Song>, String> {
    let connection = database.lock()?;
    let mut statement = connection
        .prepare("SELECT id, title, lyrics, reflow_mode FROM songs ORDER BY title COLLATE NOCASE")
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map([], |row| {
            Ok(Song {
                id: row.get(0)?,
                title: row.get(1)?,
                lyrics: row.get(2)?,
                reflow_mode: row.get(3)?,
            })
        })
        .map_err(|error| error.to_string())?;
    rows.collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn save_song(song: Song, database: State<'_, Database>) -> Result<(), String> {
    database
        .lock()?
        .execute(
            "INSERT INTO songs (id, title, lyrics, reflow_mode) VALUES (?1, ?2, ?3, ?4)
             ON CONFLICT(id) DO UPDATE SET title = excluded.title, lyrics = excluded.lyrics,
             reflow_mode = excluded.reflow_mode",
            params![song.id, song.title, song.lyrics, song.reflow_mode],
        )
        .map(|_| ())
        .map_err(|error| error.to_string())
}

#[tauri::command(rename_all = "camelCase")]
pub fn delete_song(song_id: String, database: State<'_, Database>) -> Result<(), String> {
    database
        .lock()?
        .execute("DELETE FROM songs WHERE id = ?1", [&song_id])
        .map(|_| ())
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn list_media(database: State<'_, Database>) -> Result<Vec<MediaAsset>, String> {
    let connection = database.lock()?;
    let mut statement = connection
        .prepare("SELECT id, name, kind, source, size, added_at FROM media_assets ORDER BY added_at DESC, name COLLATE NOCASE")
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map([], |row| {
            Ok(MediaAsset {
                id: row.get(0)?,
                name: row.get(1)?,
                kind: row.get(2)?,
                source: row.get(3)?,
                size: row.get(4)?,
                added_at: row.get(5)?,
            })
        })
        .map_err(|error| error.to_string())?;
    rows.collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn save_media(asset: MediaAsset, database: State<'_, Database>) -> Result<(), String> {
    database
        .lock()?
        .execute(
            "INSERT INTO media_assets (id, name, kind, source, size, added_at) VALUES (?1, ?2, ?3, ?4, ?5, ?6)
             ON CONFLICT(id) DO UPDATE SET name = excluded.name, kind = excluded.kind,
             source = excluded.source, size = excluded.size, added_at = excluded.added_at",
            params![asset.id, asset.name, asset.kind, asset.source, asset.size as i64, asset.added_at],
        )
        .map(|_| ())
        .map_err(|error| error.to_string())
}

#[tauri::command(rename_all = "camelCase")]
pub fn delete_media(media_id: String, database: State<'_, Database>) -> Result<(), String> {
    database
        .lock()?
        .execute("DELETE FROM media_assets WHERE id = ?1", [&media_id])
        .map(|_| ())
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn list_slide_themes(database: State<'_, Database>) -> Result<Vec<SlideTheme>, String> {
    let connection = database.lock()?;
    let mut statement = connection
        .prepare("SELECT id, name, background, font_size, text_align, transition, transition_duration FROM slide_themes ORDER BY name COLLATE NOCASE")
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map([], |row| {
            Ok(SlideTheme {
                id: row.get(0)?,
                name: row.get(1)?,
                background: row.get(2)?,
                font_size: row.get(3)?,
                text_align: row.get(4)?,
                transition: row.get(5)?,
                transition_duration: row.get(6)?,
            })
        })
        .map_err(|error| error.to_string())?;
    rows.collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn save_slide_theme(theme: SlideTheme, database: State<'_, Database>) -> Result<(), String> {
    let mut connection = database.lock()?;
    let transaction = connection.transaction().map_err(|error| error.to_string())?;
    transaction
        .execute(
            "INSERT INTO slide_themes (id, name, background, font_size, text_align, transition, transition_duration)
             VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)
             ON CONFLICT(id) DO UPDATE SET name = excluded.name, background = excluded.background,
             font_size = excluded.font_size, text_align = excluded.text_align, transition = excluded.transition,
             transition_duration = excluded.transition_duration,
             updated_at = strftime('%Y-%m-%dT%H:%M:%fZ', 'now')",
            params![theme.id, theme.name, theme.background, theme.font_size, theme.text_align, theme.transition, theme.transition_duration],
        )
        .map_err(|error| error.to_string())?;
    transaction
        .execute(
            "UPDATE slides SET background = ?1, font_size = ?2, text_align = ?3, transition = ?4, transition_duration = ?5 WHERE theme_id = ?6",
            params![theme.background, theme.font_size, theme.text_align, theme.transition, theme.transition_duration, theme.id],
        )
        .map_err(|error| error.to_string())?;
    transaction.commit().map_err(|error| error.to_string())
}

#[tauri::command(rename_all = "camelCase")]
pub fn delete_slide_theme(theme_id: String, database: State<'_, Database>) -> Result<(), String> {
    let mut connection = database.lock()?;
    let transaction = connection.transaction().map_err(|error| error.to_string())?;
    transaction
        .execute("UPDATE slides SET theme_id = NULL WHERE theme_id = ?1", [&theme_id])
        .map_err(|error| error.to_string())?;
    transaction
        .execute("DELETE FROM slide_themes WHERE id = ?1", [&theme_id])
        .map_err(|error| error.to_string())?;
    transaction.commit().map_err(|error| error.to_string())
}
