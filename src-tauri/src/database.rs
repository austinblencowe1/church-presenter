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
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ServiceSummary {
    pub id: String,
    pub title: String,
    pub date: String,
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
    let version: u32 = connection.pragma_query_value(None, "user_version", |row| row.get(0))?;
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
        "INSERT INTO services (id, title, date) VALUES (?1, ?2, ?3)
         ON CONFLICT(id) DO UPDATE SET title = excluded.title, date = excluded.date,
         updated_at = strftime('%Y-%m-%dT%H:%M:%fZ', 'now')",
        params![service.id, service.title, service.date],
    )?;
    transaction.execute(
        "DELETE FROM service_items WHERE service_id = ?1",
        [&service.id],
    )?;

    for (item_position, item) in service.items.iter().enumerate() {
        transaction.execute(
            "INSERT INTO service_items (id, service_id, title, item_type, position)
             VALUES (?1, ?2, ?3, ?4, ?5)",
            params![
                item.id,
                service.id,
                item.title,
                item_type_name(&item.item_type),
                item_position as i64
            ],
        )?;

        for (slide_position, slide) in item.slides.iter().enumerate() {
            transaction.execute(
                "INSERT INTO slides
                 (id, service_item_id, slide_type, text, background, text_align, font_size, position)
                 VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)",
                params![
                    slide.id,
                    item.id,
                    slide.slide_type,
                    slide.text,
                    slide.background,
                    slide.text_align,
                    slide.font_size,
                    slide_position as i64
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
        .prepare("SELECT id, title, date FROM services ORDER BY date DESC, updated_at DESC")
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map([], |row| {
            Ok(ServiceSummary {
                id: row.get(0)?,
                title: row.get(1)?,
                date: row.get(2)?,
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
    let (title, date): (String, String) = connection
        .query_row(
            "SELECT title, date FROM services WHERE id = ?1",
            [&service_id],
            |row| Ok((row.get(0)?, row.get(1)?)),
        )
        .map_err(|error| error.to_string())?;

    let mut item_statement = connection
        .prepare("SELECT id, title, item_type FROM service_items WHERE service_id = ?1 ORDER BY position")
        .map_err(|error| error.to_string())?;
    let item_rows = item_statement
        .query_map([&service_id], |row| {
            Ok((row.get::<_, String>(0)?, row.get::<_, String>(1)?, row.get::<_, String>(2)?))
        })
        .map_err(|error| error.to_string())?;
    let item_data = item_rows
        .collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|error| error.to_string())?;
    drop(item_statement);

    let mut items = Vec::with_capacity(item_data.len());
    for (id, item_title, item_type) in item_data {
        let mut slide_statement = connection
            .prepare(
                "SELECT id, slide_type, text, background, text_align, font_size
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
        });
    }

    Ok(Service {
        id: service_id,
        title,
        date,
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