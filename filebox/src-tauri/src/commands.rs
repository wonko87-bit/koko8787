use crate::model::{
    ext_of, new_id, now_millis, stem_of, tokenize, EntryStatus, Favorite, FileEntry, MoveRecord,
    Rule, Settings, Suggestion,
};
use crate::store::{move_file, unique_dest, Store};
use crate::watcher::{Msg, WatcherCtl};
use serde::Serialize;
use std::path::PathBuf;
use tauri::{AppHandle, Emitter, State};

const MAX_RECORDS: usize = 500;

// ---------- 설정 ----------

#[tauri::command]
pub fn get_settings(store: State<Store>) -> Option<Settings> {
    store.read(|d| d.settings.clone())
}

#[tauri::command]
pub fn update_settings(
    app: AppHandle,
    store: State<Store>,
    watcher: State<WatcherCtl>,
    settings: Settings,
) -> Settings {
    let dirs_changed = store.update(|d| {
        let changed = d
            .settings
            .as_ref()
            .map(|old| old.watch_dirs != settings.watch_dirs)
            .unwrap_or(true);
        d.settings = Some(settings.clone());
        changed
    });
    if dirs_changed {
        if let Ok(tx) = watcher.0.lock() {
            let _ = tx.send(Msg::Rewatch);
        }
    }
    let _ = app.emit("settings-changed", ());
    settings
}

// ---------- 항목 ----------

#[tauri::command]
pub fn list_entries(store: State<Store>) -> Vec<FileEntry> {
    store.read(|d| d.entries.clone())
}

#[tauri::command]
pub fn get_suggestions(store: State<Store>, entry_id: String) -> Vec<Suggestion> {
    store.read(|d| {
        d.entries
            .iter()
            .find(|e| e.id == entry_id)
            .map(|e| crate::suggest::suggest(&d.rules, &d.records, &d.favorites, &e.file_name, 3))
            .unwrap_or_default()
    })
}

fn file_entry_to(
    app: &AppHandle,
    store: &Store,
    entry_id: &str,
    dest_dir: PathBuf,
    learn_favorite: Option<String>,
) -> Result<FileEntry, String> {
    let entry = store
        .read(|d| d.entries.iter().find(|e| e.id == entry_id).cloned())
        .ok_or("항목을 찾을 수 없습니다")?;
    if entry.status == EntryStatus::Filed {
        return Err("이미 이동이 완료된 항목입니다".into());
    }
    if !dest_dir.is_dir() {
        return Err(format!("대상 폴더가 없습니다: {}", dest_dir.display()));
    }

    let dest = unique_dest(&dest_dir, &entry.file_name);
    move_file(&entry.path, &dest).map_err(|e| format!("이동 실패: {e}"))?;

    let updated = store.update(|d| {
        if let Some(fav_id) = &learn_favorite {
            d.records.push(MoveRecord {
                ext: ext_of(&entry.file_name),
                tokens: tokenize(&stem_of(&entry.file_name)),
                favorite_id: fav_id.clone(),
                at: now_millis(),
            });
            let len = d.records.len();
            if len > MAX_RECORDS {
                d.records.drain(0..len - MAX_RECORDS);
            }
        }
        d.entries.iter_mut().find(|e| e.id == entry_id).map(|e| {
            e.status = EntryStatus::Filed;
            e.path = dest.clone();
            e.filed_to = Some(dest_dir.clone());
            e.filed_at = Some(now_millis());
            e.clone()
        })
    });
    let _ = app.emit("entries-changed", ());
    updated.ok_or_else(|| "항목 갱신 실패".into())
}

#[tauri::command]
pub fn send_to_favorite(
    app: AppHandle,
    store: State<Store>,
    entry_id: String,
    favorite_id: String,
) -> Result<FileEntry, String> {
    let fav = store
        .read(|d| d.favorites.iter().find(|f| f.id == favorite_id).cloned())
        .ok_or("즐겨찾기를 찾을 수 없습니다")?;
    file_entry_to(&app, &store, &entry_id, fav.path, Some(fav.id))
}

#[tauri::command]
pub fn send_to_path(
    app: AppHandle,
    store: State<Store>,
    entry_id: String,
    dest_dir: String,
) -> Result<FileEntry, String> {
    let dest = PathBuf::from(dest_dir);
    // 같은 경로의 즐겨찾기가 있으면 학습에도 반영
    let fav_id = store.read(|d| {
        d.favorites.iter().find(|f| f.path == dest).map(|f| f.id.clone())
    });
    file_entry_to(&app, &store, &entry_id, dest, fav_id)
}

#[tauri::command]
pub fn set_category(
    app: AppHandle,
    store: State<Store>,
    entry_id: String,
    category: String,
) -> Result<(), String> {
    let found = store.update(|d| {
        d.entries
            .iter_mut()
            .find(|e| e.id == entry_id)
            .map(|e| e.category = category)
            .is_some()
    });
    let _ = app.emit("entries-changed", ());
    found.then_some(()).ok_or("항목을 찾을 수 없습니다".into())
}

#[tauri::command]
pub fn set_tags(
    app: AppHandle,
    store: State<Store>,
    entry_id: String,
    tags: Vec<String>,
) -> Result<(), String> {
    let found = store.update(|d| {
        d.entries
            .iter_mut()
            .find(|e| e.id == entry_id)
            .map(|e| e.tags = tags)
            .is_some()
    });
    let _ = app.emit("entries-changed", ());
    found.then_some(()).ok_or("항목을 찾을 수 없습니다".into())
}

/// 목록에서만 제거 (파일은 디스크에 그대로 남음)
#[tauri::command]
pub fn remove_entry(app: AppHandle, store: State<Store>, entry_id: String) {
    store.update(|d| d.entries.retain(|e| e.id != entry_id));
    let _ = app.emit("entries-changed", ());
}

/// 이동 완료(Filed) 기록 전체 삭제. 학습 데이터(records)는 유지된다.
#[tauri::command]
pub fn clear_history(app: AppHandle, store: State<Store>) {
    store.update(|d| d.entries.retain(|e| e.status != EntryStatus::Filed));
    let _ = app.emit("entries-changed", ());
}

// ---------- 즐겨찾기 ----------

#[tauri::command]
pub fn list_favorites(store: State<Store>) -> Vec<Favorite> {
    store.read(|d| d.favorites.clone())
}

#[tauri::command]
pub fn add_favorite(
    app: AppHandle,
    store: State<Store>,
    name: String,
    path: String,
) -> Result<Favorite, String> {
    let path = PathBuf::from(path);
    if !path.is_dir() {
        return Err("폴더가 존재하지 않습니다".into());
    }
    let fav = Favorite { id: new_id(), name, path };
    store.update(|d| d.favorites.push(fav.clone()));
    let _ = app.emit("favorites-changed", ());
    Ok(fav)
}

#[tauri::command]
pub fn remove_favorite(app: AppHandle, store: State<Store>, favorite_id: String) {
    store.update(|d| {
        d.favorites.retain(|f| f.id != favorite_id);
        d.records.retain(|r| r.favorite_id != favorite_id);
        for rule in d.rules.iter_mut() {
            if rule.favorite_id.as_deref() == Some(favorite_id.as_str()) {
                rule.favorite_id = None;
            }
        }
    });
    let _ = app.emit("favorites-changed", ());
}

// ---------- 규칙 ----------

#[tauri::command]
pub fn list_rules(store: State<Store>) -> Vec<Rule> {
    store.read(|d| d.rules.clone())
}

#[tauri::command]
pub fn upsert_rule(app: AppHandle, store: State<Store>, mut rule: Rule) -> Rule {
    if rule.id.is_empty() {
        rule.id = new_id();
    }
    rule.extensions = rule
        .extensions
        .iter()
        .map(|e| e.trim().trim_start_matches('.').to_lowercase())
        .filter(|e| !e.is_empty())
        .collect();
    rule.keywords = rule
        .keywords
        .iter()
        .map(|k| k.trim().to_lowercase())
        .filter(|k| !k.is_empty())
        .collect();
    store.update(|d| {
        if let Some(existing) = d.rules.iter_mut().find(|r| r.id == rule.id) {
            *existing = rule.clone();
        } else {
            d.rules.push(rule.clone());
        }
    });
    let _ = app.emit("rules-changed", ());
    rule
}

#[tauri::command]
pub fn remove_rule(app: AppHandle, store: State<Store>, rule_id: String) {
    store.update(|d| d.rules.retain(|r| r.id != rule_id));
    let _ = app.emit("rules-changed", ());
}

// ---------- 수동 수집 ----------

#[derive(Serialize)]
pub struct Candidate {
    pub path: String,
    pub name: String,
    pub size: u64,
}

/// 감시 폴더에 남아 있는 (아직 수집되지 않은) 파일 목록
#[tauri::command]
pub fn list_uncollected(store: State<Store>) -> Vec<Candidate> {
    let dirs = store.read(|d| {
        d.settings
            .as_ref()
            .map(|s| s.watch_dirs.clone())
            .unwrap_or_default()
    });
    let mut out = Vec::new();
    for dir in dirs {
        let Ok(read) = std::fs::read_dir(&dir) else { continue };
        for e in read.flatten() {
            let path = e.path();
            let Ok(meta) = e.metadata() else { continue };
            if !meta.is_file() {
                continue;
            }
            let name = e.file_name().to_string_lossy().to_string();
            if crate::rules::is_temp_download(&name) {
                continue;
            }
            out.push(Candidate {
                path: path.to_string_lossy().to_string(),
                name,
                size: meta.len(),
            });
        }
    }
    out.sort_by(|a, b| a.name.cmp(&b.name));
    out
}

#[tauri::command]
pub fn collect_paths(app: AppHandle, paths: Vec<String>) -> usize {
    let mut count = 0;
    for p in paths {
        if crate::watcher::collect_path(&app, std::path::Path::new(&p)).is_some() {
            count += 1;
        }
    }
    count
}

// ---------- 열기 / 표시 ----------

#[tauri::command]
pub fn open_entry(store: State<Store>, entry_id: String) -> Result<(), String> {
    let path = store
        .read(|d| d.entries.iter().find(|e| e.id == entry_id).map(|e| e.path.clone()))
        .ok_or("항목을 찾을 수 없습니다")?;
    tauri_plugin_opener::open_path(path, None::<&str>).map_err(|e| e.to_string())
}

#[tauri::command]
pub fn reveal_entry(store: State<Store>, entry_id: String) -> Result<(), String> {
    let path = store
        .read(|d| d.entries.iter().find(|e| e.id == entry_id).map(|e| e.path.clone()))
        .ok_or("항목을 찾을 수 없습니다")?;
    tauri_plugin_opener::reveal_item_in_dir(path).map_err(|e| e.to_string())
}

// ---------- 토스트 ----------

#[tauri::command]
pub fn hide_toast(app: AppHandle) {
    crate::toast::hide_toast(&app);
}

#[tauri::command]
pub fn show_main_window(app: AppHandle) {
    use tauri::Manager;
    if let Some(win) = app.get_webview_window("main") {
        let _ = win.show();
        let _ = win.unminimize();
        let _ = win.set_focus();
    }
}

/// 임의의 폴더 경로를 탐색기로 연다. (즐겨찾기 / 감시 폴더 / 관리함 폴더 열기용)
#[tauri::command]
pub fn open_folder(path: String) -> Result<(), String> {
    let path = PathBuf::from(path);
    if !path.is_dir() {
        return Err(format!("폴더를 찾을 수 없습니다: {}", path.display()));
    }
    tauri_plugin_opener::open_path(path, None::<&str>).map_err(|e| e.to_string())
}
