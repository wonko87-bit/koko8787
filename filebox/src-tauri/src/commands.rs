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
        let mut record_id = None;
        if let Some(fav_id) = &learn_favorite {
            let id = new_id();
            d.records.push(MoveRecord {
                id: id.clone(),
                ext: ext_of(&entry.file_name),
                tokens: tokenize(&stem_of(&entry.file_name)),
                favorite_id: fav_id.clone(),
                at: now_millis(),
            });
            let len = d.records.len();
            if len > MAX_RECORDS {
                d.records.drain(0..len - MAX_RECORDS);
            }
            record_id = Some(id);
        }
        d.entries.iter_mut().find(|e| e.id == entry_id).map(|e| {
            e.status = EntryStatus::Filed;
            e.path = dest.clone();
            e.filed_to = Some(dest_dir.clone());
            e.filed_at = Some(now_millis());
            e.record_id = record_id;
            e.clone()
        })
    });

    // 갈 곳이 정해진 지금이 할일을 만들 때다. 메모에 적히는 경로가 최종 경로이므로
    // 나중에 어긋나지 않는다. 관리함에 있을 때 보내면 이동하는 순간 죽는 경로가 된다.
    let updated = updated.map(|entry| register_with_flowdeck(store, entry));
    updated.ok_or_else(|| "항목 갱신 실패".into())
}

/// 이동이 끝난 항목을 Flowdeck 에 등록한다.
///
/// 규칙에 걸리면 그 규칙의 설정으로, 걸리지 않아도 사용자가 관리함에서 미리
/// 표시해 뒀으면 기한 없는 할일로 보낸다. 실패해도 이동 자체는 이미 끝난 일이라
/// 여기서 되돌리지 않는다.
fn register_with_flowdeck(store: &Store, entry: FileEntry) -> FileEntry {
    let sent = crate::flowdeck::dispatch(store, &entry);
    if sent.is_empty() {
        return entry;
    }
    let updated = store.update(|d| {
        d.entries.iter_mut().find(|e| e.id == entry.id).map(|e| {
            for todo in &sent {
                e.flowdeck_todos.retain(|t| t.todo_id != todo.todo_id);
                e.flowdeck_todos.push(todo.clone());
            }
            // 예약은 소비됐다.
            e.flowdeck_pending = false;
            e.clone()
        })
    });
    updated.unwrap_or(entry)
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
    let filed = file_entry_to(&store, &entry_id, fav.path, Some(fav.id));
    let _ = app.emit("entries-changed", ());
    filed
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
    let filed = file_entry_to(&store, &entry_id, dest, fav_id);
    let _ = app.emit("entries-changed", ());
    filed
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

/// 여러 항목을 목록에서 제거 (파일은 그대로)
#[tauri::command]
pub fn remove_entries(app: AppHandle, store: State<Store>, entry_ids: Vec<String>) {
    store.update(|d| d.entries.retain(|e| !entry_ids.contains(&e.id)));
    let _ = app.emit("entries-changed", ());
}

/// 여러 항목의 카테고리를 한 번에 변경
#[tauri::command]
pub fn set_category_many(
    app: AppHandle,
    store: State<Store>,
    entry_ids: Vec<String>,
    category: String,
) {
    store.update(|d| {
        for e in d.entries.iter_mut().filter(|e| entry_ids.contains(&e.id)) {
            e.category = category.clone();
        }
    });
    let _ = app.emit("entries-changed", ());
}

#[derive(Serialize)]
pub struct BatchResult {
    pub moved: usize,
    pub errors: Vec<String>,
}

fn move_many(
    store: &Store,
    entry_ids: Vec<String>,
    dest_dir: PathBuf,
    favorite_id: Option<String>,
) -> BatchResult {
    let mut moved = 0;
    let mut errors = Vec::new();
    for id in entry_ids {
        let name = store.read(|d| {
            d.entries.iter().find(|e| e.id == id).map(|e| e.file_name.clone())
        });
        match file_entry_to(store, &id, dest_dir.clone(), favorite_id.clone()) {
            Ok(_) => moved += 1,
            Err(e) => errors.push(format!("{}: {e}", name.unwrap_or(id))),
        }
    }
    BatchResult { moved, errors }
}

/// 여러 항목을 한 즐겨찾기로 일괄 이동
#[tauri::command]
pub fn send_many_to_favorite(
    app: AppHandle,
    store: State<Store>,
    entry_ids: Vec<String>,
    favorite_id: String,
) -> Result<BatchResult, String> {
    let fav = store
        .read(|d| d.favorites.iter().find(|f| f.id == favorite_id).cloned())
        .ok_or("즐겨찾기를 찾을 수 없습니다")?;
    let result = move_many(&store, entry_ids, fav.path, Some(fav.id));
    let _ = app.emit("entries-changed", ());
    Ok(result)
}

/// 여러 항목을 임의의 폴더로 일괄 이동
#[tauri::command]
pub fn send_many_to_path(
    app: AppHandle,
    store: State<Store>,
    entry_ids: Vec<String>,
    dest_dir: String,
) -> BatchResult {
    let dest = PathBuf::from(dest_dir);
    let fav_id = store.read(|d| {
        d.favorites.iter().find(|f| f.path == dest).map(|f| f.id.clone())
    });
    let result = move_many(&store, entry_ids, dest, fav_id);
    let _ = app.emit("entries-changed", ());
    result
}

/// 이동을 되돌린다: 파일을 관리함으로 되돌리고, 그때 남긴 학습 기록도 제거한다.
#[tauri::command]
pub fn undo_move(
    app: AppHandle,
    store: State<Store>,
    entry_id: String,
) -> Result<FileEntry, String> {
    let (entry, inbox) = store.read(|d| {
        (
            d.entries.iter().find(|e| e.id == entry_id).cloned(),
            d.settings.as_ref().map(|s| s.inbox_dir.clone()),
        )
    });
    let entry = entry.ok_or("항목을 찾을 수 없습니다")?;
    let inbox = inbox.ok_or("관리함 폴더가 설정되지 않았습니다")?;
    if entry.status != EntryStatus::Filed {
        return Err("이미 관리함에 있는 항목입니다".into());
    }
    if !entry.path.is_file() {
        return Err("파일이 이동되었거나 삭제되어 되돌릴 수 없습니다".into());
    }
    std::fs::create_dir_all(&inbox).map_err(|e| format!("관리함 폴더 생성 실패: {e}"))?;

    let dest = unique_dest(&inbox, &entry.file_name);
    move_file(&entry.path, &dest).map_err(|e| format!("되돌리기 실패: {e}"))?;

    let updated = store.update(|d| {
        if let Some(rid) = &entry.record_id {
            d.records.retain(|r| &r.id != rid);
        }
        d.entries.iter_mut().find(|e| e.id == entry_id).map(|e| {
            e.status = EntryStatus::Inbox;
            e.path = dest.clone();
            e.file_name = dest
                .file_name()
                .map(|n| n.to_string_lossy().to_string())
                .unwrap_or_else(|| e.file_name.clone());
            e.filed_to = None;
            e.filed_at = None;
            e.record_id = None;
            e.clone()
        })
    });
    let _ = app.emit("entries-changed", ());
    updated.ok_or_else(|| "항목 갱신 실패".into())
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

// ---------- Flowdeck 연동 ----------

#[derive(Serialize)]
pub struct FlowdeckStatus {
    /// 지금 쓰는 감시 폴더. 연동이 꺼져 있으면 None.
    pub inbox: Option<String>,
    /// 그 폴더가 실제로 있는지. Flowdeck 을 한 번도 켜지 않았으면 없다.
    pub exists: bool,
    /// 특별규칙이 붙은 규칙 수
    pub rule_count: usize,
}

#[tauri::command]
pub fn flowdeck_status(store: State<Store>) -> FlowdeckStatus {
    let inbox = crate::flowdeck::inbox_of(&store);
    let rule_count = store.read(|d| d.rules.iter().filter(|r| r.flowdeck.is_some()).count());
    FlowdeckStatus {
        exists: inbox.as_ref().is_some_and(|p| p.is_dir()),
        inbox: inbox.map(|p| p.to_string_lossy().to_string()),
        rule_count,
    }
}

/// 관리함에 있는 파일을 "옮길 때 Flowdeck 에 등록" 으로 표시하거나 해제한다.
///
/// 지금 바로 보내지 않는 이유는 하나다. 관리함은 거쳐 가는 곳이라 지금 보내면
/// 메모에 적히는 경로가 곧 죽는다. 표시만 해 두고 갈 곳이 정해질 때 보낸다.
#[tauri::command]
pub fn set_flowdeck_pending(
    app: AppHandle,
    store: State<Store>,
    entry_id: String,
    pending: bool,
) -> Result<FileEntry, String> {
    let updated = store.update(|d| {
        d.entries.iter_mut().find(|e| e.id == entry_id).map(|e| {
            e.flowdeck_pending = pending;
            e.clone()
        })
    });
    let _ = app.emit("entries-changed", ());
    updated.ok_or_else(|| "항목을 찾을 수 없습니다".into())
}

/// 이미 정리가 끝난 파일을 지금 Flowdeck 할일로 보낸다.
///
/// 이 파일은 최종 폴더에 있으므로 메모에 적히는 경로가 그대로 유효하다. 규칙을
/// 만들 만큼 반복되지 않는 파일을 위한 길이다.
#[tauri::command]
pub fn send_to_flowdeck(
    app: AppHandle,
    store: State<Store>,
    entry_id: String,
) -> Result<String, String> {
    let inbox = crate::flowdeck::inbox_of(&store)
        .ok_or("Flowdeck 연동이 꺼져 있습니다. 설정에서 켜 주세요.")?;
    let entry = store
        .read(|d| d.entries.iter().find(|e| e.id == entry_id).cloned())
        .ok_or("항목을 찾을 수 없습니다.")?;

    // 관리함 파일을 지금 보내면 곧 죽을 경로가 메모에 박힌다. 그쪽은 예약을 쓴다.
    if entry.status == EntryStatus::Inbox {
        return Err(
            "아직 관리함에 있는 파일입니다. 폴더로 옮기면 자동으로 등록됩니다.".into(),
        );
    }

    let rules = store.read(|d| d.rules.clone());
    let matched = crate::flowdeck::matching_specs(&rules, &entry.file_name);
    let (spec, rule_id, rule_name) = match matched.first() {
        Some(r) => (
            r.flowdeck.clone().unwrap_or_default(),
            r.id.clone(),
            r.name.clone(),
        ),
        None => (
            crate::model::FlowdeckSpec {
                due_in_days: None,
                ..Default::default()
            },
            String::new(),
            String::new(),
        ),
    };

    let todo = crate::flowdeck::send(
        &inbox,
        &spec,
        &entry.id,
        &rule_id,
        &rule_name,
        &entry.file_name,
        &entry.path,
        &entry.category,
    )
    .map_err(|e| format!("Flowdeck 폴더에 쓰지 못했습니다: {e}"))?;

    store.update(|d| {
        if let Some(e) = d.entries.iter_mut().find(|e| e.id == entry_id) {
            // 같은 규칙으로 다시 보낸 것이면 흔적도 하나만 남긴다.
            e.flowdeck_todos.retain(|t| t.todo_id != todo.todo_id);
            e.flowdeck_todos.push(todo.clone());
        }
    });
    let _ = app.emit("entries-changed", ());
    Ok(todo.todo_id)
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

// ---------- 휴지통 ----------

/// 관리함의 파일을 OS 휴지통으로 보낸다. 영구 삭제는 하지 않으며,
/// 휴지통이 없는 위치(네트워크 드라이브 등)에서는 아무것도 하지 않고 실패를 알린다.
#[tauri::command]
pub fn trash_entry(
    app: AppHandle,
    store: State<Store>,
    entry_id: String,
) -> Result<FileEntry, String> {
    let entry = store
        .read(|d| d.entries.iter().find(|e| e.id == entry_id).cloned())
        .ok_or("항목을 찾을 수 없습니다")?;
    if entry.status != EntryStatus::Inbox {
        return Err("관리함에 있는 파일만 휴지통으로 보낼 수 있습니다".into());
    }
    if !entry.path.is_file() {
        return Err("파일을 찾을 수 없습니다. 이미 옮겨졌거나 삭제된 것 같아요".into());
    }
    trash::delete(&entry.path)
        .map_err(|e| format!("휴지통으로 보내지 못했습니다: {e}"))?;

    let updated = store.update(|d| {
        d.entries.iter_mut().find(|e| e.id == entry_id).map(|e| {
            e.status = EntryStatus::Trashed;
            e.filed_to = None;
            e.filed_at = Some(now_millis());
            e.record_id = None;
            e.clone()
        })
    });
    let _ = app.emit("entries-changed", ());
    updated.ok_or_else(|| "항목 갱신 실패".into())
}

/// 여러 항목을 한 번에 휴지통으로 보낸다.
#[tauri::command]
pub fn trash_many(app: AppHandle, store: State<Store>, entry_ids: Vec<String>) -> BatchResult {
    let mut moved = 0;
    let mut errors = Vec::new();
    for id in entry_ids {
        let name = store.read(|d| {
            d.entries.iter().find(|e| e.id == id).map(|e| e.file_name.clone())
        });
        match trash_entry_inner(&app, &store, &id) {
            Ok(()) => moved += 1,
            Err(e) => errors.push(format!("{}: {e}", name.unwrap_or(id))),
        }
    }
    BatchResult { moved, errors }
}

fn trash_entry_inner(app: &AppHandle, store: &Store, entry_id: &str) -> Result<(), String> {
    let entry = store
        .read(|d| d.entries.iter().find(|e| e.id == entry_id).cloned())
        .ok_or("항목을 찾을 수 없습니다")?;
    if entry.status != EntryStatus::Inbox {
        return Err("관리함에 있는 파일만 휴지통으로 보낼 수 있습니다".into());
    }
    if !entry.path.is_file() {
        return Err("파일을 찾을 수 없습니다".into());
    }
    trash::delete(&entry.path).map_err(|e| format!("휴지통으로 보내지 못했습니다: {e}"))?;
    store.update(|d| {
        if let Some(e) = d.entries.iter_mut().find(|e| e.id == entry_id) {
            e.status = EntryStatus::Trashed;
            e.filed_to = None;
            e.filed_at = Some(now_millis());
            e.record_id = None;
        }
    });
    let _ = app.emit("entries-changed", ());
    Ok(())
}

/// 윈도우 휴지통에 들어 있는 항목 수 (FileBox가 보낸 것뿐 아니라 전체)
#[tauri::command]
pub fn trash_count() -> Result<usize, String> {
    trash::os_limited::list()
        .map(|items| items.len())
        .map_err(|e| format!("휴지통을 읽지 못했습니다: {e}"))
}

/// 윈도우 휴지통을 통째로 비운다. 되돌릴 수 없으므로 화면에서 반드시 확인을 받는다.
#[tauri::command]
pub fn empty_trash() -> Result<usize, String> {
    let items = trash::os_limited::list()
        .map_err(|e| format!("휴지통을 읽지 못했습니다: {e}"))?;
    let count = items.len();
    trash::os_limited::purge_all(items)
        .map_err(|e| format!("휴지통을 비우지 못했습니다: {e}"))?;
    Ok(count)
}

/// 윈도우 휴지통 폴더를 탐색기로 연다.
#[tauri::command]
pub fn open_trash() -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        std::process::Command::new("explorer.exe")
            .arg("shell:RecycleBinFolder")
            .spawn()
            .map(|_| ())
            .map_err(|e| format!("휴지통을 열지 못했습니다: {e}"))
    }
    #[cfg(not(target_os = "windows"))]
    {
        Err("이 플랫폼에서는 휴지통 열기를 지원하지 않습니다".into())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{FileEntry, FlowdeckSpec, Rule, Settings};

    struct Env {
        store: Store,
        root: PathBuf,
    }

    /// 관리함에 파일 하나가 수집돼 있고, 최종 폴더와 Flowdeck 감시 폴더가 준비된 상태.
    fn setup(rule: Option<Rule>, pending: bool) -> Env {
        let root = std::env::temp_dir().join(format!("fbcmd-{}", new_id()));
        let inbox = root.join("관리함");
        let dest = root.join("최종폴더");
        let fd = root.join("fd-inbox");
        for dir in [&inbox, &dest, &fd] {
            std::fs::create_dir_all(dir).unwrap();
        }

        let file = inbox.join("2026_상반기_시장분석.pdf");
        std::fs::write(&file, "리포트").unwrap();

        let store = Store::load(root.join("store.json"));
        store.update(|d| {
            let mut settings = Settings::new(root.join("dl"), inbox.clone());
            settings.flowdeck_enabled = true;
            settings.flowdeck_inbox = Some(fd.clone());
            d.settings = Some(settings);
            if let Some(r) = rule.clone() {
                d.rules.push(r);
            }
            d.entries.push(FileEntry {
                id: "e1".into(),
                file_name: "2026_상반기_시장분석.pdf".into(),
                path: file.clone(),
                origin: root.join("dl/2026_상반기_시장분석.pdf"),
                size: 6,
                added_at: now_millis(),
                category: "리포트".into(),
                tags: vec![],
                status: EntryStatus::Inbox,
                filed_to: None,
                filed_at: None,
                record_id: None,
                flowdeck_todos: vec![],
                flowdeck_pending: pending,
            });
        });

        Env { store, root }
    }

    fn transfers(env: &Env) -> Vec<String> {
        std::fs::read_dir(env.root.join("fd-inbox"))
            .map(|rd| {
                rd.filter_map(|e| e.ok())
                    .filter(|e| e.path().extension().is_some_and(|x| x == "txt"))
                    .filter_map(|e| std::fs::read_to_string(e.path()).ok())
                    .collect()
            })
            .unwrap_or_default()
    }

    fn reading_rule() -> Rule {
        Rule {
            id: "r1".into(),
            name: "리포트 읽기".into(),
            extensions: vec!["pdf".into()],
            keywords: vec![],
            category: None,
            favorite_id: None,
            flowdeck: Some(FlowdeckSpec {
                title: "[읽기] {파일명}".into(),
                due_in_days: Some(7),
                due_time: Some("10:00".into()),
                priority: "High".into(),
                tags: vec!["리포트".into()],
                reminder_minutes: Some(30),
            }),
        }
    }

    /// 이 릴리스의 요점. 이동 명령이 발송을 부르고, 메모에 최종 경로가 적혀야 한다.
    #[test]
    fn filing_sends_the_todo_with_the_destination_path() {
        let env = setup(Some(reading_rule()), false);
        assert!(transfers(&env).is_empty(), "이동 전에는 아무것도 보내지 않는다");

        let dest = env.root.join("최종폴더");
        let filed = file_entry_to(&env.store, "e1", dest.clone(), None)
            .expect("이동 실패");

        assert_eq!(filed.status, EntryStatus::Filed);
        assert_eq!(filed.flowdeck_todos.len(), 1, "이동했는데 할일이 안 만들어졌다");

        let files = transfers(&env);
        assert_eq!(files.len(), 1);
        let sent = &files[0];
        assert!(
            sent.contains(dest.join("2026_상반기_시장분석.pdf").to_str().unwrap()),
            "메모에 최종 경로가 없다: {sent}"
        );
        assert!(!sent.contains("관리함"), "관리함 경로가 새어 나갔다: {sent}");
        std::fs::remove_dir_all(&env.root).ok();
    }

    /// 규칙이 없어도 관리함에서 표시해 둔 파일은 이동할 때 나가야 한다.
    #[test]
    fn filing_a_marked_file_sends_it_even_without_a_rule() {
        let env = setup(None, true);
        let dest = env.root.join("최종폴더");
        let filed = file_entry_to(&env.store, "e1", dest, None).expect("이동 실패");

        assert_eq!(filed.flowdeck_todos.len(), 1);
        assert!(!filed.flowdeck_pending, "표시는 소비돼야 한다");
        assert_eq!(transfers(&env).len(), 1);
        std::fs::remove_dir_all(&env.root).ok();
    }

    /// 규칙도 없고 표시도 없으면 조용히 지나가야 한다.
    #[test]
    fn filing_an_ordinary_file_sends_nothing() {
        let env = setup(None, false);
        let dest = env.root.join("최종폴더");
        file_entry_to(&env.store, "e1", dest, None).expect("이동 실패");

        assert!(transfers(&env).is_empty(), "보낼 이유가 없는데 보냈다");
        std::fs::remove_dir_all(&env.root).ok();
    }

}
