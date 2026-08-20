use crate::model::{new_id, now_millis, EntryStatus, FileEntry};
use crate::rules::{categorize, is_temp_download};
use crate::store::{move_file, unique_dest, Store};
use notify::{recommended_watcher, RecursiveMode, Watcher};
use std::collections::{HashMap, HashSet};
use std::path::{Path, PathBuf};
use std::sync::mpsc::{channel, RecvTimeoutError, Sender};
use std::sync::Mutex;
use std::time::{Duration, Instant};
use tauri::{AppHandle, Emitter, Manager};

/// 커맨드 쪽에서 감시 폴더 변경을 알리는 컨트롤 채널
pub struct WatcherCtl(pub Mutex<Sender<Msg>>);

pub enum Msg {
    Fs(notify::Result<notify::Event>),
    Rewatch,
}

/// 파일 크기가 이만큼 유지되면 "다운로드 완료"로 간주
const STABLE_AFTER: Duration = Duration::from_millis(2500);

pub fn spawn(app: AppHandle) -> Sender<Msg> {
    let (tx, rx) = channel::<Msg>();
    let fs_tx = tx.clone();
    let thread_tx = tx.clone();

    std::thread::spawn(move || {
        let mut watcher = match recommended_watcher(move |res| {
            let _ = fs_tx.send(Msg::Fs(res));
        }) {
            Ok(w) => w,
            Err(e) => {
                eprintln!("[filebox] watcher init failed: {e}");
                return;
            }
        };

        let mut watched: Vec<PathBuf> = Vec::new();
        // 감시 시작 시점에 이미 존재하던 파일 (수정 이벤트가 와도 수집하지 않음)
        let mut known: HashSet<PathBuf> = HashSet::new();
        rewatch(&app, &mut watcher, &mut watched, &mut known);

        // 안정화 대기 중인 후보: path -> (마지막 크기, 마지막 크기변화 시각)
        let mut pending: HashMap<PathBuf, (u64, Instant)> = HashMap::new();

        loop {
            match rx.recv_timeout(Duration::from_millis(800)) {
                Ok(Msg::Rewatch) => rewatch(&app, &mut watcher, &mut watched, &mut known),
                Ok(Msg::Fs(Ok(event))) => {
                    for path in event.paths {
                        consider(&app, &path, &known, &mut pending);
                    }
                }
                Ok(Msg::Fs(Err(e))) => eprintln!("[filebox] watch error: {e}"),
                Err(RecvTimeoutError::Timeout) => {}
                Err(RecvTimeoutError::Disconnected) => break,
            }

            // 안정화 검사
            let now = Instant::now();
            let mut ready: Vec<PathBuf> = Vec::new();
            pending.retain(|path, (size, changed_at)| {
                match std::fs::metadata(path) {
                    Err(_) => false, // 사라짐 (임시파일 rename 등)
                    Ok(meta) => {
                        if meta.len() != *size {
                            *size = meta.len();
                            *changed_at = now;
                            true
                        } else if now.duration_since(*changed_at) >= STABLE_AFTER {
                            ready.push(path.clone());
                            false
                        } else {
                            true
                        }
                    }
                }
            });

            for path in ready {
                known.insert(path.clone());
                if let Some(entry) = collect_path(&app, &path) {
                    notify_new_entry(&app, &entry);
                }
            }
        }
    });

    let _ = thread_tx; // sender kept alive by returned clone
    tx
}

fn rewatch(
    app: &AppHandle,
    watcher: &mut notify::RecommendedWatcher,
    watched: &mut Vec<PathBuf>,
    known: &mut HashSet<PathBuf>,
) {
    for dir in watched.drain(..) {
        let _ = watcher.unwatch(&dir);
    }
    known.clear();

    let store = app.state::<Store>();
    let dirs = store.read(|d| {
        d.settings
            .as_ref()
            .map(|s| s.watch_dirs.clone())
            .unwrap_or_default()
    });

    for dir in dirs {
        if !dir.is_dir() {
            continue;
        }
        // 기존 파일은 수집 대상에서 제외
        if let Ok(read) = std::fs::read_dir(&dir) {
            for e in read.flatten() {
                known.insert(e.path());
            }
        }
        if watcher.watch(&dir, RecursiveMode::NonRecursive).is_ok() {
            watched.push(dir);
        }
    }
}

fn consider(
    app: &AppHandle,
    path: &Path,
    known: &HashSet<PathBuf>,
    pending: &mut HashMap<PathBuf, (u64, Instant)>,
) {
    if known.contains(path) || pending.contains_key(path) {
        return;
    }
    let Ok(meta) = std::fs::metadata(path) else { return };
    if !meta.is_file() {
        return;
    }
    let Some(name) = path.file_name().map(|n| n.to_string_lossy().to_string()) else {
        return;
    };
    if is_temp_download(&name) {
        return;
    }

    let store = app.state::<Store>();
    let (active, inbox) = store.read(|d| {
        let s = d.settings.as_ref();
        (
            s.map(|s| s.auto_collect && !s.paused).unwrap_or(false),
            s.map(|s| s.inbox_dir.clone()).unwrap_or_default(),
        )
    });
    if !active || path.starts_with(&inbox) {
        return;
    }

    pending.insert(path.to_path_buf(), (meta.len(), Instant::now()));
}

/// 파일 하나를 관리함으로 이동하고 항목으로 등록한다. (감시/수동 수집 공용)
pub fn collect_path(app: &AppHandle, path: &Path) -> Option<FileEntry> {
    let store = app.state::<Store>();
    let (inbox, rules) = store.read(|d| {
        (
            d.settings.as_ref().map(|s| s.inbox_dir.clone()),
            d.rules.clone(),
        )
    });
    let inbox = inbox?;

    if !path.is_file() {
        return None; // 폴더나 사라진 경로는 수집하지 않는다
    }
    if path.starts_with(&inbox) {
        return None; // 이미 관리함 안에 있는 파일
    }
    let name = path.file_name()?.to_string_lossy().to_string();
    std::fs::create_dir_all(&inbox).ok()?;
    let dest = unique_dest(&inbox, &name);
    move_file(path, &dest).ok()?;
    let size = std::fs::metadata(&dest).map(|m| m.len()).unwrap_or(0);
    let final_name = dest.file_name()?.to_string_lossy().to_string();

    let entry = FileEntry {
        id: new_id(),
        file_name: final_name.clone(),
        path: dest,
        origin: path.to_path_buf(),
        size,
        added_at: now_millis(),
        category: categorize(&rules, &final_name),
        tags: Vec::new(),
        status: EntryStatus::Inbox,
        filed_to: None,
        filed_at: None,
        record_id: None,
    };
    store.update(|d| d.entries.push(entry.clone()));
    let _ = app.emit("entries-changed", ());
    Some(entry)
}

/// 새 항목에 대해 토스트 또는 OS 알림을 띄운다.
fn notify_new_entry(app: &AppHandle, entry: &FileEntry) {
    let store = app.state::<Store>();
    let toast_enabled = store.read(|d| {
        d.settings.as_ref().map(|s| s.toast_enabled).unwrap_or(true)
    });
    let suggestions = store.read(|d| {
        crate::suggest::suggest(&d.rules, &d.records, &d.favorites, &entry.file_name, 3)
    });

    if toast_enabled {
        crate::toast::show_toast(app, entry, &suggestions);
    } else {
        use tauri_plugin_notification::NotificationExt;
        let _ = app
            .notification()
            .builder()
            .title("FileBox: 새 파일 수집됨")
            .body(&entry.file_name)
            .show();
    }
}
