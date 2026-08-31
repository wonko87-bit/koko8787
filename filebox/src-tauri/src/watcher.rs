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

/// 옮기지 못한 파일을 다시 시도하는 최대 횟수.
///
/// 간격이 두 배씩 늘어나 전부 합치면 20분쯤 된다. 저장해 둔 문서를 닫기까지
/// 그 정도면 대개 충분하고, 그 뒤에도 파일은 감시 폴더에 그대로 있으니
/// '기존 파일 수집'으로 언제든 가져올 수 있다.
const MAX_RETRIES: u32 = 9;

/// 첫 재시도까지 기다리는 시간. 시도할 때마다 두 배가 된다.
const FIRST_RETRY: Duration = Duration::from_secs(2);

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
        // 옮기지 못해 다시 시도할 것들: path -> (시도 횟수, 다음 시도 시각)
        let mut retry: HashMap<PathBuf, (u32, Instant)> = HashMap::new();

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
                match collect_path(&app, &path) {
                    Collected::Ok(entry) => {
                        known.insert(path.clone());
                        retry.remove(&path);
                        notify_new_entry(&app, &entry);
                    }
                    // 대상이 아니므로 다시 볼 필요도 없다.
                    Collected::Skipped => {
                        known.insert(path.clone());
                        retry.remove(&path);
                    }
                    // 지금 못 옮겼을 뿐이다. known 에 넣어 버리면 파일을 붙잡고 있던
                    // 프로그램이 놓아준 뒤에도 영영 들어오지 않는다.
                    Collected::Failed(reason) => {
                        let (attempts, _) = retry.get(&path).copied().unwrap_or((0, now));
                        let attempts = attempts + 1;
                        if attempts > MAX_RETRIES {
                            eprintln!("[filebox] 수집 포기 ({attempts}회 시도): {reason}");
                            known.insert(path.clone());
                            retry.remove(&path);
                        } else {
                            let wait = FIRST_RETRY * 2u32.saturating_pow(attempts - 1);
                            eprintln!("[filebox] 수집 실패, {wait:?} 뒤 다시 시도: {reason}");
                            retry.insert(path.clone(), (attempts, now + wait));
                        }
                    }
                }
            }

            // 때가 된 재시도는 다시 안정화 줄에 세운다.
            retry.retain(|path, (_, next_try)| {
                if *next_try > now {
                    return true;
                }
                match std::fs::metadata(path) {
                    // 그 사이 사라졌다면 더 볼 것이 없다.
                    Err(_) => false,
                    Ok(meta) => {
                        pending.insert(path.clone(), (meta.len(), now - STABLE_AFTER));
                        true
                    }
                }
            });
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
    if !active || path.starts_with(&inbox) || in_flowdeck_inbox(&store, path) {
        return;
    }

    pending.insert(path.to_path_buf(), (meta.len(), Instant::now()));
}

/// Flowdeck 에 넘기려고 쓴 파일인지. 그 폴더가 감시 폴더 아래에 있을 수 있다.
fn in_flowdeck_inbox(store: &Store, path: &Path) -> bool {
    crate::flowdeck::inbox_of(store).is_some_and(|dir| path.starts_with(&dir))
}

/// 수집을 시도한 결과.
pub enum Collected {
    /// 관리함으로 들어왔다.
    Ok(FileEntry),
    /// 애초에 수집 대상이 아니다 (폴더, 이미 관리함 안, 사라진 경로).
    Skipped,
    /// 대상은 맞는데 지금은 옮기지 못했다. 다시 시도할 가치가 있다.
    ///
    /// 파워포인트처럼 저장한 파일을 계속 붙잡고 있는 프로그램이 대표적이다.
    /// 윈도우에서는 읽기는 되면서 이동과 삭제만 막히므로, 여기서 포기해 버리면
    /// 그 프로그램을 닫아도 파일이 영영 들어오지 않는다.
    Failed(String),
}

/// 파일 하나를 관리함으로 이동하고 항목으로 등록한다. (감시/수동 수집 공용)
pub fn collect_path(app: &AppHandle, path: &Path) -> Collected {
    let store = app.state::<Store>();
    let (inbox, rules) = store.read(|d| {
        (
            d.settings.as_ref().map(|s| s.inbox_dir.clone()),
            d.rules.clone(),
        )
    });
    let Some(inbox) = inbox else {
        return Collected::Skipped;
    };

    if !path.is_file() {
        return Collected::Skipped; // 폴더나 사라진 경로는 수집하지 않는다
    }
    if path.starts_with(&inbox) {
        return Collected::Skipped; // 이미 관리함 안에 있는 파일
    }
    if in_flowdeck_inbox(&store, path) {
        // FileBox 가 Flowdeck 에 넘기려고 방금 쓴 파일. 감시 폴더 안에 Flowdeck 감시
        // 폴더가 있으면 자기가 쓴 것을 자기가 도로 주워 오게 된다.
        return Collected::Skipped;
    }
    let Some(name) = path.file_name().map(|n| n.to_string_lossy().to_string()) else {
        return Collected::Skipped;
    };
    if let Err(e) = std::fs::create_dir_all(&inbox) {
        return Collected::Failed(format!("관리함 폴더를 만들지 못했습니다: {e}"));
    }
    let dest = unique_dest(&inbox, &name);
    if let Err(e) = move_file(path, &dest) {
        return Collected::Failed(format!("{name}: {e}"));
    }
    let size = std::fs::metadata(&dest).map(|m| m.len()).unwrap_or(0);
    let final_name = dest.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or(name);

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
        flowdeck_todos: Vec::new(),
        flowdeck_pending: false,
        recent_cleared: false,
        pinned: false,
    };
    store.update(|d| d.entries.push(entry.clone()));

    // Flowdeck 으로는 여기서 보내지 않는다. 지금 이 파일의 경로는 관리함이고,
    // 관리함은 거쳐 가는 곳이라 최종 폴더로 옮기는 순간 그 경로가 죽는다.
    // 할일은 파일이 갈 곳이 정해진 뒤에 만든다 (commands::file_entry_to).

    let _ = app.emit("entries-changed", ());
    Collected::Ok(entry)
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
