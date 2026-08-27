use crate::model::StoreData;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;

pub struct Store {
    pub data: Mutex<StoreData>,
    path: PathBuf,
}

impl Store {
    pub fn load(path: PathBuf) -> Self {
        let data = fs::read_to_string(&path)
            .ok()
            .and_then(|s| serde_json::from_str::<StoreData>(&s).ok())
            .unwrap_or_default();
        Store { data: Mutex::new(data), path }
    }

    /// 잠금을 잡고 data를 수정한 뒤 디스크에 저장한다.
    pub fn update<T>(&self, f: impl FnOnce(&mut StoreData) -> T) -> T {
        let mut guard = self.data.lock().unwrap();
        let out = f(&mut guard);
        let json = serde_json::to_string_pretty(&*guard).unwrap_or_default();
        drop(guard);
        if let Some(parent) = self.path.parent() {
            let _ = fs::create_dir_all(parent);
        }
        // 임시 파일에 쓴 뒤 교체해 저장 도중 손상 방지
        let tmp = self.path.with_extension("json.tmp");
        if fs::write(&tmp, &json).is_ok() {
            let _ = fs::rename(&tmp, &self.path);
        }
        out
    }

    pub fn read<T>(&self, f: impl FnOnce(&StoreData) -> T) -> T {
        let guard = self.data.lock().unwrap();
        f(&guard)
    }
}

/// dest_dir 안에서 겹치지 않는 파일 경로를 만든다 (name.ext, name (1).ext, ...)
pub fn unique_dest(dest_dir: &Path, file_name: &str) -> PathBuf {
    let candidate = dest_dir.join(file_name);
    if !candidate.exists() {
        return candidate;
    }
    let stem = crate::model::stem_of(file_name);
    let ext = crate::model::ext_of(file_name);
    for i in 1..1000 {
        let name = if ext.is_empty() {
            format!("{stem} ({i})")
        } else {
            format!("{stem} ({i}).{ext}")
        };
        let candidate = dest_dir.join(name);
        if !candidate.exists() {
            return candidate;
        }
    }
    dest_dir.join(format!("{}-{}", crate::model::new_id(), file_name))
}

/// 드라이브 간 이동까지 지원하는 파일 이동 (rename 실패 시 copy+delete)
pub fn move_file(src: &Path, dest: &Path) -> std::io::Result<()> {
    if let Some(parent) = dest.parent() {
        fs::create_dir_all(parent)?;
    }
    match fs::rename(src, dest) {
        Ok(()) => Ok(()),
        Err(_) => {
            fs::copy(src, dest)?;
            fs::remove_file(src)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{now_millis, EntryStatus, FileEntry};

    fn temp_dir(name: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!("filebox-test-{name}-{}", now_millis()));
        fs::create_dir_all(&dir).unwrap();
        dir
    }

    #[test]
    fn store_roundtrip() {
        let dir = temp_dir("store");
        let path = dir.join("store.json");
        let store = Store::load(path.clone());
        store.update(|d| {
            d.entries.push(FileEntry {
                id: "e1".into(),
                file_name: "a.pdf".into(),
                path: "/inbox/a.pdf".into(),
                origin: "/dl/a.pdf".into(),
                size: 10,
                added_at: 1,
                category: "문서".into(),
                tags: vec![],
                status: EntryStatus::Inbox,
                filed_to: None,
                filed_at: None,
                record_id: None,
                flowdeck_todos: Vec::new(),
            });
        });
        let reloaded = Store::load(path);
        assert_eq!(reloaded.read(|d| d.entries.len()), 1);
        assert_eq!(reloaded.read(|d| d.entries[0].file_name.clone()), "a.pdf");
        let _ = fs::remove_dir_all(dir);
    }

    /// 구버전(record_id / MoveRecord.id 없음) 저장 파일도 손실 없이 읽혀야 한다.
    /// 파싱에 실패하면 기본값으로 덮어써져 사용자 데이터가 통째로 사라진다.
    #[test]
    fn loads_store_written_by_older_version() {
        let dir = temp_dir("compat");
        let path = dir.join("store.json");
        let legacy = r#"{
          "entries": [{
            "id": "e1", "file_name": "a.pdf", "path": "/inbox/a.pdf",
            "origin": "/dl/a.pdf", "size": 10, "added_at": 1,
            "category": "문서", "tags": ["급함"], "status": "filed",
            "filed_to": "/dest", "filed_at": 2
          }],
          "favorites": [{"id": "f1", "name": "문서함", "path": "/dest"}],
          "rules": [],
          "records": [{"ext": "pdf", "tokens": ["report"], "favorite_id": "f1", "at": 3}],
          "settings": null
        }"#;
        fs::write(&path, legacy).unwrap();

        let store = Store::load(path);
        assert_eq!(store.read(|d| d.entries.len()), 1, "구버전 항목이 유실됨");
        assert_eq!(store.read(|d| d.records.len()), 1, "구버전 학습 기록이 유실됨");
        assert_eq!(store.read(|d| d.favorites.len()), 1);
        assert_eq!(store.read(|d| d.entries[0].tags.clone()), vec!["급함"]);
        assert!(store.read(|d| d.entries[0].record_id.is_none()));
        let _ = fs::remove_dir_all(dir);
    }

    /// 구버전 설정에는 auto_update_check 필드가 없다. 이 필드 하나 때문에 파싱이
    /// 실패하면 감시 폴더·관리함 경로까지 통째로 초기화되므로 기본값으로 읽혀야 한다.
    #[test]
    fn loads_settings_without_auto_update_field() {
        let dir = temp_dir("compat-settings");
        let path = dir.join("store.json");
        let legacy = r#"{
          "entries": [], "favorites": [], "rules": [], "records": [],
          "settings": {
            "watch_dirs": ["C:/Downloads"],
            "inbox_dir": "C:/FileBox",
            "auto_collect": true,
            "toast_enabled": false,
            "paused": true
          }
        }"#;
        fs::write(&path, legacy).unwrap();

        let store = Store::load(path);
        let settings = store.read(|d| d.settings.clone()).expect("설정이 유실됨");
        assert_eq!(settings.watch_dirs.len(), 1, "감시 폴더가 유실됨");
        assert_eq!(settings.inbox_dir.to_string_lossy(), "C:/FileBox");
        assert!(!settings.toast_enabled, "기존 설정값이 덮어써짐");
        assert!(settings.paused, "기존 설정값이 덮어써짐");
        assert!(settings.auto_update_check, "새 옵션은 기본으로 켜져 있어야 한다");
        let _ = fs::remove_dir_all(dir);
    }

    #[test]
    fn loads_store_written_before_the_flowdeck_bridge() {
        // 0.2.3 이하가 쓴 파일. 파싱이 실패하면 사용자 데이터 전체가 조용히 사라진다.
        let dir = temp_dir("compat-flowdeck");
        let path = dir.join("store.json");
        let legacy = r#"{
          "entries": [{
            "id": "e1", "file_name": "a.pdf", "path": "/inbox/a.pdf", "origin": "/dl/a.pdf",
            "size": 10, "added_at": 1, "category": "문서", "tags": [], "status": "inbox",
            "filed_to": null, "filed_at": null
          }],
          "favorites": [],
          "rules": [{
            "id": "r1", "name": "청구서", "extensions": ["pdf"], "keywords": ["청구서"],
            "category": "경리", "favorite_id": null
          }],
          "records": [],
          "settings": {
            "watch_dirs": ["C:/Downloads"], "inbox_dir": "C:/FileBox",
            "auto_collect": true, "toast_enabled": true, "paused": false,
            "auto_update_check": false
          }
        }"#;
        fs::write(&path, legacy).unwrap();

        let store = Store::load(path);
        let (entries, rules, settings) =
            store.read(|d| (d.entries.clone(), d.rules.clone(), d.settings.clone()));

        assert_eq!(entries.len(), 1, "항목이 유실됨");
        assert!(entries[0].flowdeck_todos.is_empty());
        assert_eq!(rules.len(), 1, "규칙이 유실됨");
        assert_eq!(rules[0].category.as_deref(), Some("경리"), "기존 규칙이 뭉개짐");
        assert!(rules[0].flowdeck.is_none(), "옛 규칙은 특별규칙이 없어야 한다");

        let settings = settings.expect("설정이 유실됨");
        assert!(!settings.auto_update_check, "기존 설정값이 덮어써짐");
        assert!(settings.flowdeck_enabled, "연동은 기본으로 켜져 있어야 한다");
        assert!(settings.flowdeck_inbox.is_none(), "경로는 기본값을 쓴다");
        let _ = fs::remove_dir_all(dir);
    }

    #[test]
    fn unique_dest_appends_counter() {
        let dir = temp_dir("unique");
        fs::write(dir.join("a.txt"), "x").unwrap();
        let next = unique_dest(&dir, "a.txt");
        assert_eq!(next.file_name().unwrap().to_string_lossy(), "a (1).txt");
        let fresh = unique_dest(&dir, "b.txt");
        assert_eq!(fresh.file_name().unwrap().to_string_lossy(), "b.txt");
        let _ = fs::remove_dir_all(dir);
    }

    #[test]
    fn move_file_across_dirs() {
        let dir = temp_dir("move");
        let src = dir.join("src.txt");
        let dest = dir.join("sub").join("dest.txt");
        fs::write(&src, "hello").unwrap();
        move_file(&src, &dest).unwrap();
        assert!(!src.exists());
        assert_eq!(fs::read_to_string(&dest).unwrap(), "hello");
        let _ = fs::remove_dir_all(dir);
    }
}
