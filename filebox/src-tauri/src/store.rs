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
