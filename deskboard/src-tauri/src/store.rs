//! 보드 내용을 앱 데이터 폴더의 JSON 한 파일에 보관한다.

use crate::model::Board;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::{SystemTime, UNIX_EPOCH};

pub struct Store {
    path: PathBuf,
    board: Mutex<Board>,
}

impl Store {
    /// 파일이 없으면 시작용 보드로 연다. 읽을 수 없는 파일은 덮어쓰지 않고
    /// 옆에 따로 남겨 둔다 — 메모를 잃는 것보다 파일이 하나 더 생기는 편이 낫다.
    pub fn open(path: PathBuf) -> Self {
        let board = match fs::read_to_string(&path) {
            Ok(text) => match serde_json::from_str::<Board>(&text) {
                Ok(mut board) => {
                    board.sanitize();
                    board
                }
                Err(_) => {
                    set_aside(&path);
                    Board::starter()
                }
            },
            Err(_) => Board::starter(),
        };
        Store { path, board: Mutex::new(board) }
    }

    pub fn get(&self) -> Board {
        self.board.lock().unwrap().clone()
    }

    /// 디스크에 먼저 쓰고, 성공했을 때만 메모리 내용을 바꾼다.
    pub fn replace(&self, mut board: Board) -> Result<(), String> {
        board.sanitize();
        let json = serde_json::to_string_pretty(&board).map_err(|e| e.to_string())?;
        let mut guard = self.board.lock().unwrap();
        write_atomic(&self.path, json.as_bytes())
            .map_err(|e| format!("저장하지 못했습니다: {e}"))?;
        *guard = board;
        Ok(())
    }
}

/// 백업 파일의 내용을 보드로 읽는다. 아무 JSON 이나 빈 보드로 받아들이지 않도록 모양을 먼저 본다.
pub fn parse_backup(text: &str) -> Result<Board, String> {
    let not_a_backup = || "DeskBoard 백업 파일이 아닙니다.".to_string();
    let value: serde_json::Value = serde_json::from_str(text).map_err(|_| not_a_backup())?;
    if !value["categories"].is_array() || !value["notes"].is_array() {
        return Err(not_a_backup());
    }
    let mut board: Board = serde_json::from_value(value).map_err(|_| not_a_backup())?;
    board.sanitize();
    Ok(board)
}

/// 임시 파일에 쓴 뒤 교체해, 쓰는 도중 꺼져도 기존 파일이 깨지지 않게 한다.
pub fn write_atomic(path: &Path, bytes: &[u8]) -> io::Result<()> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    let tmp = path.with_extension("json.tmp");
    fs::write(&tmp, bytes)?;
    fs::rename(&tmp, path)
}

fn set_aside(path: &Path) {
    let stamp = SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_secs()).unwrap_or(0);
    let _ = fs::rename(path, path.with_extension(format!("corrupt-{stamp}.json")));
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{Category, Tile};

    fn temp_dir(name: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!("deskboard-test-{name}-{}", std::process::id()));
        let _ = fs::remove_dir_all(&dir);
        fs::create_dir_all(&dir).unwrap();
        dir
    }

    #[test]
    fn first_run_opens_the_starter_board() {
        let dir = temp_dir("first-run");
        let store = Store::open(dir.join("board.json"));
        assert_eq!(store.get(), Board::starter());
    }

    #[test]
    fn saved_board_comes_back_after_reopening() {
        let dir = temp_dir("roundtrip");
        let path = dir.join("nested").join("board.json");
        let mut board = Board::starter();
        board.categories.push(Category {
            id: "c".into(),
            name: "업무".into(),
            tiles: vec![Tile {
                id: "t".into(),
                name: "보고서".into(),
                target: r"C:\업무\보고서".into(),
                kind: "folder".into(),
                color: "teal".into(),
                icon: "emoji:📁".into(),
            }],
        });
        board.notes[0].text = "장보기\n우유".into();
        Store::open(path.clone()).replace(board.clone()).unwrap();

        assert_eq!(Store::open(path).get(), board);
        assert!(!dir.join("nested").join("board.json.tmp").exists());
    }

    #[test]
    fn unreadable_file_is_kept_aside_instead_of_overwritten() {
        let dir = temp_dir("corrupt");
        let path = dir.join("board.json");
        fs::write(&path, "{ not json").unwrap();

        let store = Store::open(path.clone());
        assert_eq!(store.get(), Board::starter());
        store.replace(Board::default()).unwrap();

        let kept: Vec<_> = fs::read_dir(&dir)
            .unwrap()
            .filter_map(|e| e.ok())
            .filter(|e| e.file_name().to_string_lossy().contains("corrupt-"))
            .collect();
        assert_eq!(kept.len(), 1);
        assert_eq!(fs::read_to_string(kept[0].path()).unwrap(), "{ not json");
    }

    #[test]
    fn exported_board_can_be_imported_back() {
        let mut board = Board::starter();
        board.settings.theme = "mint".into();
        let json = serde_json::to_string_pretty(&board).unwrap();
        assert_eq!(parse_backup(&json).unwrap(), board);
    }

    #[test]
    fn files_that_are_not_backups_are_refused() {
        for text in ["", "not json", "{}", "[]", r#"{"x":1,"y":2,"width":3,"height":4}"#, r#"{"categories":1,"notes":[]}"#] {
            assert!(parse_backup(text).is_err(), "{text}");
        }
    }

    #[test]
    fn replace_stores_the_sanitized_board() {
        let dir = temp_dir("sanitize");
        let store = Store::open(dir.join("board.json"));
        let mut board = Board::starter();
        board.settings.split = 9.0;
        store.replace(board).unwrap();
        assert_eq!(store.get().settings.split, 0.8);
    }
}
