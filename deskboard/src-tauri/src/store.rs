//! 보드 내용을 앱 데이터 폴더의 JSON 한 파일에 보관한다.

use crate::model::Board;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::{SystemTime, UNIX_EPOCH};

/// 자동 백업을 며칠 치 남겨 둘지
const BACKUPS_KEPT: usize = 14;

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

    /// 자동 백업이 쌓이는 폴더
    pub fn backup_dir(&self) -> PathBuf {
        backup_dir(&self.path)
    }

    pub fn get(&self) -> Board {
        self.board.lock().unwrap().clone()
    }

    /// 디스크에 먼저 쓰고, 성공했을 때만 메모리 내용을 바꾼다.
    pub fn replace(&self, mut board: Board) -> Result<(), String> {
        board.sanitize();
        let json = serde_json::to_string_pretty(&board).map_err(|e| e.to_string())?;
        let mut guard = self.board.lock().unwrap();
        back_up_once_a_day(&self.path, &today());
        write_atomic(&self.path, json.as_bytes())
            .map_err(|e| format!("저장하지 못했습니다: {e}"))?;
        *guard = board;
        Ok(())
    }
}

fn backup_dir(board_file: &Path) -> PathBuf {
    board_file.parent().unwrap_or(Path::new(".")).join("backups")
}

/// 그날 처음 저장하기 직전의 보드 파일을 `backups/board-YYYYMMDD.json` 으로 남겨 둔다.
/// 실수로 지우거나 망가뜨린 내용을 "가져오기"로 되살릴 수 있게 하려는 것이다. 오래된 것은 지운다.
pub fn back_up_once_a_day(board_file: &Path, today: &str) {
    let dir = backup_dir(board_file);
    let target = dir.join(format!("board-{today}.json"));
    if target.exists() || !board_file.exists() {
        return;
    }
    if fs::create_dir_all(&dir).is_err() || fs::copy(board_file, &target).is_err() {
        return;
    }
    let mut backups: Vec<PathBuf> = fs::read_dir(&dir)
        .into_iter()
        .flatten()
        .filter_map(|entry| entry.ok().map(|e| e.path()))
        .filter(|path| {
            let name = path.file_name().unwrap_or_default().to_string_lossy().to_string();
            name.starts_with("board-") && name.ends_with(".json")
        })
        .collect();
    // 이름에 날짜가 들어 있어 이름순이 곧 날짜순이다.
    backups.sort();
    for old in backups.iter().rev().skip(BACKUPS_KEPT) {
        let _ = fs::remove_file(old);
    }
}

/// 오늘 날짜를 YYYYMMDD 로 (UTC 기준).
fn today() -> String {
    let seconds = SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_secs()).unwrap_or(0);
    date_of_day((seconds / 86_400) as i64)
}

/// 1970-01-01 부터 센 날수를 YYYYMMDD 로 바꾼다.
fn date_of_day(days_since_epoch: i64) -> String {
    // Howard Hinnant 의 civil_from_days
    let z = days_since_epoch + 719_468;
    let era = z.div_euclid(146_097);
    let day_of_era = z.rem_euclid(146_097);
    let year_of_era = (day_of_era - day_of_era / 1_460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_index = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_index + 2) / 5 + 1;
    let month = if month_index < 10 { month_index + 3 } else { month_index - 9 };
    let year = year_of_era + era * 400 + i64::from(month <= 2);
    format!("{year:04}{month:02}{day:02}")
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
        board.notes[0].boxes[0].text = "장보기\n우유".into();
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
    fn days_are_turned_into_calendar_dates() {
        assert_eq!(date_of_day(0), "19700101");
        assert_eq!(date_of_day(59), "19700301");
        assert_eq!(date_of_day(11_016), "20000229"); // 윤년
        assert_eq!(date_of_day(20_777), "20261120");
    }

    #[test]
    fn the_first_save_of_a_day_keeps_what_was_there_before() {
        let dir = temp_dir("backup");
        let path = dir.join("board.json");
        back_up_once_a_day(&path, "20260101");
        assert!(!dir.join("backups").exists(), "nothing to back up yet");

        fs::write(&path, "first").unwrap();
        back_up_once_a_day(&path, "20260101");
        fs::write(&path, "second").unwrap();
        back_up_once_a_day(&path, "20260101");
        fs::write(&path, "third").unwrap();
        back_up_once_a_day(&path, "20260102");

        let read = |name: &str| fs::read_to_string(dir.join("backups").join(name)).unwrap();
        assert_eq!(read("board-20260101.json"), "first");
        assert_eq!(read("board-20260102.json"), "third");
    }

    #[test]
    fn only_the_newest_backups_are_kept() {
        let dir = temp_dir("backup-prune");
        let path = dir.join("board.json");
        fs::write(&path, "x").unwrap();
        for day in 1..=20 {
            back_up_once_a_day(&path, &format!("202601{day:02}"));
        }
        let mut names: Vec<String> = fs::read_dir(dir.join("backups"))
            .unwrap()
            .map(|e| e.unwrap().file_name().to_string_lossy().to_string())
            .collect();
        names.sort();
        assert_eq!(names.len(), BACKUPS_KEPT);
        assert_eq!(names.first().unwrap(), "board-20260107.json");
        assert_eq!(names.last().unwrap(), "board-20260120.json");
    }

    #[test]
    fn saving_through_the_store_makes_a_backup() {
        let dir = temp_dir("backup-store");
        let path = dir.join("board.json");
        fs::write(&path, serde_json::to_string(&Board::starter()).unwrap()).unwrap();
        let store = Store::open(path);
        store.replace(Board::default()).unwrap();
        let backups: Vec<_> = fs::read_dir(store.backup_dir()).unwrap().collect();
        assert_eq!(backups.len(), 1);
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
