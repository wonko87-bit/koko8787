//! 위젯이 기억하는 내용 전체. 프런트엔드와 같은 모양(camelCase JSON)을 주고받는다.

use serde::{Deserialize, Serialize};

pub const NOTE_MIN_W: f64 = 120.0;
pub const NOTE_MIN_H: f64 = 90.0;
const NOTE_MAX: f64 = 4000.0;
const NOTE_COLORS: [&str; 5] = ["yellow", "pink", "green", "blue", "purple"];
const TILE_KINDS: [&str; 4] = ["folder", "program", "file", "link"];
const TILE_SIZES: [&str; 3] = ["s", "m", "l"];

#[derive(Serialize, Deserialize, Clone, Debug, PartialEq)]
#[serde(rename_all = "camelCase", default)]
pub struct Board {
    pub version: u32,
    pub categories: Vec<Category>,
    pub notes: Vec<Note>,
    pub settings: Settings,
}

#[derive(Serialize, Deserialize, Clone, Debug, PartialEq, Default)]
#[serde(rename_all = "camelCase", default)]
pub struct Category {
    pub id: String,
    pub name: String,
    pub tiles: Vec<Tile>,
}

#[derive(Serialize, Deserialize, Clone, Debug, PartialEq, Default)]
#[serde(rename_all = "camelCase", default)]
pub struct Tile {
    pub id: String,
    pub name: String,
    /// 폴더·파일·프로그램 경로 또는 URL
    pub target: String,
    /// folder | program | file | link — 타일 그림을 고르는 데만 쓴다
    pub kind: String,
}

#[derive(Serialize, Deserialize, Clone, Debug, PartialEq)]
#[serde(rename_all = "camelCase", default)]
pub struct Note {
    pub id: String,
    pub text: String,
    /// 메모판 왼쪽 위를 기준으로 한 위치와 크기 (CSS px)
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
    pub color: String,
    /// 겹쳤을 때 위에 오는 순서
    pub z: u32,
}

#[derive(Serialize, Deserialize, Clone, Debug, PartialEq)]
#[serde(rename_all = "camelCase", default)]
pub struct Settings {
    /// 바로가기 영역이 차지하는 너비 비율
    pub split: f64,
    /// s | m | l
    pub tile_size: String,
    /// 위젯 배경의 불투명도
    pub opacity: f64,
    /// 잠그면 창 이동·크기 조절이 막힌다
    pub locked: bool,
}

impl Default for Board {
    fn default() -> Self {
        Board { version: 1, categories: Vec::new(), notes: Vec::new(), settings: Settings::default() }
    }
}

impl Default for Note {
    fn default() -> Self {
        Note {
            id: String::new(),
            text: String::new(),
            x: 16.0,
            y: 16.0,
            w: 190.0,
            h: 150.0,
            color: "yellow".into(),
            z: 1,
        }
    }
}

impl Default for Settings {
    fn default() -> Self {
        Settings { split: 0.56, tile_size: "m".into(), opacity: 0.86, locked: true }
    }
}

impl Board {
    /// 처음 실행했을 때 보여 줄 내용. 빈 화면 대신 쓰는 법을 메모 한 장에 적어 둔다.
    pub fn starter() -> Self {
        Board {
            categories: vec![
                Category { id: "cat-folders".into(), name: "자주 가는 폴더".into(), tiles: Vec::new() },
                Category { id: "cat-programs".into(), name: "프로그램".into(), tiles: Vec::new() },
            ],
            notes: vec![Note {
                id: "note-welcome".into(),
                text: "DeskBoard 사용법\n\n\
                       · 왼쪽 ＋ 로 폴더·프로그램 바로가기 추가\n\
                       · 메모는 위쪽 띠를 끌어 옮기고, 오른쪽 아래 모서리로 크기 조절\n\
                       · 자물쇠 버튼을 풀면 위젯 위치·크기 변경"
                    .into(),
                w: 250.0,
                h: 170.0,
                ..Note::default()
            }],
            ..Board::default()
        }
    }

    /// 손으로 고친 파일이나 잘못된 값이 들어와도 화면이 깨지지 않게 범위를 맞춘다.
    pub fn sanitize(&mut self) {
        self.version = 1;
        let s = &mut self.settings;
        s.split = clamp_or(s.split, 0.2, 0.8, Settings::default().split);
        s.opacity = clamp_or(s.opacity, 0.2, 1.0, Settings::default().opacity);
        if !TILE_SIZES.contains(&s.tile_size.as_str()) {
            s.tile_size = "m".into();
        }
        for note in &mut self.notes {
            let d = Note::default();
            note.x = clamp_or(note.x, 0.0, NOTE_MAX, d.x);
            note.y = clamp_or(note.y, 0.0, NOTE_MAX, d.y);
            note.w = clamp_or(note.w, NOTE_MIN_W, NOTE_MAX, d.w);
            note.h = clamp_or(note.h, NOTE_MIN_H, NOTE_MAX, d.h);
            if !NOTE_COLORS.contains(&note.color.as_str()) {
                note.color = d.color;
            }
        }
        for tile in self.categories.iter_mut().flat_map(|c| c.tiles.iter_mut()) {
            if !TILE_KINDS.contains(&tile.kind.as_str()) {
                tile.kind = "file".into();
            }
        }
    }
}

fn clamp_or(value: f64, min: f64, max: f64, fallback: f64) -> f64 {
    if value.is_finite() {
        value.clamp(min, max)
    } else {
        fallback
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn missing_fields_fall_back_to_defaults() {
        let board: Board = serde_json::from_str(r#"{"notes":[{"id":"a","text":"hi"}]}"#).unwrap();
        assert_eq!(board.settings, Settings::default());
        assert_eq!(board.notes[0].text, "hi");
        assert_eq!(board.notes[0].w, Note::default().w);
        assert!(board.categories.is_empty());
    }

    #[test]
    fn json_uses_camel_case_names_the_frontend_expects() {
        let json = serde_json::to_value(Board::starter()).unwrap();
        assert!(json["settings"]["tileSize"].is_string());
        assert!(json["categories"][0]["tiles"].is_array());
    }

    #[test]
    fn sanitize_pulls_out_of_range_values_back_in() {
        let mut board = Board::starter();
        board.settings.split = 5.0;
        board.settings.opacity = f64::NAN;
        board.settings.tile_size = "huge".into();
        board.notes[0].x = -40.0;
        board.notes[0].w = 3.0;
        board.notes[0].color = "black".into();
        board.categories[0].tiles.push(Tile { kind: "weird".into(), ..Tile::default() });
        board.sanitize();
        assert_eq!(board.settings.split, 0.8);
        assert_eq!(board.settings.opacity, Settings::default().opacity);
        assert_eq!(board.settings.tile_size, "m");
        assert_eq!(board.notes[0].x, 0.0);
        assert_eq!(board.notes[0].w, NOTE_MIN_W);
        assert_eq!(board.notes[0].color, "yellow");
        assert_eq!(board.categories[0].tiles[0].kind, "file");
    }

    #[test]
    fn sanitize_leaves_a_valid_board_alone() {
        let mut board = Board::starter();
        board.sanitize();
        assert_eq!(board, Board::starter());
    }
}
