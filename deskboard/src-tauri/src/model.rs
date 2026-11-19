//! 위젯이 기억하는 내용 전체. 프런트엔드와 같은 모양(camelCase JSON)을 주고받는다.

use serde::{Deserialize, Serialize};

pub const NOTE_MIN_W: f64 = 190.0;
pub const NOTE_MIN_H: f64 = 90.0;
const NOTE_MAX: f64 = 4000.0;
const NOTE_COLORS: [&str; 5] = ["yellow", "pink", "green", "blue", "purple"];
const TILE_KINDS: [&str; 4] = ["folder", "program", "file", "link"];
const TILE_SIZES: [&str; 3] = ["s", "m", "l"];
/// 타일 바탕색. 빈 문자열은 테마 기본색이다.
const TILE_COLORS: [&str; 9] = ["", "red", "orange", "yellow", "green", "teal", "blue", "purple", "pink"];
const THEMES: [&str; 4] = ["dark", "light", "mint", "lavender"];
const NOTE_KINDS: [&str; 2] = ["text", "list"];
/// 메모 글자색. 빈 문자열은 기본색이다.
const INK_COLORS: [&str; 6] = ["", "red", "orange", "green", "blue", "purple"];
pub const NOTE_FONT_MIN: f64 = 10.0;
pub const NOTE_FONT_MAX: f64 = 32.0;
pub const DEFAULT_HOTKEY: &str = "Ctrl+Alt+D";

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
    /// folder | program | file | link — 아이콘을 구하지 못했을 때 보일 그림을 고르는 데 쓴다
    pub kind: String,
    /// 타일 바탕색 이름. 비어 있으면 테마 기본색
    pub color: String,
    /// 비어 있으면 대상의 아이콘을 그대로 쓴다. `emoji:🙂` 는 글자 하나, `file:경로` 는 그 파일의 그림
    pub icon: String,
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
    /// text: 자유롭게 쓰는 메모 / list: 체크리스트
    pub kind: String,
    /// 체크리스트 항목 (kind 가 list 일 때 쓴다)
    pub items: Vec<NoteItem>,
    pub font_size: f64,
    /// text 의 어느 부분이 굵은지·기울었는지·무슨 색인지
    pub marks: Vec<Mark>,
}

#[derive(Serialize, Deserialize, Clone, Debug, PartialEq, Default)]
#[serde(rename_all = "camelCase", default)]
pub struct NoteItem {
    pub id: String,
    pub text: String,
    pub done: bool,
    pub marks: Vec<Mark>,
}

/// 글의 한 구간에 입힌 서식. 위치는 화면(JavaScript)이 세는 방식대로 UTF-16 단위다.
#[derive(Serialize, Deserialize, Clone, Debug, PartialEq, Default)]
#[serde(rename_all = "camelCase", default)]
pub struct Mark {
    pub start: u32,
    pub end: u32,
    pub bold: bool,
    pub italic: bool,
    /// 글자색 이름. 비어 있으면 기본색
    pub color: String,
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
    /// dark | light | mint | lavender
    pub theme: String,
    /// 메모 글꼴 이름. 비어 있으면 기본 글꼴
    pub note_font: String,
    /// 위젯을 앞으로 불러오는 전역 단축키. 비어 있으면 쓰지 않는다
    pub hotkey: String,
    /// 시작할 때 새 버전이 있는지 확인한다
    pub check_updates: bool,
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
            kind: "text".into(),
            items: Vec::new(),
            font_size: 13.0,
            marks: Vec::new(),
        }
    }
}

impl Default for Settings {
    fn default() -> Self {
        Settings {
            split: 0.56,
            tile_size: "m".into(),
            opacity: 0.86,
            locked: true,
            theme: "dark".into(),
            note_font: String::new(),
            hotkey: DEFAULT_HOTKEY.into(),
            check_updates: true,
        }
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
                       · ＋ 바로가기, 또는 탐색기에서 파일·폴더를 끌어다 놓아 추가\n\
                       · 메모는 위쪽 띠를 끌어 옮기고, 오른쪽 아래 모서리로 크기 조절\n\
                       · 자물쇠 버튼을 풀면 위젯 위치·크기 변경\n\
                       · Ctrl+Alt+D 로 어디서든 위젯 불러오기"
                    .into(),
                w: 270.0,
                h: 190.0,
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
        if !THEMES.contains(&s.theme.as_str()) {
            s.theme = "dark".into();
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
            if !NOTE_KINDS.contains(&note.kind.as_str()) {
                note.kind = d.kind;
            }
            note.font_size = clamp_or(note.font_size, NOTE_FONT_MIN, NOTE_FONT_MAX, d.font_size);
            clean_marks(&note.text, &mut note.marks);
            for item in &mut note.items {
                clean_marks(&item.text, &mut item.marks);
            }
        }
        for tile in self.categories.iter_mut().flat_map(|c| c.tiles.iter_mut()) {
            if !TILE_KINDS.contains(&tile.kind.as_str()) {
                tile.kind = "file".into();
            }
            if !TILE_COLORS.contains(&tile.color.as_str()) {
                tile.color = String::new();
            }
        }
    }
}

/// 글 밖을 가리키거나 아무 서식도 없는 구간은 버린다.
fn clean_marks(text: &str, marks: &mut Vec<Mark>) {
    let length = text.encode_utf16().count() as u32;
    marks.retain_mut(|mark| {
        if !INK_COLORS.contains(&mark.color.as_str()) {
            mark.color = String::new();
        }
        mark.end = mark.end.min(length);
        mark.start < mark.end && (mark.bold || mark.italic || !mark.color.is_empty())
    });
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
        assert_eq!(board.notes[0].kind, "text");
        assert_eq!(board.notes[0].font_size, 13.0);
        assert_eq!(board.settings.hotkey, DEFAULT_HOTKEY);
        assert!(board.categories.is_empty());
    }

    #[test]
    fn json_uses_camel_case_names_the_frontend_expects() {
        let json = serde_json::to_value(Board::starter()).unwrap();
        assert!(json["settings"]["tileSize"].is_string());
        assert!(json["settings"]["noteFont"].is_string());
        assert!(json["settings"]["checkUpdates"].is_boolean());
        assert!(json["notes"][0]["fontSize"].is_number());
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
        board.notes[0].kind = "table".into();
        board.notes[0].font_size = 400.0;
        board.settings.theme = "neon".into();
        board.categories[0].tiles.push(Tile { kind: "weird".into(), color: "glitter".into(), ..Tile::default() });
        board.sanitize();
        assert_eq!(board.settings.split, 0.8);
        assert_eq!(board.settings.opacity, Settings::default().opacity);
        assert_eq!(board.settings.tile_size, "m");
        assert_eq!(board.notes[0].x, 0.0);
        assert_eq!(board.notes[0].w, NOTE_MIN_W);
        assert_eq!(board.notes[0].color, "yellow");
        assert_eq!(board.notes[0].kind, "text");
        assert_eq!(board.notes[0].font_size, NOTE_FONT_MAX);
        assert_eq!(board.settings.theme, "dark");
        assert_eq!(board.categories[0].tiles[0].kind, "file");
        assert_eq!(board.categories[0].tiles[0].color, "");
    }

    #[test]
    fn marks_outside_the_text_or_without_formatting_are_dropped() {
        let mut board = Board::starter();
        board.notes[0].text = "가나다 abc".into(); // 7 UTF-16 units
        board.notes[0].marks = vec![
            Mark { start: 0, end: 3, bold: true, ..Mark::default() },
            Mark { start: 4, end: 99, italic: true, color: "red".into(), ..Mark::default() },
            Mark { start: 9, end: 12, bold: true, ..Mark::default() },
            Mark { start: 1, end: 2, ..Mark::default() },
            Mark { start: 5, end: 6, color: "neon".into(), ..Mark::default() },
        ];
        board.sanitize();
        assert_eq!(
            board.notes[0].marks,
            vec![
                Mark { start: 0, end: 3, bold: true, ..Mark::default() },
                Mark { start: 4, end: 7, italic: true, color: "red".into(), ..Mark::default() },
            ]
        );
    }

    #[test]
    fn sanitize_leaves_a_valid_board_alone() {
        let mut board = Board::starter();
        board.sanitize();
        assert_eq!(board, Board::starter());
    }
}
