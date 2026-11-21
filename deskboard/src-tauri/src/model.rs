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
    /// 메모판의 페이지(탭)들. 언제나 하나 이상 있다.
    pub pages: Vec<Page>,
    pub notes: Vec<Note>,
    pub settings: Settings,
}

/// 메모판의 한 페이지. 메모는 저마다 한 페이지에 붙어 있다.
#[derive(Serialize, Deserialize, Clone, Debug, PartialEq, Default)]
#[serde(rename_all = "camelCase", default)]
pub struct Page {
    pub id: String,
    pub name: String,
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
    /// 글 메모의 내용. 메모지 안 아무 데나 놓을 수 있는 글상자들이다.
    /// (예전 판은 글 전체를 text 에 두었다. 읽을 때 글상자 하나로 옮긴다.)
    pub boxes: Vec<TextBox>,
    /// 이 메모가 붙어 있는 페이지
    pub page: String,
}

/// 메모지 안의 글상자. 위치는 메모지 글 영역의 왼쪽 위 기준이다 (CSS px).
#[derive(Serialize, Deserialize, Clone, Debug, PartialEq, Default)]
#[serde(rename_all = "camelCase", default)]
pub struct TextBox {
    pub id: String,
    pub x: f64,
    pub y: f64,
    /// 사용자가 정한 폭. 0 이면 내용에 맞춘다.
    pub w: f64,
    pub text: String,
    pub marks: Vec<Mark>,
}

/// 폭을 정한 글상자의 가장 좁은 폭
pub const BOX_MIN_W: f64 = 40.0;

/// 글상자를 따로 놓지 않았을 때의 자리 (예전 글 메모를 옮길 때도 여기에 둔다)
pub const BOX_ORIGIN: (f64, f64) = (8.0, 6.0);

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
    /// 지금 보고 있는 메모판 페이지
    pub page: String,
}

impl Default for Board {
    fn default() -> Self {
        Board {
            version: 1,
            categories: Vec::new(),
            pages: Vec::new(),
            notes: Vec::new(),
            settings: Settings::default(),
        }
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
            boxes: Vec::new(),
            page: String::new(),
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
            page: String::new(),
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
        .sanitized()
    }

    fn sanitized(mut self) -> Self {
        self.sanitize();
        self
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
        // 페이지는 언제나 하나 이상 있고, 모든 메모와 "지금 페이지"는 있는 페이지를 가리킨다.
        self.pages.retain(|page| !page.id.is_empty());
        if self.pages.is_empty() {
            self.pages.push(Page { id: "page-1".into(), name: "메모".into() });
        }
        let first_page = self.pages[0].id.clone();
        let known = |id: &str| self.pages.iter().any(|page| page.id == id);
        if !known(&self.settings.page) {
            self.settings.page = first_page.clone();
        }
        for note in &mut self.notes {
            if !self.pages.iter().any(|page| page.id == note.page) {
                note.page = first_page.clone();
            }
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
            // 예전 판의 글 메모는 글 전체가 text 에 있다. 글상자 하나로 옮긴다.
            if note.kind == "text" && !note.text.trim().is_empty() {
                note.boxes.insert(
                    0,
                    TextBox {
                        id: format!("{}-text", note.id),
                        x: BOX_ORIGIN.0,
                        y: BOX_ORIGIN.1,
                        w: 0.0,
                        text: std::mem::take(&mut note.text),
                        marks: std::mem::take(&mut note.marks),
                    },
                );
            }
            // 아무것도 적지 않은 글상자는 남기지 않는다.
            note.boxes.retain(|b| !b.text.trim().is_empty());
            for b in &mut note.boxes {
                b.x = clamp_or(b.x, 0.0, NOTE_MAX, BOX_ORIGIN.0);
                b.y = clamp_or(b.y, 0.0, NOTE_MAX, BOX_ORIGIN.1);
                b.w = if b.w.is_finite() && b.w > 0.0 { b.w.clamp(BOX_MIN_W, NOTE_MAX) } else { 0.0 };
                clean_marks(&b.text, &mut b.marks);
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
        let text_box = &mut board.notes[0].boxes[0];
        text_box.text = "가나다 abc".into(); // 7 UTF-16 units
        text_box.marks = vec![
            Mark { start: 0, end: 3, bold: true, ..Mark::default() },
            Mark { start: 4, end: 99, italic: true, color: "red".into(), ..Mark::default() },
            Mark { start: 9, end: 12, bold: true, ..Mark::default() },
            Mark { start: 1, end: 2, ..Mark::default() },
            Mark { start: 5, end: 6, color: "neon".into(), ..Mark::default() },
        ];
        board.sanitize();
        assert_eq!(
            board.notes[0].boxes[0].marks,
            vec![
                Mark { start: 0, end: 3, bold: true, ..Mark::default() },
                Mark { start: 4, end: 7, italic: true, color: "red".into(), ..Mark::default() },
            ]
        );
    }

    #[test]
    fn old_text_notes_become_a_single_text_box() {
        let mut board: Board = serde_json::from_str(
            r#"{"notes":[{"id":"n","text":"장보기\n우유","marks":[{"start":0,"end":3,"bold":true}]}]}"#,
        )
        .unwrap();
        board.sanitize();
        let note = &board.notes[0];
        assert_eq!(note.text, "");
        assert!(note.marks.is_empty());
        assert_eq!(note.boxes.len(), 1);
        assert_eq!(note.boxes[0].text, "장보기\n우유");
        assert_eq!((note.boxes[0].x, note.boxes[0].y), BOX_ORIGIN);
        assert_eq!(note.boxes[0].marks, vec![Mark { start: 0, end: 3, bold: true, ..Mark::default() }]);
    }

    #[test]
    fn empty_text_boxes_are_not_kept_and_positions_stay_inside() {
        let mut board = Board::starter();
        board.notes[0].boxes = vec![
            TextBox { id: "a".into(), x: 40.0, y: 30.0, text: "남는다".into(), ..TextBox::default() },
            TextBox { id: "b".into(), x: 10.0, y: 80.0, text: "  \n ".into(), ..TextBox::default() },
            TextBox { id: "c".into(), x: -5.0, y: f64::NAN, text: "밖".into(), ..TextBox::default() },
        ];
        board.sanitize();
        let boxes = &board.notes[0].boxes;
        assert_eq!(boxes.iter().map(|b| b.id.as_str()).collect::<Vec<_>>(), ["a", "c"]);
        assert_eq!((boxes[1].x, boxes[1].y), (0.0, BOX_ORIGIN.1));
    }

    #[test]
    fn boards_from_before_pages_get_one_page_holding_every_note() {
        let mut board: Board =
            serde_json::from_str(r#"{"notes":[{"id":"a"},{"id":"b","page":"gone"}],"settings":{"page":"gone"}}"#).unwrap();
        board.sanitize();
        assert_eq!(board.pages, vec![Page { id: "page-1".into(), name: "메모".into() }]);
        assert!(board.notes.iter().all(|note| note.page == "page-1"));
        assert_eq!(board.settings.page, "page-1");
    }

    #[test]
    fn notes_keep_their_page_when_it_exists() {
        let mut board = Board::starter();
        board.pages.push(Page { id: "work".into(), name: "업무".into() });
        board.notes[0].page = "work".into();
        board.settings.page = "work".into();
        board.sanitize();
        assert_eq!(board.notes[0].page, "work");
        assert_eq!(board.settings.page, "work");
        assert_eq!(board.pages.len(), 2);
    }

    #[test]
    fn text_box_width_is_automatic_or_at_least_the_minimum() {
        let mut board = Board::starter();
        let template = board.notes[0].boxes[0].clone();
        board.notes[0].boxes = [0.0, -3.0, 5.0, 120.0, f64::NAN]
            .iter()
            .map(|w| TextBox { w: *w, ..template.clone() })
            .collect();
        board.sanitize();
        let widths: Vec<f64> = board.notes[0].boxes.iter().map(|b| b.w).collect();
        assert_eq!(widths, [0.0, 0.0, BOX_MIN_W, 120.0, 0.0]);
    }

    #[test]
    fn sanitize_leaves_a_valid_board_alone() {
        let mut board = Board::starter();
        board.sanitize();
        assert_eq!(board, Board::starter());
    }
}
