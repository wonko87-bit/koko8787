// src-tauri/src/model.rs 와 같은 모양이다.

export type TileKind = "folder" | "program" | "file" | "link";

/** 타일 바탕색. 빈 문자열은 테마 기본색이다. */
export const TILE_COLORS = ["", "red", "orange", "yellow", "green", "teal", "blue", "purple", "pink"] as const;
export type TileColor = (typeof TILE_COLORS)[number];

export interface Tile {
  id: string;
  name: string;
  /** 폴더·파일·프로그램 경로 또는 URL */
  target: string;
  /** 아이콘을 구하지 못했을 때 보일 그림을 고르는 데 쓴다 */
  kind: TileKind;
  color: TileColor;
  /** 비어 있으면 대상의 아이콘. `emoji:🙂` 는 글자 하나, `file:경로` 는 그 파일의 그림 */
  icon: string;
}

export interface Category {
  id: string;
  name: string;
  tiles: Tile[];
}

export const NOTE_COLORS = ["yellow", "pink", "green", "blue", "purple"] as const;
export type NoteColor = (typeof NOTE_COLORS)[number];

export const NOTE_MIN_W = 150;
export const NOTE_MIN_H = 90;
export const NOTE_FONT_MIN = 10;
export const NOTE_FONT_MAX = 32;

export interface NoteItem {
  id: string;
  text: string;
  done: boolean;
}

export interface Note {
  id: string;
  text: string;
  /** 메모판 왼쪽 위 기준 위치와 크기 (px) */
  x: number;
  y: number;
  w: number;
  h: number;
  color: NoteColor;
  /** 겹쳤을 때 위에 오는 순서 */
  z: number;
  /** text: 자유롭게 쓰는 메모 / list: 체크리스트 */
  kind: "text" | "list";
  items: NoteItem[];
  fontSize: number;
}

export type TileSize = "s" | "m" | "l";

export const THEMES = [
  { id: "dark", name: "다크" },
  { id: "light", name: "라이트" },
  { id: "mint", name: "파스텔 민트" },
  { id: "lavender", name: "파스텔 라벤더" },
] as const;
export type Theme = (typeof THEMES)[number]["id"];

export interface Settings {
  /** 바로가기 영역이 차지하는 너비 비율 (0.2 ~ 0.8) */
  split: number;
  tileSize: TileSize;
  /** 위젯 배경 불투명도 (0.2 ~ 1) */
  opacity: number;
  /** 잠그면 창 이동·크기 조절이 막힌다 */
  locked: boolean;
  theme: Theme;
  /** 메모 글꼴 이름. 비어 있으면 기본 글꼴 */
  noteFont: string;
  /** 위젯을 앞으로 불러오는 전역 단축키. 비어 있으면 쓰지 않는다 */
  hotkey: string;
  /** 시작할 때 새 버전이 있는지 확인한다 */
  checkUpdates: boolean;
}

export interface Board {
  version: number;
  categories: Category[];
  notes: Note[];
  settings: Settings;
}
