// src-tauri/src/model.rs 와 같은 모양이다.

export type TileKind = "folder" | "program" | "file" | "link";

export interface Tile {
  id: string;
  name: string;
  /** 폴더·파일·프로그램 경로 또는 URL */
  target: string;
  kind: TileKind;
}

export interface Category {
  id: string;
  name: string;
  tiles: Tile[];
}

export const NOTE_COLORS = ["yellow", "pink", "green", "blue", "purple"] as const;
export type NoteColor = (typeof NOTE_COLORS)[number];

export const NOTE_MIN_W = 120;
export const NOTE_MIN_H = 90;

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
}

export type TileSize = "s" | "m" | "l";

export interface Settings {
  /** 바로가기 영역이 차지하는 너비 비율 (0.2 ~ 0.8) */
  split: number;
  tileSize: TileSize;
  /** 위젯 배경 불투명도 (0.2 ~ 1) */
  opacity: number;
  /** 잠그면 창 이동·크기 조절이 막힌다 */
  locked: boolean;
}

export interface Board {
  version: number;
  categories: Category[];
  notes: Note[];
  settings: Settings;
}
