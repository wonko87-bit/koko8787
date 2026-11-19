// 보드를 바꾸는 순수 함수들. 원본은 건드리지 않고 새 보드를 돌려준다.

import { NOTE_FONT_MAX, NOTE_FONT_MIN, NOTE_MIN_H, NOTE_MIN_W } from "./types";
import { joinRich, sliceRich } from "./richText";
import type { Board, Category, Note, NoteItem, Rich, Settings, Tile } from "./types";

export const newId = () => crypto.randomUUID();

const mapCategory = (board: Board, id: string, fn: (c: Category) => Category): Board => ({
  ...board,
  categories: board.categories.map((c) => (c.id === id ? fn(c) : c)),
});

export function addCategory(board: Board, name: string): Board {
  return { ...board, categories: [...board.categories, { id: newId(), name, tiles: [] }] };
}

export function renameCategory(board: Board, id: string, name: string): Board {
  return mapCategory(board, id, (c) => ({ ...c, name }));
}

export function removeCategory(board: Board, id: string): Board {
  return { ...board, categories: board.categories.filter((c) => c.id !== id) };
}

/** 카테고리를 한 칸 위(-1)나 아래(+1)로 옮긴다. */
export function shiftCategory(board: Board, id: string, delta: -1 | 1): Board {
  const from = board.categories.findIndex((c) => c.id === id);
  const to = from + delta;
  if (from < 0 || to < 0 || to >= board.categories.length) return board;
  const categories = [...board.categories];
  [categories[from], categories[to]] = [categories[to], categories[from]];
  return { ...board, categories };
}

export function findTile(board: Board, tileId: string): { tile: Tile; categoryId: string } | null {
  for (const category of board.categories) {
    const tile = category.tiles.find((t) => t.id === tileId);
    if (tile) return { tile, categoryId: category.id };
  }
  return null;
}

/** 새 타일이면 카테고리 끝에 붙이고, 있던 타일이면 내용을 바꾼다 (카테고리가 달라졌으면 옮긴다). */
export function saveTile(board: Board, categoryId: string, tile: Tile): Board {
  const existing = findTile(board, tile.id);
  if (existing?.categoryId === categoryId) {
    return mapCategory(board, categoryId, (c) => ({
      ...c,
      tiles: c.tiles.map((t) => (t.id === tile.id ? tile : t)),
    }));
  }
  const without = existing ? removeTile(board, tile.id) : board;
  return mapCategory(without, categoryId, (c) => ({ ...c, tiles: [...c.tiles, tile] }));
}

export function removeTile(board: Board, tileId: string): Board {
  return {
    ...board,
    categories: board.categories.map((c) =>
      c.tiles.some((t) => t.id === tileId) ? { ...c, tiles: c.tiles.filter((t) => t.id !== tileId) } : c,
    ),
  };
}

/** 타일을 다른 자리로 옮긴다. beforeTileId 가 null 이면 카테고리 맨 끝. */
export function moveTile(
  board: Board,
  tileId: string,
  toCategoryId: string,
  beforeTileId: string | null,
): Board {
  const found = findTile(board, tileId);
  if (!found || beforeTileId === tileId) return board;
  if (!board.categories.some((c) => c.id === toCategoryId)) return board;
  return mapCategory(removeTile(board, tileId), toCategoryId, (c) => {
    const index = beforeTileId ? c.tiles.findIndex((t) => t.id === beforeTileId) : -1;
    const tiles = [...c.tiles];
    tiles.splice(index < 0 ? tiles.length : index, 0, found.tile);
    return { ...c, tiles };
  });
}

const topZ = (board: Board) => board.notes.reduce((max, n) => Math.max(max, n.z), 0);

export function addNote(board: Board, at: { x: number; y: number }): { board: Board; note: Note } {
  const note: Note = {
    id: newId(),
    text: "",
    x: Math.max(0, Math.round(at.x)),
    y: Math.max(0, Math.round(at.y)),
    w: 190,
    h: 150,
    color: "yellow",
    z: topZ(board) + 1,
    kind: "text",
    items: [],
    fontSize: 13,
    marks: [],
  };
  return { board: { ...board, notes: [...board.notes, note] }, note };
}

export function updateNote(board: Board, id: string, patch: Partial<Note>): Board {
  return {
    ...board,
    notes: board.notes.map((n) => {
      if (n.id !== id) return n;
      const next = { ...n, ...patch };
      return {
        ...next,
        x: Math.max(0, Math.round(next.x)),
        y: Math.max(0, Math.round(next.y)),
        w: Math.max(NOTE_MIN_W, Math.round(next.w)),
        h: Math.max(NOTE_MIN_H, Math.round(next.h)),
        fontSize: Math.min(NOTE_FONT_MAX, Math.max(NOTE_FONT_MIN, next.fontSize)),
      };
    }),
  };
}

export function removeNote(board: Board, id: string): Board {
  return { ...board, notes: board.notes.filter((n) => n.id !== id) };
}

/** 메모를 다른 메모들 위로 올린다. 이미 맨 위면 그대로 둔다. */
export function raiseNote(board: Board, id: string): Board {
  const note = board.notes.find((n) => n.id === id);
  if (!note) return board;
  const others = board.notes.filter((n) => n.id !== id);
  if (others.every((n) => n.z < note.z)) return board;
  return updateNote(board, id, { z: topZ(board) + 1 });
}

/** 체크리스트에서 끝낸 항목을 글 메모로 바꿀 때 줄 앞에 붙이는 표시 */
const DONE_MARK = "✓ ";

export const newItem = (rich: Rich = { text: "", marks: [] }, done = false): NoteItem => ({
  id: newId(),
  text: rich.text,
  marks: rich.marks,
  done,
});

/** 글 메모의 각 줄을 체크리스트 항목으로 바꾼다. 빈 줄은 버리고, 항목이 하나도 없으면 빈 항목 하나를 둔다. 서식은 따라간다. */
export function textToItems(note: Rich): NoteItem[] {
  const items: NoteItem[] = [];
  let start = 0;
  for (const line of note.text.split("\n")) {
    const end = start + line.length;
    if (line.trim()) {
      const done = line.startsWith(DONE_MARK);
      items.push(newItem(sliceRich(note, done ? start + DONE_MARK.length : start, end), done));
    }
    start = end + 1;
  }
  return items.length > 0 ? items : [newItem()];
}

export function itemsToText(items: NoteItem[]): Rich {
  const lines = items
    .filter((item) => item.text.trim().length > 0)
    .map((item) => (item.done ? joinRich([{ text: DONE_MARK.trimEnd(), marks: [] }, item], " ") : item));
  return joinRich(lines, "\n");
}

/** 글 메모와 체크리스트를 서로 바꾼다. 내용은 줄 단위로 옮겨 간다. */
export function switchNoteKind(board: Board, id: string): Board {
  const note = board.notes.find((n) => n.id === id);
  if (!note) return board;
  return note.kind === "list"
    ? updateNote(board, id, { kind: "text", ...itemsToText(note.items), items: [] })
    : updateNote(board, id, { kind: "list", items: textToItems(note), text: "", marks: [] });
}

export function patchSettings(board: Board, patch: Partial<Settings>): Board {
  return { ...board, settings: { ...board.settings, ...patch } };
}
