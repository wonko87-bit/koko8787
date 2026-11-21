// 보드를 바꾸는 순수 함수들. 원본은 건드리지 않고 새 보드를 돌려준다.

import { BOX_ORIGIN, NOTE_FONT_MAX, NOTE_FONT_MIN, NOTE_MIN_H, NOTE_MIN_W } from "./types";
import { joinRich, sliceRich } from "./richText";
import type { Board, Category, Note, NoteItem, Page, Rich, Settings, TextBox, Tile } from "./types";

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
    // 바로 적을 수 있게 빈 글상자 하나를 놓아 둔다. 아무것도 안 적으면 사라진다.
    boxes: [newBox(BOX_ORIGIN)],
    page: board.settings.page,
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

export const newBox = (at: { x: number; y: number }, rich: Rich = { text: "", marks: [] }): TextBox => ({
  id: newId(),
  x: Math.max(0, Math.round(at.x)),
  y: Math.max(0, Math.round(at.y)),
  w: 0,
  text: rich.text,
  marks: rich.marks,
});

/** 값이 후보 중 하나와 `within` 안으로 가까우면 그 후보에 붙인다 (글상자 줄 맞춤). 붙었으면 guide 에 그 자리가 온다. */
export function snap(value: number, candidates: number[], within: number): { value: number; guide: number | null } {
  let nearest: number | null = null;
  for (const candidate of candidates) {
    const distance = Math.abs(candidate - value);
    if (distance <= within && (nearest === null || distance < Math.abs(nearest - value))) nearest = candidate;
  }
  return nearest === null ? { value, guide: null } : { value: nearest, guide: nearest };
}

// ---- 메모판 페이지 ----

export function addPage(board: Board, name: string): Board {
  const page: Page = { id: newId(), name };
  return { ...board, pages: [...board.pages, page], settings: { ...board.settings, page: page.id } };
}

export function renamePage(board: Board, id: string, name: string): Board {
  return { ...board, pages: board.pages.map((p) => (p.id === id ? { ...p, name } : p)) };
}

/** 페이지와 거기 붙은 메모를 지운다. 마지막 남은 페이지는 지우지 않는다. */
export function removePage(board: Board, id: string): Board {
  if (board.pages.length <= 1 || !board.pages.some((p) => p.id === id)) return board;
  const pages = board.pages.filter((p) => p.id !== id);
  const page = board.settings.page === id ? pages[0].id : board.settings.page;
  return { ...board, pages, notes: board.notes.filter((n) => n.page !== id), settings: { ...board.settings, page } };
}

/** 페이지 탭을 한 칸 왼쪽(-1)이나 오른쪽(+1)으로 옮긴다. */
export function shiftPage(board: Board, id: string, delta: -1 | 1): Board {
  const from = board.pages.findIndex((p) => p.id === id);
  const to = from + delta;
  if (from < 0 || to < 0 || to >= board.pages.length) return board;
  const pages = [...board.pages];
  [pages[from], pages[to]] = [pages[to], pages[from]];
  return { ...board, pages };
}

// ---- 지운 것 되살리기 ----

/** 지웠던 메모를 다시 붙인다. 있던 페이지가 없어졌으면 지금 페이지에 붙인다. */
export function restoreNote(board: Board, note: Note): Board {
  if (board.notes.some((n) => n.id === note.id)) return board;
  const page = board.pages.some((p) => p.id === note.page) ? note.page : board.settings.page;
  return { ...board, notes: [...board.notes, { ...note, page }] };
}

/** 지웠던 타일을 있던 자리에 다시 넣는다. 카테고리가 없어졌으면 첫 카테고리 끝에 넣는다. */
export function restoreTile(board: Board, tile: Tile, categoryId: string, index: number): Board {
  if (findTile(board, tile.id) || board.categories.length === 0) return board;
  const target = board.categories.some((c) => c.id === categoryId) ? categoryId : board.categories[0].id;
  return {
    ...board,
    categories: board.categories.map((c) => {
      if (c.id !== target) return c;
      const tiles = [...c.tiles];
      tiles.splice(c.id === categoryId ? Math.min(index, tiles.length) : tiles.length, 0, tile);
      return { ...c, tiles };
    }),
  };
}

export function restoreCategory(board: Board, category: Category, index: number): Board {
  if (board.categories.some((c) => c.id === category.id)) return board;
  const categories = [...board.categories];
  categories.splice(Math.min(index, categories.length), 0, category);
  return { ...board, categories };
}

export function restorePage(board: Board, page: Page, index: number, notes: Note[]): Board {
  if (board.pages.some((p) => p.id === page.id)) return board;
  const pages = [...board.pages];
  pages.splice(Math.min(index, pages.length), 0, page);
  const missing = notes.filter((note) => !board.notes.some((n) => n.id === note.id));
  return { ...board, pages, notes: [...board.notes, ...missing], settings: { ...board.settings, page: page.id } };
}

/** 적힌 것이 하나도 없는 메모인가 */
export function isNoteEmpty(note: Note): boolean {
  const parts: Rich[] = note.kind === "list" ? note.items : note.boxes;
  return parts.every((part) => !part.text.trim());
}

/** 글의 각 줄을 체크리스트 항목으로 바꾼다. 빈 줄은 버린다. 서식은 따라간다. */
function linesToItems(rich: Rich): NoteItem[] {
  const items: NoteItem[] = [];
  let start = 0;
  for (const line of rich.text.split("\n")) {
    const end = start + line.length;
    if (line.trim()) {
      const done = line.startsWith(DONE_MARK);
      items.push(newItem(sliceRich(rich, done ? start + DONE_MARK.length : start, end), done));
    }
    start = end + 1;
  }
  return items;
}

/** 글상자들을 위에서 아래로, 같은 높이면 왼쪽에서 오른쪽으로 읽어 체크리스트 항목으로 바꾼다. */
export function boxesToItems(boxes: TextBox[]): NoteItem[] {
  const ordered = [...boxes].sort((a, b) => a.y - b.y || a.x - b.x);
  const items = ordered.flatMap(linesToItems);
  return items.length > 0 ? items : [newItem()];
}

export function itemsToText(items: NoteItem[]): Rich {
  const lines = items
    .filter((item) => item.text.trim().length > 0)
    .map((item) => (item.done ? joinRich([{ text: DONE_MARK.trimEnd(), marks: [] }, item], " ") : item));
  return joinRich(lines, "\n");
}

/** 글 메모와 체크리스트를 서로 바꾼다. 내용은 줄 단위로 옮겨 가고, 글상자의 자리는 남지 않는다. */
export function switchNoteKind(board: Board, id: string): Board {
  const note = board.notes.find((n) => n.id === id);
  if (!note) return board;
  if (note.kind === "list") {
    const text = itemsToText(note.items);
    return updateNote(board, id, { kind: "text", items: [], boxes: text.text ? [newBox(BOX_ORIGIN, text)] : [] });
  }
  return updateNote(board, id, { kind: "list", items: boxesToItems(note.boxes), boxes: [] });
}

export function patchSettings(board: Board, patch: Partial<Settings>): Board {
  return { ...board, settings: { ...board.settings, ...patch } };
}
