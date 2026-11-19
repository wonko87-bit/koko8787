import { useCallback, useEffect, useRef, useState } from "react";
import type { CSSProperties, MouseEvent, PointerEvent as ReactPointerEvent } from "react";
import { api } from "./api";
import type { ResizeEdge } from "./api";
import * as ops from "./board";
import { ConfirmDialog, NameDialog, TileDialog } from "./components/Dialogs";
import { LockIcon, PlusIcon, SlidersIcon } from "./components/icons";
import { ContextMenu } from "./components/Menu";
import type { MenuState } from "./components/Menu";
import { NoteBoard } from "./components/Notes";
import { SettingsPanel } from "./components/Settings";
import { Shortcuts } from "./components/Shortcuts";
import type { Board, Category, Note, Tile } from "./types";

type Dialog =
  | { type: "tile"; tile?: Tile; categoryId: string }
  | { type: "category"; category?: Category }
  | { type: "confirm"; title: string; message: string; confirmLabel: string; onConfirm: () => void };

const EDGES: ResizeEdge[] = ["North", "South", "East", "West", "NorthEast", "NorthWest", "SouthEast", "SouthWest"];
const SAVE_DELAY_MS = 400;
const clampSplit = (value: number) => Math.min(0.8, Math.max(0.2, value));

export default function App() {
  const [board, setBoard] = useState<Board | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [dialog, setDialog] = useState<Dialog | null>(null);
  const [menu, setMenu] = useState<MenuState | null>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [toast, setToast] = useState<string | null>(null);
  const [focusNoteId, setFocusNoteId] = useState<string | null>(null);
  const [liveSplit, setLiveSplit] = useState<number | null>(null);
  const notesRef = useRef<HTMLDivElement>(null);
  /** 디스크에 마지막으로 저장된 보드 */
  const savedRef = useRef<Board | null>(null);
  const boardRef = useRef<Board | null>(null);
  boardRef.current = board;

  const update = useCallback((change: (board: Board) => Board) => {
    setBoard((current) => (current ? change(current) : current));
  }, []);

  useEffect(() => {
    api.loadBoard().then(
      (loaded) => {
        savedRef.current = loaded;
        setBoard(loaded);
      },
      (e) => setLoadError(String(e)),
    );
  }, []);

  // 바뀐 내용은 입력이 잠깐 멈췄을 때 저장한다.
  useEffect(() => {
    if (!board || board === savedRef.current) return;
    const timer = setTimeout(() => {
      api.saveBoard(board).then(
        () => (savedRef.current = board),
        (e) => setToast(String(e)),
      );
    }, SAVE_DELAY_MS);
    return () => clearTimeout(timer);
  }, [board]);

  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => setToast(null), 4500);
    return () => clearTimeout(timer);
  }, [toast]);

  const locked = board?.settings.locked ?? true;
  useEffect(() => {
    if (board) api.setResizable(!locked).catch(() => {});
  }, [board === null, locked]); // eslint-disable-line react-hooks/exhaustive-deps

  const closeMenu = useCallback(() => setMenu(null), []);
  const closeDialog = useCallback(() => setDialog(null), []);
  const closeSettings = useCallback(() => setSettingsOpen(false), []);

  // ---- 메모 ----
  const changeNote = useCallback(
    (id: string, patch: Partial<Note>) => update((b) => ops.updateNote(b, id, patch)),
    [update],
  );
  const raiseNote = useCallback((id: string) => update((b) => ops.raiseNote(b, id)), [update]);
  const addNote = useCallback(
    (at?: { x: number; y: number }) => {
      const current = boardRef.current;
      if (!current) return;
      // 버튼으로 만들 때는 지금 보이는 자리에 조금씩 어긋나게 붙인다.
      const el = notesRef.current;
      const step = (current.notes.length % 6) * 22;
      const spot = at ?? { x: (el?.scrollLeft ?? 0) + 16 + step, y: (el?.scrollTop ?? 0) + 14 + step };
      const { board: next, note } = ops.addNote(current, spot);
      setBoard(next);
      setFocusNoteId(note.id);
    },
    [],
  );
  const removeNote = useCallback(
    (note: Note) => {
      const remove = () => update((b) => ops.removeNote(b, note.id));
      if (!note.text.trim()) return remove();
      setDialog({
        type: "confirm",
        title: "메모 떼기",
        message: "이 메모를 떼어 버릴까요? 적어 둔 내용은 되돌릴 수 없습니다.",
        confirmLabel: "떼기",
        onConfirm: remove,
      });
    },
    [update],
  );

  if (loadError) return <div className="widget fatal">보드를 불러오지 못했습니다: {loadError}</div>;
  if (!board) return null;

  const { settings, categories } = board;

  // ---- 바로가기 ----
  const launch = (tile: Tile) => api.launch(tile.target).catch((e) => setToast(String(e)));

  const openTileDialog = (categoryId?: string) => {
    if (categoryId ?? categories[0]) {
      setDialog({ type: "tile", categoryId: categoryId ?? categories[0].id });
      return;
    }
    // 카테고리가 하나도 없으면 담을 곳부터 만든다.
    const next = ops.addCategory(board, "바로가기");
    setBoard(next);
    setDialog({ type: "tile", categoryId: next.categories[0].id });
  };

  const tileMenu = (at: { x: number; y: number }, tile: Tile, categoryId: string) =>
    setMenu({
      ...at,
      items: [
        { label: "열기", onSelect: () => launch(tile) },
        { label: "편집…", onSelect: () => setDialog({ type: "tile", tile, categoryId }) },
        {
          label: "삭제",
          danger: true,
          onSelect: () =>
            setDialog({
              type: "confirm",
              title: "바로가기 삭제",
              message: `'${tile.name}' 바로가기를 삭제할까요? 원본 파일·폴더는 그대로 남습니다.`,
              confirmLabel: "삭제",
              onConfirm: () => update((b) => ops.removeTile(b, tile.id)),
            }),
        },
      ],
    });

  const categoryMenu = (at: { x: number; y: number }, category: Category) => {
    const index = categories.findIndex((c) => c.id === category.id);
    const remove = () => update((b) => ops.removeCategory(b, category.id));
    setMenu({
      ...at,
      items: [
        { label: "바로가기 추가…", onSelect: () => openTileDialog(category.id) },
        { label: "이름 바꾸기…", onSelect: () => setDialog({ type: "category", category }) },
        { label: "위로", disabled: index === 0, onSelect: () => update((b) => ops.shiftCategory(b, category.id, -1)) },
        {
          label: "아래로",
          disabled: index === categories.length - 1,
          onSelect: () => update((b) => ops.shiftCategory(b, category.id, 1)),
        },
        {
          label: "카테고리 삭제",
          danger: true,
          onSelect: () =>
            category.tiles.length === 0
              ? remove()
              : setDialog({
                  type: "confirm",
                  title: "카테고리 삭제",
                  message: `'${category.name}' 카테고리와 안에 있는 바로가기 ${category.tiles.length}개를 삭제할까요?`,
                  confirmLabel: "삭제",
                  onConfirm: remove,
                }),
        },
      ],
    });
  };

  // ---- 위젯 배치 ----
  const beginMove = (e: MouseEvent) => {
    if (locked || e.button !== 0 || (e.target as HTMLElement).closest("button")) return;
    e.preventDefault();
    api.startMove().catch(() => {});
  };

  const beginResize = (edge: ResizeEdge) => (e: MouseEvent) => {
    if (e.button !== 0) return;
    e.preventDefault();
    api.startResize(edge).catch(() => {});
  };

  const beginSplit = (e: ReactPointerEvent<HTMLDivElement>) => {
    if (locked || e.button !== 0) return;
    e.preventDefault();
    const handle = e.currentTarget;
    handle.setPointerCapture(e.pointerId);
    const body = handle.parentElement!.getBoundingClientRect();
    let latest: number | null = null;
    const move = (ev: PointerEvent) => {
      latest = clampSplit((ev.clientX - body.left) / body.width);
      setLiveSplit(latest);
    };
    const end = () => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", end);
      handle.removeEventListener("pointercancel", end);
      if (latest !== null) {
        const split = latest;
        update((b) => ops.patchSettings(b, { split }));
      }
      setLiveSplit(null);
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", end);
    handle.addEventListener("pointercancel", end);
  };

  const quit = async () => {
    // 아직 저장을 기다리는 변경이 있으면 쓰고 나서 끈다.
    const current = boardRef.current;
    if (current && current !== savedRef.current) await api.saveBoard(current).catch(() => {});
    await api.quit();
  };

  const widgetStyle = {
    "--alpha": settings.opacity,
    "--split": liveSplit ?? settings.split,
  } as CSSProperties;

  return (
    <div className={locked ? "widget" : "widget editing"} style={widgetStyle}>
      <header className="header" onMouseDown={beginMove}>
        <span className="title">DeskBoard</span>
        <button className="text-button" onClick={() => openTileDialog()}>
          <PlusIcon /> 바로가기
        </button>
        <button className="text-button" onClick={() => addNote()}>
          <PlusIcon /> 메모
        </button>
        <span className="header-hint">
          {locked ? "" : "여기를 끌어 이동 · 가장자리를 끌어 크기 조절 · 가운데 선으로 영역 조절"}
        </span>
        <button
          className={locked ? "icon-button" : "icon-button active"}
          title={locked ? "위치·크기 조절하기 (잠금 풀기)" : "위치·크기 고정하기 (잠그기)"}
          aria-pressed={!locked}
          onClick={() => update((b) => ops.patchSettings(b, { locked: !locked }))}
        >
          <LockIcon open={!locked} />
        </button>
        <button
          className={settingsOpen ? "icon-button active" : "icon-button"}
          title="설정"
          data-settings-toggle
          onClick={() => setSettingsOpen((open) => !open)}
        >
          <SlidersIcon />
        </button>
      </header>

      <div className="body">
        <Shortcuts
          categories={categories}
          tileSize={settings.tileSize}
          onLaunch={launch}
          onTileMenu={tileMenu}
          onCategoryMenu={categoryMenu}
          onAddTile={openTileDialog}
          onAddCategory={() => setDialog({ type: "category" })}
          onMoveTile={(tileId, toCategoryId, beforeTileId) =>
            update((b) => ops.moveTile(b, tileId, toCategoryId, beforeTileId))
          }
        />
        <div className="splitter" onPointerDown={beginSplit} />
        <NoteBoard
          notes={board.notes}
          focusId={focusNoteId}
          boardRef={notesRef}
          onAdd={addNote}
          onChange={changeNote}
          onRemove={removeNote}
          onRaise={raiseNote}
        />
      </div>

      {!locked &&
        EDGES.map((edge) => <div key={edge} className={`resize-edge edge-${edge}`} onMouseDown={beginResize(edge)} />)}

      {settingsOpen && (
        <SettingsPanel
          settings={settings}
          onChange={(patch) => update((b) => ops.patchSettings(b, patch))}
          onResetWindow={() => api.resetWindow().catch((e) => setToast(String(e)))}
          onQuit={quit}
          onClose={closeSettings}
        />
      )}

      {dialog?.type === "tile" && (
        <TileDialog
          categories={categories}
          tile={dialog.tile}
          categoryId={dialog.categoryId}
          onSave={(categoryId, tile) => {
            update((b) => ops.saveTile(b, categoryId, tile));
            closeDialog();
          }}
          onClose={closeDialog}
        />
      )}
      {dialog?.type === "category" && (
        <NameDialog
          title={dialog.category ? "카테고리 이름 바꾸기" : "카테고리 추가"}
          initial={dialog.category?.name ?? ""}
          confirmLabel={dialog.category ? "저장" : "추가"}
          onSubmit={(name) => {
            const category = dialog.category;
            update((b) => (category ? ops.renameCategory(b, category.id, name) : ops.addCategory(b, name)));
            closeDialog();
          }}
          onClose={closeDialog}
        />
      )}
      {dialog?.type === "confirm" && (
        <ConfirmDialog
          title={dialog.title}
          message={dialog.message}
          confirmLabel={dialog.confirmLabel}
          onConfirm={() => {
            dialog.onConfirm();
            closeDialog();
          }}
          onClose={closeDialog}
        />
      )}

      {menu && <ContextMenu menu={menu} onClose={closeMenu} />}
      {toast && (
        <div className="toast" role="alert" onClick={() => setToast(null)}>
          {toast}
        </div>
      )}
    </div>
  );
}
