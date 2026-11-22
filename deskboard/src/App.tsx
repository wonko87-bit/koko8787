import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { CSSProperties, MouseEvent, PointerEvent as ReactPointerEvent } from "react";
import { api } from "./api";
import type { AvailableUpdate, FileDrop, ResizeEdge } from "./api";
import * as ops from "./board";
import { ConfirmDialog, NameDialog, suggestName, TileDialog } from "./components/Dialogs";
import { LockIcon, PlusIcon, SlidersIcon } from "./components/icons";
import { ContextMenu } from "./components/Menu";
import type { MenuState } from "./components/Menu";
import { NoteBoard } from "./components/Notes";
import { SettingsPanel } from "./components/Settings";
import { matchesTile, Shortcuts } from "./components/Shortcuts";
import { useTileIcons } from "./tileIcons";
import type { Board, Category, Note, Page, Tile } from "./types";

interface Toast {
  text: string;
  kind: "error" | "info";
  /** 알림에 붙는 단추 (지운 것 되돌리기) */
  action?: { label: string; run: () => void };
}

/** 지운 것을 되돌릴 수 있게 알림을 띄워 두는 시간 */
const UNDO_MS = 8000;

const categoryAt = (x: number, y: number) =>
  document.elementFromPoint(x, y)?.closest<HTMLElement>("[data-category-id]")?.dataset.categoryId ?? null;

type Dialog =
  | { type: "tile"; tile?: Tile; categoryId: string }
  | { type: "category"; category?: Category }
  | { type: "page"; page?: Page }
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
  const [toast, setToast] = useState<Toast | null>(null);
  const [fileDropCategoryId, setFileDropCategoryId] = useState<string | null>(null);
  const [update, setUpdate] = useState<AvailableUpdate | null>(null);
  const [installing, setInstalling] = useState(false);
  const [query, setQuery] = useState("");
  const searchRef = useRef<HTMLInputElement>(null);
  const [focusNoteId, setFocusNoteId] = useState<string | null>(null);
  const [liveSplit, setLiveSplit] = useState<number | null>(null);
  const notesRef = useRef<HTMLDivElement>(null);
  /** 디스크에 마지막으로 저장된 보드 */
  const savedRef = useRef<Board | null>(null);
  const boardRef = useRef<Board | null>(null);
  boardRef.current = board;

  const change = useCallback((apply: (board: Board) => Board) => {
    setBoard((current) => (current ? apply(current) : current));
  }, []);
  const showError = useCallback((e: unknown) => setToast({ text: String(e), kind: "error" }), []);
  /** 방금 지운 것을 잠깐 동안 되돌릴 수 있게 한다. */
  const offerUndo = useCallback(
    (text: string, restore: (board: Board) => Board) =>
      setToast({
        text,
        kind: "info",
        action: {
          label: "되돌리기",
          run: () => {
            change(restore);
            setToast(null);
          },
        },
      }),
    [change],
  );

  // 되돌리기 알림이 떠 있는 동안에는 Ctrl+Z 로도 되돌린다 (글을 쓰는 중의 Ctrl+Z 는 글자 되돌리기로 둔다).
  useEffect(() => {
    const undo = toast?.action;
    if (!undo) return;
    const onKey = (e: KeyboardEvent) => {
      if (!e.ctrlKey || e.key.toLowerCase() !== "z") return;
      if ((e.target as HTMLElement | null)?.closest("input, textarea, [contenteditable]")) return;
      e.preventDefault();
      undo.run();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [toast]);

  // 트레이 메뉴에서 "업데이트 확인"을 고르면 할 일. 실제 내용은 아래에서 채운다.
  const checkFromTray = useRef<() => void>(() => {});
  useEffect(() => {
    let stop: (() => void) | undefined;
    let cancelled = false;
    api.onUpdateRequest(() => checkFromTray.current()).then(
      (unlisten) => (cancelled ? unlisten() : (stop = unlisten)),
      () => {},
    );
    return () => {
      cancelled = true;
      stop?.();
    };
  }, []);

  // 단축키나 트레이로 불러오면 바로 타자를 쳐서 바로가기를 찾을 수 있게 찾기 칸에 커서를 둔다.
  useEffect(() => {
    let stop: (() => void) | undefined;
    let cancelled = false;
    api.onSummon(() => {
      searchRef.current?.focus();
      searchRef.current?.select();
    }).then((unlisten) => (cancelled ? unlisten() : (stop = unlisten)), () => {});
    return () => {
      cancelled = true;
      stop?.();
    };
  }, []);

  useEffect(() => {
    api.loadBoard().then(
      (loaded) => {
        savedRef.current = loaded;
        setBoard(loaded);
        // 시작할 때의 업데이트 확인은 조용히 한다. 오프라인이거나 아직 릴리스가 없을 수 있다.
        if (loaded.settings.checkUpdates) api.checkUpdate().then(setUpdate, () => {});
      },
      (e) => setLoadError(String(e)),
    );
  }, []);

  // 탐색기에서 끌어다 놓은 파일·폴더를 바로가기로 만든다.
  useEffect(() => {
    const onDrop = async (drop: FileDrop) => {
      if (drop.type === "leave") return setFileDropCategoryId(null);
      const under = categoryAt(drop.x, drop.y) ?? boardRef.current?.categories[0]?.id ?? null;
      if (drop.type === "over") return setFileDropCategoryId(under);
      setFileDropCategoryId(null);
      const tiles = await Promise.all(
        drop.paths.map(async (path): Promise<Tile> => {
          const kind = await api.classify(path).then((info) => info.kind, () => "file" as const);
          return { id: ops.newId(), name: suggestName(path), target: path, kind, color: "", icon: "" };
        }),
      );
      change((b) => {
        // 놓은 자리의 카테고리에, 없으면 첫 카테고리에 넣는다. 카테고리가 하나도 없으면 만든다.
        let next = b.categories.length > 0 ? b : ops.addCategory(b, "바로가기");
        const categoryId = next.categories.some((c) => c.id === under) ? under! : next.categories[0].id;
        for (const tile of tiles) next = ops.saveTile(next, categoryId, tile);
        return next;
      });
    };
    let stop: (() => void) | undefined;
    let cancelled = false;
    api.onFileDrop(onDrop).then((unlisten) => (cancelled ? unlisten() : (stop = unlisten)), () => {});
    return () => {
      cancelled = true;
      stop?.();
    };
  }, [change]);

  const allTiles = useMemo(() => board?.categories.flatMap((c) => c.tiles) ?? [], [board?.categories]);
  const iconFor = useTileIcons(allTiles);

  // 바뀐 내용은 입력이 잠깐 멈췄을 때 저장한다.
  useEffect(() => {
    if (!board || board === savedRef.current) return;
    const timer = setTimeout(() => {
      api.saveBoard(board).then(
        () => (savedRef.current = board),
        (e) => showError(e),
      );
    }, SAVE_DELAY_MS);
    return () => clearTimeout(timer);
  }, [board]);

  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => setToast(null), toast.action ? UNDO_MS : toast.kind === "error" ? 6000 : 3500);
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
    (id: string, patch: Partial<Note>) => change((b) => ops.updateNote(b, id, patch)),
    [change],
  );
  const updateNoteWith = useCallback(
    (id: string, decide: (note: Note) => Partial<Note>) =>
      change((b) => {
        const note = b.notes.find((n) => n.id === id);
        return note ? ops.updateNote(b, id, decide(note)) : b;
      }),
    [change],
  );
  const raiseNote = useCallback((id: string) => change((b) => ops.raiseNote(b, id)), [change]);
  const switchNoteKind = useCallback((id: string) => change((b) => ops.switchNoteKind(b, id)), [change]);
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
      change((b) => ops.removeNote(b, note.id));
      if (!ops.isNoteEmpty(note)) offerUndo("메모를 뗐습니다.", (b) => ops.restoreNote(b, note));
    },
    [change, offerUndo],
  );

  if (loadError) return <div className="widget fatal">보드를 불러오지 못했습니다: {loadError}</div>;
  if (!board) return null;

  const { settings, categories } = board;

  // ---- 바로가기 ----
  const launch = (tile: Tile, elevated = false) => api.launch(tile.target, elevated).catch(showError);

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
        // 관리자 권한 실행은 프로그램에만 뜻이 있다. Windows 가 실행 전에 확인 창을 띄운다.
        ...(tile.kind === "program" ? [{ label: "관리자 권한으로 열기", onSelect: () => launch(tile, true) }] : []),
        { label: "편집…", onSelect: () => setDialog({ type: "tile", tile, categoryId }) },
        {
          label: "삭제",
          danger: true,
          onSelect: () => {
            const index = categories.find((c) => c.id === categoryId)?.tiles.findIndex((t) => t.id === tile.id) ?? 0;
            change((b) => ops.removeTile(b, tile.id));
            offerUndo(`'${tile.name}' 바로가기를 지웠습니다.`, (b) => ops.restoreTile(b, tile, categoryId, index));
          },
        },
      ],
    });

  const categoryMenu = (at: { x: number; y: number }, category: Category) => {
    const index = categories.findIndex((c) => c.id === category.id);
    const remove = () => {
      change((b) => ops.removeCategory(b, category.id));
      offerUndo(`'${category.name}' 카테고리를 지웠습니다.`, (b) => ops.restoreCategory(b, category, index));
    };
    setMenu({
      ...at,
      items: [
        { label: "바로가기 추가…", onSelect: () => openTileDialog(category.id) },
        { label: "이름 바꾸기…", onSelect: () => setDialog({ type: "category", category }) },
        { label: "위로", disabled: index === 0, onSelect: () => change((b) => ops.shiftCategory(b, category.id, -1)) },
        {
          label: "아래로",
          disabled: index === categories.length - 1,
          onSelect: () => change((b) => ops.shiftCategory(b, category.id, 1)),
        },
        {
          label: "카테고리 삭제",
          danger: true,
          onSelect: remove,
        },
      ],
    });
  };

  // ---- 메모판 페이지 ----
  const pageNotes = board.notes.filter((n) => n.page === settings.page);

  const pageMenu = (at: { x: number; y: number }, page: Page) => {
    const index = board.pages.findIndex((p) => p.id === page.id);
    const notesOnPage = board.notes.filter((n) => n.page === page.id);
    setMenu({
      ...at,
      items: [
        { label: "이름 바꾸기…", onSelect: () => setDialog({ type: "page", page }) },
        { label: "왼쪽으로", disabled: index === 0, onSelect: () => change((b) => ops.shiftPage(b, page.id, -1)) },
        {
          label: "오른쪽으로",
          disabled: index === board.pages.length - 1,
          onSelect: () => change((b) => ops.shiftPage(b, page.id, 1)),
        },
        {
          label: notesOnPage.length > 0 ? `페이지 삭제 (메모 ${notesOnPage.length}장 포함)` : "페이지 삭제",
          danger: true,
          disabled: board.pages.length <= 1,
          onSelect: () => {
            change((b) => ops.removePage(b, page.id));
            offerUndo(`'${page.name}' 페이지를 지웠습니다.`, (b) => ops.restorePage(b, page, index, notesOnPage));
          },
        },
      ],
    });
  };

  /** 메모 띠를 우클릭하면 다른 페이지로 보낼 수 있다. */
  const noteBarMenu = (at: { x: number; y: number }, note: Note) => {
    const others = board.pages.filter((p) => p.id !== note.page);
    if (others.length === 0) return;
    setMenu({
      ...at,
      items: others.map((p) => ({
        label: `'${p.name}' 페이지로 보내기`,
        onSelect: () => changeNote(note.id, { page: p.id }),
      })),
    });
  };

  // ---- 바로가기 찾기 ----
  const firstMatch = query.trim()
    ? categories.flatMap((c) => c.tiles).find((t) => matchesTile(t, query))
    : undefined;

  // ---- 위젯 배치 ----
  const beginMove = (e: MouseEvent) => {
    if (locked || e.button !== 0 || (e.target as HTMLElement).closest("button, input")) return;
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
        change((b) => ops.patchSettings(b, { split }));
      }
      setLiveSplit(null);
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", end);
    handle.addEventListener("pointercancel", end);
  };

  const quit = async () => {
    // 아직 저장을 기다리는 변경이 있으면 쓰고 나서 끈다.
    await flushPendingSave().catch(() => {});
    await api.quit();
  };

  // ---- 백업 · 업데이트 ----
  const flushPendingSave = async () => {
    const current = boardRef.current;
    if (current && current !== savedRef.current) {
      await api.saveBoard(current);
      savedRef.current = current;
    }
  };

  const exportBoard = async () => {
    try {
      await flushPendingSave();
      const path = await api.exportBoard();
      if (path) setToast({ text: `내보냈습니다: ${path}`, kind: "info" });
    } catch (e) {
      showError(e);
    }
  };

  const importBoard = () =>
    setDialog({
      type: "confirm",
      title: "보드 가져오기",
      message: "지금의 바로가기·메모·설정이 모두 백업 파일의 내용으로 바뀝니다. 계속할까요?",
      confirmLabel: "파일 선택…",
      onConfirm: async () => {
        try {
          const imported = await api.importBoard();
          if (!imported) return;
          savedRef.current = imported;
          setBoard(imported);
          setToast({ text: "백업에서 보드를 가져왔습니다.", kind: "info" });
        } catch (e) {
          showError(e);
        }
      },
    });

  /** 설정의 "지금 확인". 결과를 한 줄 글로 돌려준다. */
  const checkUpdateNow = async () => {
    try {
      const found = await api.checkUpdate();
      setUpdate(found);
      return found ? "" : `이미 최신 버전입니다 (v${await api.version()}).`;
    } catch (e) {
      return `업데이트를 확인하지 못했습니다. (${String(e)})`;
    }
  };

  /** 설치하기 전에 한 번 묻는다. */
  const askToInstall = (found: AvailableUpdate) => {
    if (installing) return;
    setDialog({
      type: "confirm",
      title: "업데이트 설치",
      message: `새 버전 v${found.version} 을 내려받아 설치합니다. 설치가 끝나면 DeskBoard 가 다시 시작됩니다.`,
      confirmLabel: "설치",
      onConfirm: async () => {
        setInstalling(true);
        try {
          await flushPendingSave();
          await found.install();
        } catch (e) {
          showError(e);
          setInstalling(false);
        }
      },
    });
  };

  const installUpdate = () => {
    if (update) askToInstall(update);
  };

  // 트레이 메뉴의 "업데이트 확인". 위젯의 단추가 화면 밖에 있어도 업데이트할 수 있게,
  // 새 버전이 있으면 바로 설치할지 묻고 없으면 그렇다고 알린다.
  checkFromTray.current = async () => {
    setToast({ text: "업데이트를 확인하는 중…", kind: "info" });
    try {
      const found = await api.checkUpdate();
      setUpdate(found);
      if (found) {
        setToast(null);
        askToInstall(found);
      } else {
        setToast({ text: `이미 최신 버전입니다 (v${await api.version()}).`, kind: "info" });
      }
    } catch (e) {
      showError(`업데이트를 확인하지 못했습니다. (${String(e)})`);
    }
  };

  const widgetStyle = {
    "--alpha": settings.opacity,
    "--split": liveSplit ?? settings.split,
    ...(settings.noteFont ? { "--note-font": JSON.stringify(settings.noteFont) } : {}),
  } as CSSProperties;

  return (
    <div className={locked ? "widget" : "widget editing"} data-theme={settings.theme} style={widgetStyle}>
      <header className="header" onMouseDown={beginMove}>
        <span className="title">DeskBoard</span>
        <button className="text-button" onClick={() => openTileDialog()}>
          <PlusIcon /> 바로가기
        </button>
        <button className="text-button" onClick={() => addNote()}>
          <PlusIcon /> 메모
        </button>
        <input
          ref={searchRef}
          className="search"
          type="search"
          value={query}
          placeholder="바로가기 찾기"
          aria-label="바로가기 찾기"
          spellCheck={false}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => {
            // 한글을 조합하는 중의 Enter 는 글자를 확정하는 것이다.
            if (e.nativeEvent.isComposing) return;
            if (e.key === "Enter" && firstMatch) {
              launch(firstMatch);
              setQuery("");
            } else if (e.key === "Escape") {
              setQuery("");
              e.currentTarget.blur();
            }
          }}
        />
        <span className="header-hint">
          {locked ? "" : "여기를 끌어 이동 · 가장자리를 끌어 크기 조절 · 가운데 선으로 영역 조절"}
        </span>
        {update && (
          <button className="update-pill" title="새 버전 설치" disabled={installing} onClick={installUpdate}>
            {installing ? "설치하는 중…" : `v${update.version} 업데이트`}
          </button>
        )}
        <button
          className={locked ? "icon-button" : "icon-button active"}
          title={locked ? "위치·크기 조절하기 (잠금 풀기)" : "위치·크기 고정하기 (잠그기)"}
          aria-pressed={!locked}
          onClick={() => change((b) => ops.patchSettings(b, { locked: !locked }))}
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
          iconFor={iconFor}
          fileDropCategoryId={fileDropCategoryId}
          query={query}
          onLaunch={launch}
          onTileMenu={tileMenu}
          onCategoryMenu={categoryMenu}
          onAddTile={openTileDialog}
          onAddCategory={() => setDialog({ type: "category" })}
          onMoveTile={(tileId, toCategoryId, beforeTileId) =>
            change((b) => ops.moveTile(b, tileId, toCategoryId, beforeTileId))
          }
        />
        <div className="splitter" onPointerDown={beginSplit} />
        <NoteBoard
          notes={pageNotes}
          pages={board.pages}
          activePage={settings.page}
          focusId={focusNoteId}
          boardRef={notesRef}
          onAdd={addNote}
          onChange={changeNote}
          onUpdate={updateNoteWith}
          onSwitchKind={switchNoteKind}
          onRemove={removeNote}
          onRaise={raiseNote}
          onBarMenu={noteBarMenu}
          onSelectPage={(page) => change((b) => ops.patchSettings(b, { page }))}
          onAddPage={() => setDialog({ type: "page" })}
          onPageMenu={pageMenu}
        />
      </div>

      {!locked &&
        EDGES.map((edge) => <div key={edge} className={`resize-edge edge-${edge}`} onMouseDown={beginResize(edge)} />)}

      {settingsOpen && (
        <SettingsPanel
          settings={settings}
          updateVersion={update?.version ?? null}
          onChange={(patch) => change((b) => ops.patchSettings(b, patch))}
          onCheckUpdate={checkUpdateNow}
          onInstallUpdate={installUpdate}
          onExport={exportBoard}
          onImport={importBoard}
          onOpenBackups={() => api.openBackupFolder().catch(showError)}
          onResetWindow={() => api.resetWindow().catch(showError)}
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
            change((b) => ops.saveTile(b, categoryId, tile));
            closeDialog();
          }}
          onClose={closeDialog}
        />
      )}
      {dialog?.type === "page" && (
        <NameDialog
          title={dialog.page ? "페이지 이름 바꾸기" : "페이지 추가"}
          initial={dialog.page?.name ?? ""}
          confirmLabel={dialog.page ? "저장" : "추가"}
          onSubmit={(name) => {
            const page = dialog.page;
            change((b) => (page ? ops.renamePage(b, page.id, name) : ops.addPage(b, name)));
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
            change((b) => (category ? ops.renameCategory(b, category.id, name) : ops.addCategory(b, name)));
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
        <div className={`toast ${toast.kind}`} role={toast.kind === "error" ? "alert" : "status"} onClick={() => setToast(null)}>
          <span>{toast.text}</span>
          {toast.action && (
            <button
              className="toast-action"
              onClick={(e) => {
                e.stopPropagation();
                toast.action?.run();
              }}
            >
              {toast.action.label}
            </button>
          )}
        </div>
      )}
    </div>
  );
}
