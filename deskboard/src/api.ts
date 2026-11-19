// 백엔드(Tauri)와 이야기하는 유일한 통로.
// 브라우저에서 `npm run dev` 로만 띄웠을 때는 localStorage 로 흉내 내서 화면 작업을 할 수 있게 한다.

import { getVersion } from "@tauri-apps/api/app";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { getCurrentWindow } from "@tauri-apps/api/window";
import * as autostart from "@tauri-apps/plugin-autostart";
import { open, save } from "@tauri-apps/plugin-dialog";
import { relaunch } from "@tauri-apps/plugin-process";
import { check } from "@tauri-apps/plugin-updater";
import type { Board, TileKind } from "./types";

export interface TargetInfo {
  kind: TileKind;
  /** 경로가 지금 실제로 있는지 */
  exists: boolean;
}

export type ResizeEdge =
  | "North"
  | "South"
  | "East"
  | "West"
  | "NorthEast"
  | "NorthWest"
  | "SouthEast"
  | "SouthWest";

/** 탐색기에서 파일을 끌어 위젯 위로 가져왔을 때의 일. 좌표는 CSS 픽셀이다. */
export type FileDrop =
  | { type: "over"; x: number; y: number }
  | { type: "drop"; x: number; y: number; paths: string[] }
  | { type: "leave" };

export interface AvailableUpdate {
  version: string;
  /** 내려받아 설치하고 앱을 다시 시작한다. */
  install(): Promise<void>;
}

export interface Api {
  loadBoard(): Promise<Board>;
  saveBoard(board: Board): Promise<void>;
  launch(target: string, elevated: boolean): Promise<void>;
  classify(target: string): Promise<TargetInfo>;
  /** 타일에 그릴 아이콘 (PNG data URL). 구할 수 없으면 null */
  tileIcon(path: string, thumbnail: boolean): Promise<string | null>;
  listFonts(): Promise<string[]>;
  pickFile(): Promise<string | null>;
  pickFolder(): Promise<string | null>;
  /** 저장할 곳을 물어 보드를 파일로 내보낸다. 취소하면 null, 아니면 저장한 경로 */
  exportBoard(): Promise<string | null>;
  /** 백업 파일을 골라 보드를 통째로 바꾼다. 취소하면 null */
  importBoard(): Promise<Board | null>;
  setHotkey(hotkey: string): Promise<void>;
  /** 시작 프로그램에 등록되어 있는지. 알 수 없으면 null */
  getAutostart(): Promise<boolean | null>;
  setAutostart(enabled: boolean): Promise<boolean>;
  checkUpdate(): Promise<AvailableUpdate | null>;
  onFileDrop(handler: (drop: FileDrop) => void): Promise<() => void>;
  startMove(): Promise<void>;
  startResize(edge: ResizeEdge): Promise<void>;
  setResizable(resizable: boolean): Promise<void>;
  resetWindow(): Promise<void>;
  quit(): Promise<void>;
  version(): Promise<string>;
}

const BACKUP_FILTER = [{ name: "DeskBoard 백업", extensions: ["json"] }];

function backupName() {
  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `deskboard-${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}.json`;
}

const tauriApi: Api = {
  loadBoard: () => invoke<Board>("load_board"),
  saveBoard: (board) => invoke("save_board", { board }),
  launch: (target, elevated) => invoke("launch_target", { target, elevated }),
  classify: (target) => invoke<TargetInfo>("classify_target", { target }),
  tileIcon: (path, thumbnail) => invoke<string | null>("tile_icon", { path, thumbnail }),
  listFonts: () => invoke<string[]>("list_fonts"),
  pickFile: () => open({ multiple: false, directory: false, title: "파일·프로그램 선택" }),
  pickFolder: () => open({ multiple: false, directory: true, title: "폴더 선택" }),
  async exportBoard() {
    const path = await save({ title: "보드 내보내기", defaultPath: backupName(), filters: BACKUP_FILTER });
    if (!path) return null;
    await invoke("export_board", { path });
    return path;
  },
  async importBoard() {
    const path = await open({ multiple: false, directory: false, title: "보드 가져오기", filters: BACKUP_FILTER });
    return path ? invoke<Board>("import_board", { path }) : null;
  },
  setHotkey: (hotkey) => invoke("set_hotkey", { hotkey }),
  getAutostart: () => autostart.isEnabled().catch(() => null),
  async setAutostart(enabled) {
    await (enabled ? autostart.enable() : autostart.disable());
    return autostart.isEnabled();
  },
  async checkUpdate() {
    const update = await check();
    if (!update) return null;
    return {
      version: update.version,
      async install() {
        await update.downloadAndInstall();
        await relaunch();
      },
    };
  },
  onFileDrop(handler) {
    return getCurrentWebview().onDragDropEvent(({ payload }) => {
      if (payload.type === "leave") return handler({ type: "leave" });
      const x = payload.position.x / window.devicePixelRatio;
      const y = payload.position.y / window.devicePixelRatio;
      if (payload.type === "drop") handler({ type: "drop", x, y, paths: payload.paths });
      else handler({ type: "over", x, y });
    });
  },
  startMove: () => getCurrentWindow().startDragging(),
  startResize: (edge) => getCurrentWindow().startResizeDragging(edge),
  setResizable: (resizable) => getCurrentWindow().setResizable(resizable),
  resetWindow: () => invoke("reset_window"),
  quit: () => invoke("quit_app"),
  version: () => getVersion(),
};

// ---------------------------------------------------------------------------
// 브라우저 미리보기용 흉내
// ---------------------------------------------------------------------------

const STORAGE_KEY = "deskboard.preview.board";

const tile = (id: string, name: string, target: string, kind: TileKind, extra: object = {}) => ({
  id,
  name,
  target,
  kind,
  color: "" as const,
  icon: "",
  ...extra,
});

const previewBoard: Board = {
  version: 1,
  categories: [
    {
      id: "c1",
      name: "자주 가는 폴더",
      tiles: [
        tile("t1", "다운로드", "C:\\Users\\me\\Downloads", "folder"),
        tile("t2", "프로젝트", "D:\\Projects", "folder", { color: "teal" }),
        tile("t3", "2026 보고서", "D:\\업무\\2026 보고서", "folder", { icon: "emoji:📊" }),
      ],
    },
    {
      id: "c2",
      name: "프로그램",
      tiles: [
        tile("t4", "메모장", "C:\\Windows\\notepad.exe", "program"),
        tile("t5", "주간 회의록", "D:\\업무\\회의록.docx", "file"),
        tile("t6", "사내 포털", "https://example.com", "link", { color: "purple" }),
      ],
    },
  ],
  notes: [
    {
      id: "n1",
      text: "오후 3시 회의\n자료 출력하기",
      x: 16,
      y: 14,
      w: 180,
      h: 130,
      color: "yellow",
      z: 1,
      kind: "text",
      items: [],
      fontSize: 13,
    },
    {
      id: "n2",
      text: "",
      x: 150,
      y: 120,
      w: 170,
      h: 130,
      color: "pink",
      z: 2,
      kind: "list",
      items: [
        { id: "i1", text: "택배 찾기", done: true },
        { id: "i2", text: "우유 사기", done: false },
      ],
      fontSize: 13,
    },
  ],
  settings: {
    split: 0.56,
    tileSize: "m",
    opacity: 0.86,
    locked: true,
    theme: "dark",
    noteFont: "",
    hotkey: "Ctrl+Alt+D",
    checkUpdates: true,
  },
};

/** 미리보기에서 테스트가 밖에서 불러 쓸 수 있게 열어 둔 손잡이 */
interface PreviewHooks {
  drop?: (drop: FileDrop) => void;
  update?: string | null;
}
const hooks: PreviewHooks = ((window as unknown as { __deskboardPreview?: PreviewHooks }).__deskboardPreview ??= {});
let previewAutostart = false;

const previewApi: Api = {
  async loadBoard() {
    const saved = localStorage.getItem(STORAGE_KEY);
    return saved ? (JSON.parse(saved) as Board) : previewBoard;
  },
  async saveBoard(board) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(board));
  },
  async launch(target, elevated) {
    if (target.includes("없는")) throw `대상을 찾을 수 없습니다: ${target}`;
    console.info("[preview] launch", target, elevated ? "(elevated)" : "");
  },
  async classify(target) {
    if (/^[a-z][a-z0-9+.-]+:./i.test(target)) return { kind: "link", exists: true };
    const ext = /\.([a-z0-9-]+)$/i.exec(target)?.[1].toLowerCase();
    const kind = !ext ? "folder" : ["exe", "lnk", "bat", "cmd", "msi"].includes(ext) ? "program" : "file";
    return { kind, exists: !target.includes("없는") };
  },
  // 미리보기에는 진짜 아이콘이 없다. 메모장 하나만 그림을 돌려줘 아이콘 표시를 볼 수 있게 한다.
  async tileIcon(path) {
    if (!path.toLowerCase().includes("notepad")) return null;
    const svg =
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48"><rect x="8" y="4" width="32" height="40" rx="4" fill="#9fd3ff"/>' +
      '<rect x="8" y="4" width="32" height="9" rx="4" fill="#3d8bd9"/><path d="M15 21h18M15 28h18M15 35h11" stroke="#1d4f86" stroke-width="3" stroke-linecap="round"/></svg>';
    return `data:image/svg+xml;utf8,${encodeURIComponent(svg)}`;
  },
  listFonts: async () => ["Consolas", "Segoe UI", "굴림", "궁서", "맑은 고딕", "바탕"],
  pickFile: async () => "C:\\Program Files\\Example\\example.exe",
  pickFolder: async () => "D:\\자료\\사진",
  async exportBoard() {
    console.info("[preview] export board");
    return "C:\\backup\\deskboard.json";
  },
  async importBoard() {
    const board: Board = { ...previewBoard, notes: [], settings: { ...previewBoard.settings, theme: "mint" } };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(board));
    return board;
  },
  async setHotkey(hotkey) {
    if (hotkey.includes("F12")) throw `단축키 ${hotkey} 를 등록하지 못했습니다. 다른 프로그램이 쓰고 있을 수 있습니다.`;
    console.info("[preview] hotkey", hotkey || "(none)");
  },
  getAutostart: async () => previewAutostart,
  async setAutostart(enabled) {
    previewAutostart = enabled;
    return previewAutostart;
  },
  async checkUpdate() {
    const version = hooks.update;
    if (!version) return null;
    return { version, install: async () => console.info("[preview] install update", version) };
  },
  async onFileDrop(handler) {
    hooks.drop = handler;
    return () => {
      hooks.drop = undefined;
    };
  },
  startMove: async () => console.info("[preview] start move"),
  startResize: async (edge) => console.info("[preview] start resize", edge),
  setResizable: async () => {},
  resetWindow: async () => console.info("[preview] reset window"),
  quit: async () => console.info("[preview] quit"),
  version: async () => "preview",
};

export const api: Api = "__TAURI_INTERNALS__" in window ? tauriApi : previewApi;
