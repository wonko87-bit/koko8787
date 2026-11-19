// 백엔드(Tauri)와 이야기하는 유일한 통로.
// 브라우저에서 `npm run dev` 로만 띄웠을 때는 localStorage 로 흉내 내서 화면 작업을 할 수 있게 한다.

import { getVersion } from "@tauri-apps/api/app";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { open } from "@tauri-apps/plugin-dialog";
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

export interface Api {
  loadBoard(): Promise<Board>;
  saveBoard(board: Board): Promise<void>;
  launch(target: string): Promise<void>;
  classify(target: string): Promise<TargetInfo>;
  pickFile(): Promise<string | null>;
  pickFolder(): Promise<string | null>;
  startMove(): Promise<void>;
  startResize(edge: ResizeEdge): Promise<void>;
  setResizable(resizable: boolean): Promise<void>;
  resetWindow(): Promise<void>;
  quit(): Promise<void>;
  version(): Promise<string>;
}

const tauriApi: Api = {
  loadBoard: () => invoke<Board>("load_board"),
  saveBoard: (board) => invoke("save_board", { board }),
  launch: (target) => invoke("launch_target", { target }),
  classify: (target) => invoke<TargetInfo>("classify_target", { target }),
  pickFile: () => open({ multiple: false, directory: false, title: "바로가기로 만들 파일·프로그램 선택" }),
  pickFolder: () => open({ multiple: false, directory: true, title: "바로가기로 만들 폴더 선택" }),
  startMove: () => getCurrentWindow().startDragging(),
  startResize: (edge) => getCurrentWindow().startResizeDragging(edge),
  setResizable: (resizable) => getCurrentWindow().setResizable(resizable),
  resetWindow: () => invoke("reset_window"),
  quit: () => invoke("quit_app"),
  version: () => getVersion(),
};

const STORAGE_KEY = "deskboard.preview.board";

const previewBoard: Board = {
  version: 1,
  categories: [
    {
      id: "c1",
      name: "자주 가는 폴더",
      tiles: [
        { id: "t1", name: "다운로드", target: "C:\\Users\\me\\Downloads", kind: "folder" },
        { id: "t2", name: "프로젝트", target: "D:\\Projects", kind: "folder" },
        { id: "t3", name: "2026 보고서", target: "D:\\업무\\2026 보고서", kind: "folder" },
      ],
    },
    {
      id: "c2",
      name: "프로그램",
      tiles: [
        { id: "t4", name: "메모장", target: "C:\\Windows\\notepad.exe", kind: "program" },
        { id: "t5", name: "주간 회의록", target: "D:\\업무\\회의록.docx", kind: "file" },
        { id: "t6", name: "사내 포털", target: "https://example.com", kind: "link" },
      ],
    },
  ],
  notes: [
    { id: "n1", text: "오후 3시 회의\n자료 출력하기", x: 16, y: 14, w: 180, h: 130, color: "yellow", z: 1 },
    { id: "n2", text: "택배 찾기", x: 150, y: 120, w: 150, h: 110, color: "pink", z: 2 },
  ],
  settings: { split: 0.56, tileSize: "m", opacity: 0.86, locked: true },
};

const previewApi: Api = {
  async loadBoard() {
    const saved = localStorage.getItem(STORAGE_KEY);
    return saved ? (JSON.parse(saved) as Board) : previewBoard;
  },
  async saveBoard(board) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(board));
  },
  async launch(target) {
    if (target.includes("없는")) throw `대상을 찾을 수 없습니다: ${target}`;
    console.info("[preview] launch", target);
  },
  async classify(target) {
    if (/^[a-z][a-z0-9+.-]+:./i.test(target)) return { kind: "link", exists: true };
    const ext = /\.([a-z0-9-]+)$/i.exec(target)?.[1].toLowerCase();
    const kind = !ext ? "folder" : ["exe", "lnk", "bat", "cmd", "msi"].includes(ext) ? "program" : "file";
    return { kind, exists: !target.includes("없는") };
  },
  pickFile: async () => "C:\\Program Files\\Example\\example.exe",
  pickFolder: async () => "D:\\자료\\사진",
  startMove: async () => console.info("[preview] start move"),
  startResize: async (edge) => console.info("[preview] start resize", edge),
  setResizable: async () => {},
  resetWindow: async () => console.info("[preview] reset window"),
  quit: async () => console.info("[preview] quit"),
  version: async () => "preview",
};

export const api: Api = "__TAURI_INTERNALS__" in window ? tauriApi : previewApi;
