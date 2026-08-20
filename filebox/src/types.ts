export type EntryStatus = "inbox" | "filed";

export interface FileEntry {
  id: string;
  file_name: string;
  path: string;
  origin: string;
  size: number;
  added_at: number;
  category: string;
  tags: string[];
  status: EntryStatus;
  filed_to: string | null;
  filed_at: number | null;
  record_id: string | null;
}

export interface BatchResult {
  moved: number;
  errors: string[];
}

export interface Favorite {
  id: string;
  name: string;
  path: string;
}

export interface Rule {
  id: string;
  name: string;
  extensions: string[];
  keywords: string[];
  category: string | null;
  favorite_id: string | null;
}

export interface Settings {
  watch_dirs: string[];
  inbox_dir: string;
  auto_collect: boolean;
  toast_enabled: boolean;
  paused: boolean;
}

export interface Suggestion {
  favorite: Favorite;
  score: number;
}

export interface Candidate {
  path: string;
  name: string;
  size: number;
}

export const DEFAULT_CATEGORIES = [
  "문서",
  "이미지",
  "영상",
  "음악",
  "압축",
  "설치파일",
  "코드/스크립트",
  "기타",
];

export function humanSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  const units = ["KB", "MB", "GB", "TB"];
  let v = bytes / 1024;
  let i = 0;
  while (v >= 1024 && i < units.length - 1) {
    v /= 1024;
    i++;
  }
  return `${v.toFixed(v >= 100 ? 0 : 1)} ${units[i]}`;
}

export function formatDate(ms: number): string {
  const d = new Date(ms);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

export function categoryIcon(category: string): string {
  switch (category) {
    case "문서":
      return "📄";
    case "이미지":
      return "🖼️";
    case "영상":
      return "🎬";
    case "음악":
      return "🎵";
    case "압축":
      return "🗜️";
    case "설치파일":
      return "💿";
    case "코드/스크립트":
      return "🧩";
    default:
      return "📦";
  }
}
