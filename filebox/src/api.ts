import { invoke } from "@tauri-apps/api/core";
import type {
  BatchResult,
  Candidate,
  Favorite,
  FileEntry,
  FlowdeckStatus,
  Rule,
  Settings,
  Suggestion,
} from "./types";

export const api = {
  getSettings: () => invoke<Settings | null>("get_settings"),
  updateSettings: (settings: Settings) =>
    invoke<Settings>("update_settings", { settings }),

  listEntries: () => invoke<FileEntry[]>("list_entries"),
  getSuggestions: (entryId: string) =>
    invoke<Suggestion[]>("get_suggestions", { entryId }),
  sendToFavorite: (entryId: string, favoriteId: string) =>
    invoke<FileEntry>("send_to_favorite", { entryId, favoriteId }),
  sendToPath: (entryId: string, destDir: string) =>
    invoke<FileEntry>("send_to_path", { entryId, destDir }),
  setCategory: (entryId: string, category: string) =>
    invoke<void>("set_category", { entryId, category }),
  setTags: (entryId: string, tags: string[]) =>
    invoke<void>("set_tags", { entryId, tags }),
  removeEntry: (entryId: string) => invoke<void>("remove_entry", { entryId }),
  removeEntries: (entryIds: string[]) =>
    invoke<void>("remove_entries", { entryIds }),
  setCategoryMany: (entryIds: string[], category: string) =>
    invoke<void>("set_category_many", { entryIds, category }),
  sendManyToFavorite: (entryIds: string[], favoriteId: string) =>
    invoke<BatchResult>("send_many_to_favorite", { entryIds, favoriteId }),
  sendManyToPath: (entryIds: string[], destDir: string) =>
    invoke<BatchResult>("send_many_to_path", { entryIds, destDir }),
  undoMove: (entryId: string) => invoke<FileEntry>("undo_move", { entryId }),

  trashEntry: (entryId: string) => invoke<FileEntry>("trash_entry", { entryId }),
  trashMany: (entryIds: string[]) => invoke<BatchResult>("trash_many", { entryIds }),
  trashCount: () => invoke<number>("trash_count"),
  emptyTrash: () => invoke<number>("empty_trash"),
  openTrash: () => invoke<void>("open_trash"),
  clearHistory: () => invoke<void>("clear_history"),

  listFavorites: () => invoke<Favorite[]>("list_favorites"),
  addFavorite: (name: string, path: string) =>
    invoke<Favorite>("add_favorite", { name, path }),
  removeFavorite: (favoriteId: string) =>
    invoke<void>("remove_favorite", { favoriteId }),

  listRules: () => invoke<Rule[]>("list_rules"),
  upsertRule: (rule: Rule) => invoke<Rule>("upsert_rule", { rule }),
  removeRule: (ruleId: string) => invoke<void>("remove_rule", { ruleId }),

  flowdeckStatus: () => invoke<FlowdeckStatus>("flowdeck_status"),
  sendToFlowdeck: (entryId: string) =>
    invoke<string>("send_to_flowdeck", { entryId }),
  setFlowdeckPending: (entryId: string, pending: boolean) =>
    invoke<FileEntry>("set_flowdeck_pending", { entryId, pending }),
  clearRecent: (entryIds: string[]) =>
    invoke<void>("clear_recent", { entryIds }),

  listUncollected: () => invoke<Candidate[]>("list_uncollected"),
  collectPaths: (paths: string[]) => invoke<number>("collect_paths", { paths }),

  openEntry: (entryId: string) => invoke<void>("open_entry", { entryId }),
  openFolder: (path: string) => invoke<void>("open_folder", { path }),
  revealEntry: (entryId: string) => invoke<void>("reveal_entry", { entryId }),

  hideToast: () => invoke<void>("hide_toast"),
  showMainWindow: () => invoke<void>("show_main_window"),
};
