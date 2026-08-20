import { useEffect, useMemo, useState } from "react";
import { open } from "@tauri-apps/plugin-dialog";
import { api } from "../api";
import {
  categoryIcon,
  DEFAULT_CATEGORIES,
  formatDate,
  humanSize,
} from "../types";
import type { FileEntry, Favorite, Suggestion } from "../types";

function FileCard({
  entry,
  favorites,
}: {
  entry: FileEntry;
  favorites: Favorite[];
}) {
  const [suggestions, setSuggestions] = useState<Suggestion[]>([]);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    api.getSuggestions(entry.id).then(setSuggestions).catch(() => {});
  }, [entry.id, entry.file_name]);

  const run = async (fn: () => Promise<unknown>) => {
    setBusy(true);
    try {
      await fn();
    } catch (e) {
      alert(String(e));
    } finally {
      setBusy(false);
    }
  };

  const sendToOther = () =>
    run(async () => {
      const dir = await open({ directory: true, title: "이동할 폴더 선택" });
      if (typeof dir === "string") await api.sendToPath(entry.id, dir);
    });

  const categories = useMemo(() => {
    const set = new Set(DEFAULT_CATEGORIES);
    set.add(entry.category);
    return [...set];
  }, [entry.category]);

  const suggested = new Set(suggestions.map((s) => s.favorite.id));
  const otherFavorites = favorites.filter((f) => !suggested.has(f.id));

  return (
    <div className="file-card">
      <div className="file-row">
        <span className="file-icon">{categoryIcon(entry.category)}</span>
        <span className="file-name grow">{entry.file_name}</span>
        <span className="file-meta">
          {humanSize(entry.size)} · {formatDate(entry.added_at)}
        </span>
      </div>
      <div className="file-actions">
        {suggestions.map((s) => (
          <button
            key={s.favorite.id}
            className="suggestion-btn"
            disabled={busy}
            title={s.favorite.path}
            onClick={() => run(() => api.sendToFavorite(entry.id, s.favorite.id))}
          >
            ⭐ {s.favorite.name}
          </button>
        ))}
        {otherFavorites.length > 0 && (
          <select
            disabled={busy}
            value=""
            onChange={(e) => {
              const id = e.target.value;
              if (id) run(() => api.sendToFavorite(entry.id, id));
            }}
          >
            <option value="">즐겨찾기로 이동…</option>
            {otherFavorites.map((f) => (
              <option key={f.id} value={f.id}>
                {f.name}
              </option>
            ))}
          </select>
        )}
        <button disabled={busy} onClick={sendToOther}>
          다른 폴더로…
        </button>
        <span className="sep" />
        <select
          disabled={busy}
          value={entry.category}
          title="카테고리 변경"
          onChange={(e) => run(() => api.setCategory(entry.id, e.target.value))}
        >
          {categories.map((c) => (
            <option key={c} value={c}>
              {c}
            </option>
          ))}
        </select>
        <button disabled={busy} onClick={() => api.openEntry(entry.id)}>
          열기
        </button>
        <button disabled={busy} onClick={() => api.revealEntry(entry.id)}>
          위치 보기
        </button>
        <button
          className="ghost"
          disabled={busy}
          title="목록에서만 제거 (파일은 관리함 폴더에 남음)"
          onClick={() => run(() => api.removeEntry(entry.id))}
        >
          제거
        </button>
      </div>
    </div>
  );
}

export default function InboxView({
  entries,
  favorites,
}: {
  entries: FileEntry[];
  favorites: Favorite[];
}) {
  const inbox = entries.filter((e) => e.status === "inbox");

  const groups = useMemo(() => {
    const map = new Map<string, FileEntry[]>();
    for (const e of [...inbox].sort((a, b) => b.added_at - a.added_at)) {
      const list = map.get(e.category) ?? [];
      list.push(e);
      map.set(e.category, list);
    }
    return [...map.entries()];
  }, [inbox]);

  if (inbox.length === 0) {
    return (
      <div className="empty">
        <div className="big">📭</div>
        관리함이 비어 있어요.
        <br />
        감시 폴더에 새 파일이 생기면 자동으로 수집됩니다.
      </div>
    );
  }

  return (
    <>
      {groups.map(([category, list]) => (
        <div className="category-group" key={category}>
          <h2>
            {categoryIcon(category)} {category} ({list.length})
          </h2>
          {list.map((e) => (
            <FileCard key={e.id} entry={e} favorites={favorites} />
          ))}
        </div>
      ))}
    </>
  );
}
