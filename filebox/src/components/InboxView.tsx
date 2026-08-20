import { useEffect, useMemo, useState } from "react";
import { open } from "@tauri-apps/plugin-dialog";
import { api } from "../api";
import {
  categoryIcon,
  DEFAULT_CATEGORIES,
  formatDate,
  humanSize,
} from "../types";
import type { BatchResult, FileEntry, Favorite, Suggestion } from "../types";

function reportBatch(result: BatchResult) {
  if (result.errors.length > 0) {
    alert(
      `${result.moved}개 이동 완료, ${result.errors.length}개 실패\n\n` +
        result.errors.slice(0, 10).join("\n"),
    );
  }
}

function TagEditor({ entry }: { entry: FileEntry }) {
  const [adding, setAdding] = useState(false);
  const [text, setText] = useState("");

  const commit = () => {
    const tag = text.trim();
    setText("");
    setAdding(false);
    if (tag && !entry.tags.includes(tag)) {
      api.setTags(entry.id, [...entry.tags, tag]).catch((e) => alert(String(e)));
    }
  };

  const removeTag = (tag: string) =>
    api
      .setTags(
        entry.id,
        entry.tags.filter((t) => t !== tag),
      )
      .catch((e) => alert(String(e)));

  return (
    <div className="tag-row">
      {entry.tags.map((t) => (
        <span className="tag" key={t}>
          #{t}
          <button
            className="tag-x"
            title="태그 제거"
            onClick={() => removeTag(t)}
          >
            ×
          </button>
        </span>
      ))}
      {adding ? (
        <input
          className="tag-input"
          autoFocus
          value={text}
          placeholder="태그 입력 후 Enter"
          onChange={(e) => setText(e.target.value)}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === "Enter") commit();
            if (e.key === "Escape") {
              setText("");
              setAdding(false);
            }
          }}
        />
      ) : (
        <button className="tag-add" onClick={() => setAdding(true)}>
          + 태그
        </button>
      )}
    </div>
  );
}

function FileCard({
  entry,
  favorites,
  categories,
  selected,
  onToggle,
}: {
  entry: FileEntry;
  favorites: Favorite[];
  categories: string[];
  selected: boolean;
  onToggle: () => void;
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

  const suggested = new Set(suggestions.map((s) => s.favorite.id));
  const otherFavorites = favorites.filter((f) => !suggested.has(f.id));

  return (
    <div className={`file-card${selected ? " selected" : ""}`}>
      <div className="file-row">
        <input type="checkbox" checked={selected} onChange={onToggle} />
        <span className="file-icon">{categoryIcon(entry.category)}</span>
        <span className="file-name grow">{entry.file_name}</span>
        <span className="file-meta">
          {humanSize(entry.size)} · {formatDate(entry.added_at)}
        </span>
      </div>
      <TagEditor entry={entry} />
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
  const [query, setQuery] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("");
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);

  const inbox = useMemo(
    () =>
      entries
        .filter((e) => e.status === "inbox")
        .sort((a, b) => b.added_at - a.added_at),
    [entries],
  );

  const allCategories = useMemo(
    () => [...new Set([...DEFAULT_CATEGORIES, ...inbox.map((e) => e.category)])],
    [inbox],
  );

  const visible = useMemo(() => {
    const q = query.trim().toLowerCase();
    return inbox.filter((e) => {
      if (categoryFilter && e.category !== categoryFilter) return false;
      if (!q) return true;
      return (
        e.file_name.toLowerCase().includes(q) ||
        e.tags.some((t) => t.toLowerCase().includes(q))
      );
    });
  }, [inbox, query, categoryFilter]);

  // 화면에서 사라진 항목은 선택에서도 제외해, 보이지 않는 파일이 일괄 처리되지 않게 한다.
  useEffect(() => {
    setSelected((prev) => {
      const visibleIds = new Set(visible.map((e) => e.id));
      const next = new Set([...prev].filter((id) => visibleIds.has(id)));
      return next.size === prev.size ? prev : next;
    });
  }, [visible]);

  const groups = useMemo(() => {
    const map = new Map<string, FileEntry[]>();
    for (const e of visible) {
      const list = map.get(e.category) ?? [];
      list.push(e);
      map.set(e.category, list);
    }
    return [...map.entries()];
  }, [visible]);

  const toggle = (id: string) =>
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

  const toggleAll = () =>
    setSelected((prev) =>
      prev.size === visible.length ? new Set() : new Set(visible.map((e) => e.id)),
    );

  const ids = [...selected];
  const runBatch = async (fn: () => Promise<unknown>) => {
    setBusy(true);
    try {
      await fn();
      setSelected(new Set());
    } catch (e) {
      alert(String(e));
    } finally {
      setBusy(false);
    }
  };

  const batchToOther = () =>
    runBatch(async () => {
      const dir = await open({ directory: true, title: "이동할 폴더 선택" });
      if (typeof dir === "string") reportBatch(await api.sendManyToPath(ids, dir));
    });

  if (inbox.length === 0) {
    return (
      <div className="empty">
        <div className="big">📭</div>
        관리함이 비어 있어요.
        <br />
        감시 폴더에 새 파일이 생기거나, 창에 파일을 끌어다 놓으면 수집됩니다.
      </div>
    );
  }

  return (
    <>
      <div className="toolbar">
        <input
          className="search"
          value={query}
          placeholder="파일명 · 태그 검색"
          onChange={(e) => setQuery(e.target.value)}
        />
        <select
          value={categoryFilter}
          onChange={(e) => setCategoryFilter(e.target.value)}
        >
          <option value="">모든 카테고리</option>
          {allCategories.map((c) => (
            <option key={c} value={c}>
              {c}
            </option>
          ))}
        </select>
        <button onClick={toggleAll} disabled={visible.length === 0}>
          {selected.size === visible.length && visible.length > 0
            ? "선택 해제"
            : `전체 선택 (${visible.length})`}
        </button>
        {(query || categoryFilter) && (
          <span className="file-meta">
            {visible.length} / {inbox.length}개 표시
          </span>
        )}
      </div>

      {selected.size > 0 && (
        <div className="batch-bar">
          <strong>{selected.size}개 선택됨</strong>
          <select
            disabled={busy}
            value=""
            onChange={(e) => {
              const favId = e.target.value;
              if (favId)
                runBatch(async () =>
                  reportBatch(await api.sendManyToFavorite(ids, favId)),
                );
            }}
          >
            <option value="">즐겨찾기로 일괄 이동…</option>
            {favorites.map((f) => (
              <option key={f.id} value={f.id}>
                {f.name}
              </option>
            ))}
          </select>
          <button disabled={busy} onClick={batchToOther}>
            다른 폴더로…
          </button>
          <select
            disabled={busy}
            value=""
            onChange={(e) => {
              const cat = e.target.value;
              if (cat) runBatch(() => api.setCategoryMany(ids, cat));
            }}
          >
            <option value="">카테고리 일괄 변경…</option>
            {allCategories.map((c) => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </select>
          <button
            className="ghost"
            disabled={busy}
            title="목록에서만 제거 (파일은 관리함 폴더에 남음)"
            onClick={() => runBatch(() => api.removeEntries(ids))}
          >
            목록에서 제거
          </button>
          <div className="grow" />
          <button className="ghost" onClick={() => setSelected(new Set())}>
            선택 해제
          </button>
        </div>
      )}

      {visible.length === 0 ? (
        <div className="empty">
          <div className="big">🔍</div>
          조건에 맞는 파일이 없어요.
        </div>
      ) : (
        groups.map(([category, list]) => (
          <div className="category-group" key={category}>
            <h2>
              {categoryIcon(category)} {category} ({list.length})
            </h2>
            {list.map((e) => (
              <FileCard
                key={e.id}
                entry={e}
                favorites={favorites}
                categories={allCategories}
                selected={selected.has(e.id)}
                onToggle={() => toggle(e.id)}
              />
            ))}
          </div>
        ))
      )}
    </>
  );
}
