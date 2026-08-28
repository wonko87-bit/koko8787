import { useMemo, useState } from "react";
import { confirm } from "@tauri-apps/plugin-dialog";
import { api } from "../api";
import { openFolder } from "../openFolder";
import { formatDate, humanSize } from "../types";
import type { FileEntry } from "../types";

export default function HistoryView({ entries }: { entries: FileEntry[] }) {
  const [query, setQuery] = useState("");
  const [busy, setBusy] = useState("");

  const done = useMemo(
    () =>
      entries
        .filter((e) => e.status === "filed" || e.status === "trashed")
        .sort((a, b) => (b.filed_at ?? 0) - (a.filed_at ?? 0)),
    [entries],
  );

  // 할일의 경로가 어긋났을 때 파일이 어디로 갔는지 여기서 찾게 된다.
  // 그래서 파일명뿐 아니라 옮긴 폴더도 검색 대상에 넣는다.
  const visible = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return done;
    return done.filter(
      (e) =>
        e.file_name.toLowerCase().includes(q) ||
        (e.filed_to ?? "").toLowerCase().includes(q) ||
        e.tags.some((t) => t.toLowerCase().includes(q)),
    );
  }, [done, query]);

  const clearAll = async () => {
    const ok = await confirm(
      `기록 ${done.length}개를 모두 삭제할까요?\n(파일은 그대로 있고, 추천 학습 데이터도 유지됩니다)`,
      { title: "기록 전체 삭제", kind: "warning" },
    );
    if (ok) await api.clearHistory();
  };

  const sendToFlowdeck = async (entry: FileEntry) => {
    setBusy(entry.id);
    try {
      await api.sendToFlowdeck(entry.id);
    } catch (e) {
      alert(String(e));
    } finally {
      setBusy("");
    }
  };

  if (done.length === 0) {
    return (
      <div className="empty">
        <div className="big">🗂️</div>
        아직 처리 기록이 없어요.
      </div>
    );
  }

  return (
    <div className="panel">
      <div className="toolbar">
        <input
          className="search"
          value={query}
          placeholder="파일명 · 옮긴 폴더 · 태그 검색"
          onChange={(e) => setQuery(e.target.value)}
        />
        {query && (
          <span className="file-meta">
            {visible.length} / {done.length}개
          </span>
        )}
        <button className="ghost" onClick={clearAll}>
          전체 삭제
        </button>
      </div>

      {visible.length === 0 ? (
        <div className="empty">
          <div className="big">🔍</div>
          조건에 맞는 기록이 없어요.
        </div>
      ) : (
        visible.map((e) => {
          const registered = e.flowdeck_todos.length > 0;
          return (
            <div className="row" key={e.id}>
              <div className="grow">
                <div>
                  {e.file_name}
                  {registered && <span className="flowdeck-chip">📋 등록됨</span>}
                </div>
                <div className="path">
                  {e.status === "trashed"
                    ? "🗑 윈도우 휴지통으로 보냄"
                    : `→ ${e.filed_to}`}
                </div>
              </div>
              <span className="file-meta">
                {humanSize(e.size)}
                {e.filed_at ? ` · ${formatDate(e.filed_at)}` : ""}
              </span>
              {e.status === "trashed" ? (
                <button
                  title="윈도우 휴지통을 열어 직접 복구할 수 있어요"
                  onClick={() => api.openTrash().catch((err) => alert(String(err)))}
                >
                  휴지통 열기
                </button>
              ) : (
                <>
                  <button
                    className={registered ? "flowdeck-sent" : ""}
                    disabled={busy === e.id}
                    title={
                      registered
                        ? "이미 등록했어요. 다시 눌러도 할일이 늘어나지 않습니다"
                        : "이 파일을 Flowdeck 할일로 등록합니다"
                    }
                    onClick={() => sendToFlowdeck(e)}
                  >
                    📋
                  </button>
                  {e.filed_to && (
                    <button
                      title="옮긴 폴더 열기"
                      onClick={() => openFolder(e.filed_to!)}
                    >
                      열기
                    </button>
                  )}
                  <button
                    title="파일을 관리함으로 되돌리고 이 이동의 학습 기록도 취소"
                    onClick={() =>
                      api.undoMove(e.id).catch((err) => alert(String(err)))
                    }
                  >
                    되돌리기
                  </button>
                </>
              )}
              <button
                className="ghost"
                title="이 기록만 삭제 (파일은 그대로)"
                onClick={() => api.removeEntry(e.id)}
              >
                삭제
              </button>
            </div>
          );
        })
      )}
    </div>
  );
}
