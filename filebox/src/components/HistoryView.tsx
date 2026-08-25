import { confirm } from "@tauri-apps/plugin-dialog";
import { api } from "../api";
import { formatDate, humanSize } from "../types";
import type { FileEntry } from "../types";

export default function HistoryView({ entries }: { entries: FileEntry[] }) {
  const done = entries
    .filter((e) => e.status === "filed" || e.status === "trashed")
    .sort((a, b) => (b.filed_at ?? 0) - (a.filed_at ?? 0));

  const clearAll = async () => {
    const ok = await confirm(
      `기록 ${done.length}개를 모두 삭제할까요?\n(파일은 그대로 있고, 추천 학습 데이터도 유지됩니다)`,
      { title: "기록 전체 삭제", kind: "warning" },
    );
    if (ok) await api.clearHistory();
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
      <div style={{ display: "flex", alignItems: "center", marginBottom: 12 }}>
        <h2 style={{ margin: 0, flex: 1 }}>처리한 파일 ({done.length})</h2>
        <button className="ghost" onClick={clearAll}>
          전체 삭제
        </button>
      </div>
      {done.map((e) => (
        <div className="row" key={e.id}>
          <div className="grow">
            <div>{e.file_name}</div>
            <div className="path">
              {e.status === "trashed" ? "🗑 윈도우 휴지통으로 보냄" : `→ ${e.filed_to}`}
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
            <button
              title="파일을 관리함으로 되돌리고 이 이동의 학습 기록도 취소"
              onClick={() => api.undoMove(e.id).catch((err) => alert(String(err)))}
            >
              되돌리기
            </button>
          )}
          <button
            className="ghost"
            title="이 기록만 삭제 (파일은 그대로)"
            onClick={() => api.removeEntry(e.id)}
          >
            삭제
          </button>
        </div>
      ))}
    </div>
  );
}
