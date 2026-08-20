import { confirm } from "@tauri-apps/plugin-dialog";
import { api } from "../api";
import { formatDate, humanSize } from "../types";
import type { FileEntry } from "../types";

export default function HistoryView({ entries }: { entries: FileEntry[] }) {
  const filed = entries
    .filter((e) => e.status === "filed")
    .sort((a, b) => (b.filed_at ?? 0) - (a.filed_at ?? 0));

  const clearAll = async () => {
    const ok = await confirm(
      `이동 기록 ${filed.length}개를 모두 삭제할까요?\n(파일은 이동된 폴더에 그대로 남고, 추천 학습 데이터도 유지됩니다)`,
      { title: "기록 전체 삭제", kind: "warning" },
    );
    if (ok) await api.clearHistory();
  };

  if (filed.length === 0) {
    return (
      <div className="empty">
        <div className="big">🗂️</div>
        아직 이동 기록이 없어요.
      </div>
    );
  }

  return (
    <div className="panel">
      <div style={{ display: "flex", alignItems: "center", marginBottom: 12 }}>
        <h2 style={{ margin: 0, flex: 1 }}>이동 완료된 파일 ({filed.length})</h2>
        <button className="ghost" onClick={clearAll}>
          전체 삭제
        </button>
      </div>
      {filed.map((e) => (
        <div className="row" key={e.id}>
          <div className="grow">
            <div>{e.file_name}</div>
            <div className="path">→ {e.filed_to}</div>
          </div>
          <span className="file-meta">
            {humanSize(e.size)}
            {e.filed_at ? ` · ${formatDate(e.filed_at)}` : ""}
          </span>
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
