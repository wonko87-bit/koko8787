import { formatDate, humanSize } from "../types";
import type { FileEntry } from "../types";

export default function HistoryView({ entries }: { entries: FileEntry[] }) {
  const filed = entries
    .filter((e) => e.status === "filed")
    .sort((a, b) => (b.filed_at ?? 0) - (a.filed_at ?? 0));

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
      <h2>이동 완료된 파일 ({filed.length})</h2>
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
        </div>
      ))}
    </div>
  );
}
