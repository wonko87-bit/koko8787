import { useEffect, useState } from "react";
import { api } from "../api";
import { humanSize } from "../types";
import type { Candidate } from "../types";

export default function CollectModal({ onClose }: { onClose: () => void }) {
  const [candidates, setCandidates] = useState<Candidate[] | null>(null);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    api.listUncollected().then(setCandidates).catch(() => setCandidates([]));
  }, []);

  const toggle = (path: string) => {
    const next = new Set(selected);
    if (next.has(path)) next.delete(path);
    else next.add(path);
    setSelected(next);
  };

  const toggleAll = () => {
    if (!candidates) return;
    if (selected.size === candidates.length) setSelected(new Set());
    else setSelected(new Set(candidates.map((c) => c.path)));
  };

  const collect = async () => {
    setBusy(true);
    try {
      await api.collectPaths([...selected]);
      onClose();
    } catch (e) {
      alert(String(e));
      setBusy(false);
    }
  };

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <header>감시 폴더의 기존 파일 수집</header>
        <div className="modal-body">
          {candidates === null && <div className="hint">불러오는 중…</div>}
          {candidates?.length === 0 && (
            <div className="hint">수집할 파일이 없어요.</div>
          )}
          {candidates && candidates.length > 0 && (
            <>
              <div className="row">
                <input
                  type="checkbox"
                  checked={selected.size === candidates.length}
                  onChange={toggleAll}
                />
                <div className="grow">전체 선택 ({candidates.length}개)</div>
              </div>
              {candidates.map((c) => (
                <div className="row" key={c.path}>
                  <input
                    type="checkbox"
                    checked={selected.has(c.path)}
                    onChange={() => toggle(c.path)}
                  />
                  <div className="grow">
                    <div>{c.name}</div>
                    <div className="path">{c.path}</div>
                  </div>
                  <span className="file-meta">{humanSize(c.size)}</span>
                </div>
              ))}
            </>
          )}
        </div>
        <footer>
          <button onClick={onClose} disabled={busy}>
            닫기
          </button>
          <button
            className="primary"
            onClick={collect}
            disabled={busy || selected.size === 0}
          >
            {selected.size}개 수집
          </button>
        </footer>
      </div>
    </div>
  );
}
