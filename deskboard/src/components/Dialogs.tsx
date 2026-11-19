import { useEffect, useRef, useState } from "react";
import type { FormEvent, ReactNode } from "react";
import { api } from "../api";
import type { TargetInfo } from "../api";
import { newId } from "../board";
import type { Category, Tile } from "../types";

function Modal({ title, onClose, children }: { title: string; onClose: () => void; children: ReactNode }) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && onClose();
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);

  return (
    <div className="overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="dialog" role="dialog" aria-modal="true" aria-label={title}>
        <h2>{title}</h2>
        {children}
      </div>
    </div>
  );
}

/** 경로나 주소에서 타일에 붙일 만한 이름을 뽑는다. */
export function suggestName(target: string): string {
  const trimmed = target.trim();
  const url = /^[a-z][a-z0-9+.-]+:\/\/([^/?#]+)/i.exec(trimmed);
  if (url) return url[1].replace(/^www\./, "");
  const last = trimmed.replace(/[\\/]+$/, "").split(/[\\/]/).pop() ?? "";
  return last.replace(/\.(exe|lnk|bat|cmd|url)$/i, "") || trimmed;
}

interface TileDialogProps {
  categories: Category[];
  /** 고칠 타일. 없으면 새로 만든다. */
  tile?: Tile;
  categoryId: string;
  onSave: (categoryId: string, tile: Tile) => void;
  onClose: () => void;
}

export function TileDialog({ categories, tile, categoryId, onSave, onClose }: TileDialogProps) {
  const [name, setName] = useState(tile?.name ?? "");
  const [target, setTarget] = useState(tile?.target ?? "");
  const [category, setCategory] = useState(categoryId);
  const [info, setInfo] = useState<TargetInfo | null>(null);
  const [saving, setSaving] = useState(false);
  const targetRef = useRef<HTMLInputElement>(null);

  useEffect(() => targetRef.current?.focus(), []);

  // 입력이 잠깐 멈추면 대상이 실제로 있는지 확인해 알려 준다.
  useEffect(() => {
    setInfo(null);
    const value = target.trim();
    if (!value) return;
    let stale = false;
    const timer = setTimeout(() => {
      api.classify(value).then((result) => !stale && setInfo(result), () => {});
    }, 250);
    return () => {
      stale = true;
      clearTimeout(timer);
    };
  }, [target]);

  const pick = async (choose: () => Promise<string | null>) => {
    const picked = await choose().catch(() => null);
    if (!picked) return;
    setTarget(picked);
    if (!name.trim()) setName(suggestName(picked));
  };

  const submit = async (e: FormEvent) => {
    e.preventDefault();
    const value = target.trim();
    if (!value || saving) return;
    setSaving(true);
    const kind = await api.classify(value).then((r) => r.kind, () => tile?.kind ?? "file");
    onSave(category, { id: tile?.id ?? newId(), name: name.trim() || suggestName(value), target: value, kind });
  };

  return (
    <Modal title={tile ? "바로가기 편집" : "바로가기 추가"} onClose={onClose}>
      <form onSubmit={submit}>
        <label>
          <span>대상</span>
          <input
            ref={targetRef}
            value={target}
            onChange={(e) => setTarget(e.target.value)}
            placeholder="폴더·파일 경로 또는 https:// 주소"
            spellCheck={false}
          />
        </label>
        <div className="dialog-pickers">
          <button type="button" onClick={() => pick(api.pickFolder)}>
            폴더 선택…
          </button>
          <button type="button" onClick={() => pick(api.pickFile)}>
            파일·프로그램 선택…
          </button>
        </div>
        {info && !info.exists && (
          <p className="dialog-warning">이 경로를 찾을 수 없습니다. 그대로 저장할 수는 있습니다.</p>
        )}
        <label>
          <span>이름</span>
          <input
            value={name}
            onChange={(e) => setName(e.target.value)}
            placeholder={suggestName(target) || "타일에 표시할 이름"}
          />
        </label>
        <label>
          <span>카테고리</span>
          <select value={category} onChange={(e) => setCategory(e.target.value)}>
            {categories.map((c) => (
              <option key={c.id} value={c.id}>
                {c.name}
              </option>
            ))}
          </select>
        </label>
        <div className="dialog-actions">
          <button type="button" onClick={onClose}>
            취소
          </button>
          <button type="submit" className="primary" disabled={!target.trim() || saving}>
            {tile ? "저장" : "추가"}
          </button>
        </div>
      </form>
    </Modal>
  );
}

interface NameDialogProps {
  title: string;
  initial: string;
  confirmLabel: string;
  onSubmit: (name: string) => void;
  onClose: () => void;
}

export function NameDialog({ title, initial, confirmLabel, onSubmit, onClose }: NameDialogProps) {
  const [name, setName] = useState(initial);
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => ref.current?.select(), []);

  return (
    <Modal title={title} onClose={onClose}>
      <form
        onSubmit={(e) => {
          e.preventDefault();
          if (name.trim()) onSubmit(name.trim());
        }}
      >
        <label>
          <span>이름</span>
          <input ref={ref} value={name} onChange={(e) => setName(e.target.value)} maxLength={40} />
        </label>
        <div className="dialog-actions">
          <button type="button" onClick={onClose}>
            취소
          </button>
          <button type="submit" className="primary" disabled={!name.trim()}>
            {confirmLabel}
          </button>
        </div>
      </form>
    </Modal>
  );
}

interface ConfirmDialogProps {
  title: string;
  message: string;
  confirmLabel: string;
  onConfirm: () => void;
  onClose: () => void;
}

export function ConfirmDialog({ title, message, confirmLabel, onConfirm, onClose }: ConfirmDialogProps) {
  const ref = useRef<HTMLButtonElement>(null);
  useEffect(() => ref.current?.focus(), []);

  return (
    <Modal title={title} onClose={onClose}>
      <p className="dialog-message">{message}</p>
      <div className="dialog-actions">
        <button type="button" onClick={onClose}>
          취소
        </button>
        <button ref={ref} type="button" className="primary danger" onClick={onConfirm}>
          {confirmLabel}
        </button>
      </div>
    </Modal>
  );
}
