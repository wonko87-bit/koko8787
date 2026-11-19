import { memo, useEffect, useRef, useState } from "react";
import type { PointerEvent as ReactPointerEvent, RefObject } from "react";
import { NOTE_COLORS, NOTE_MIN_H, NOTE_MIN_W } from "../types";
import type { Note } from "../types";
import { CloseIcon } from "./icons";

const COLOR_NAMES: Record<Note["color"], string> = {
  yellow: "노랑",
  pink: "분홍",
  green: "연두",
  blue: "하늘",
  purple: "보라",
};

interface CardProps {
  note: Note;
  focusOnMount: boolean;
  onChange: (id: string, patch: Partial<Note>) => void;
  onRemove: (note: Note) => void;
  onRaise: (id: string) => void;
}

const NoteCard = memo(function NoteCard({ note, focusOnMount, onChange, onRemove, onRaise }: CardProps) {
  // 끄는 동안에는 화면에서만 움직이고, 손을 뗄 때 한 번 저장한다.
  const [live, setLive] = useState<Partial<Note> | null>(null);
  const textRef = useRef<HTMLTextAreaElement>(null);
  const shown = live ? { ...note, ...live } : note;

  useEffect(() => {
    if (focusOnMount) textRef.current?.focus();
  }, [focusOnMount]);

  const beginDrag = (mode: "move" | "size") => (e: ReactPointerEvent<HTMLElement>) => {
    if (e.button !== 0 || (e.target as HTMLElement).closest("button")) return;
    e.preventDefault();
    const handle = e.currentTarget;
    handle.setPointerCapture(e.pointerId);
    const start = { px: e.clientX, py: e.clientY };
    let latest: Partial<Note> | null = null;

    const move = (ev: PointerEvent) => {
      const dx = ev.clientX - start.px;
      const dy = ev.clientY - start.py;
      latest =
        mode === "move"
          ? { x: Math.max(0, note.x + dx), y: Math.max(0, note.y + dy) }
          : { w: Math.max(NOTE_MIN_W, note.w + dx), h: Math.max(NOTE_MIN_H, note.h + dy) };
      setLive(latest);
    };
    const end = () => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", end);
      handle.removeEventListener("pointercancel", end);
      if (latest) onChange(note.id, latest);
      setLive(null);
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", end);
    handle.addEventListener("pointercancel", end);
  };

  return (
    <div
      className={`note note-${note.color}${live ? " moving" : ""}`}
      style={{ left: shown.x, top: shown.y, width: shown.w, height: shown.h, zIndex: note.z }}
      onPointerDownCapture={() => onRaise(note.id)}
    >
      <div className="note-bar" title="끌어서 옮기기" onPointerDown={beginDrag("move")}>
        <div className="note-colors">
          {NOTE_COLORS.map((color) => (
            <button
              key={color}
              className={`note-color note-${color}${color === note.color ? " current" : ""}`}
              title={`${COLOR_NAMES[color]}색으로`}
              aria-label={`${COLOR_NAMES[color]}색으로`}
              onClick={() => onChange(note.id, { color })}
            />
          ))}
        </div>
        <button className="note-remove" title="메모 떼기" aria-label="메모 떼기" onClick={() => onRemove(note)}>
          <CloseIcon size={12} />
        </button>
      </div>
      <textarea
        ref={textRef}
        value={note.text}
        placeholder="메모…"
        spellCheck={false}
        onChange={(e) => onChange(note.id, { text: e.target.value })}
      />
      <div className="note-grip" title="끌어서 크기 조절" onPointerDown={beginDrag("size")} />
    </div>
  );
});

interface BoardProps {
  notes: Note[];
  /** 방금 만든 메모. 바로 글을 쓸 수 있게 커서를 넣는다. */
  focusId: string | null;
  boardRef: RefObject<HTMLDivElement>;
  onAdd: (at: { x: number; y: number }) => void;
  onChange: (id: string, patch: Partial<Note>) => void;
  onRemove: (note: Note) => void;
  onRaise: (id: string) => void;
}

export function NoteBoard({ notes, focusId, boardRef, onAdd, onChange, onRemove, onRaise }: BoardProps) {
  return (
    <div
      ref={boardRef}
      className="notes"
      onDoubleClick={(e) => {
        if (e.target !== e.currentTarget) return;
        const el = e.currentTarget;
        const box = el.getBoundingClientRect();
        onAdd({ x: e.clientX - box.left + el.scrollLeft - 24, y: e.clientY - box.top + el.scrollTop - 12 });
      }}
    >
      {notes.length === 0 && <p className="notes-empty">빈 곳을 더블클릭하면 메모가 붙습니다</p>}
      {notes.map((note) => (
        <NoteCard
          key={note.id}
          note={note}
          focusOnMount={note.id === focusId}
          onChange={onChange}
          onRemove={onRemove}
          onRaise={onRaise}
        />
      ))}
    </div>
  );
}
