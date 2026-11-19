import { memo, useEffect, useRef, useState } from "react";
import type { KeyboardEvent, PointerEvent as ReactPointerEvent, RefObject } from "react";
import { newItem } from "../board";
import { INK_HEX } from "../richText";
import { INK_COLORS, NOTE_COLORS, NOTE_FONT_MAX, NOTE_FONT_MIN, NOTE_MIN_H, NOTE_MIN_W } from "../types";
import type { InkColor, Note, NoteItem } from "../types";
import { ChecklistIcon, CloseIcon, TextIcon } from "./icons";
import { format, RichText } from "./RichText";

const INK_NAMES: Record<InkColor, string> = {
  red: "빨강",
  orange: "주황",
  green: "초록",
  blue: "파랑",
  purple: "보라",
};

/** 서식 단추를 눌러도 편집 칸의 선택과 포커스가 풀리지 않게 한다. */
const keepSelection = (e: { preventDefault: () => void }) => e.preventDefault();

const COLOR_NAMES: Record<Note["color"], string> = {
  yellow: "노랑",
  pink: "분홍",
  green: "연두",
  blue: "하늘",
  purple: "보라",
};

/** 체크리스트. Enter 로 아래에 항목을 만들고, 빈 항목에서 Backspace 로 지운다. */
function Checklist({ note, onChange }: { note: Note; onChange: (id: string, patch: Partial<Note>) => void }) {
  const inputs = useRef(new Map<string, HTMLDivElement>());
  const [focusId, setFocusId] = useState<string | null>(null);
  const items = note.items;

  useEffect(() => {
    if (!focusId) return;
    inputs.current.get(focusId)?.focus();
    setFocusId(null);
  }, [focusId]);

  const setItems = (next: NoteItem[]) => onChange(note.id, { items: next });
  const patch = (id: string, change: Partial<NoteItem>) =>
    setItems(items.map((item) => (item.id === id ? { ...item, ...change } : item)));

  const onKeyDown = (item: NoteItem, index: number) => (e: KeyboardEvent<HTMLDivElement>) => {
    // 한글을 조합하는 중의 Enter 는 글자를 확정하는 것이지 줄을 바꾸는 것이 아니다.
    if (e.nativeEvent.isComposing) return;
    if (e.key === "Enter") {
      e.preventDefault();
      const added = newItem();
      setItems([...items.slice(0, index + 1), added, ...items.slice(index + 1)]);
      setFocusId(added.id);
    } else if (e.key === "Backspace" && item.text === "" && items.length > 1) {
      e.preventDefault();
      setItems(items.filter((other) => other.id !== item.id));
      setFocusId((items[index - 1] ?? items[index + 1]).id);
    } else if (e.key === "ArrowUp" && index > 0) {
      e.preventDefault();
      setFocusId(items[index - 1].id);
    } else if (e.key === "ArrowDown" && index < items.length - 1) {
      e.preventDefault();
      setFocusId(items[index + 1].id);
    }
  };

  return (
    <div className="note-list" style={{ fontSize: note.fontSize }}>
      {items.map((item, index) => (
        <div key={item.id} className={item.done ? "note-item done" : "note-item"}>
          <input
            type="checkbox"
            checked={item.done}
            aria-label={item.done ? "끝냄 표시 지우기" : "끝냄으로 표시"}
            onChange={(e) => patch(item.id, { done: e.target.checked })}
          />
          <RichText
            singleLine
            value={item}
            placeholder="할 일…"
            elementRef={(el) => {
              if (el) inputs.current.set(item.id, el);
              else inputs.current.delete(item.id);
            }}
            onChange={(rich) => patch(item.id, rich)}
            onKeyDown={onKeyDown(item, index)}
          />
        </div>
      ))}
    </div>
  );
}

interface CardProps {
  note: Note;
  focusOnMount: boolean;
  onChange: (id: string, patch: Partial<Note>) => void;
  onSwitchKind: (id: string) => void;
  onRemove: (note: Note) => void;
  onRaise: (id: string) => void;
}

const NoteCard = memo(function NoteCard(props: CardProps) {
  const { note, focusOnMount, onChange, onSwitchKind, onRemove, onRaise } = props;
  // 끄는 동안에는 화면에서만 움직이고, 손을 뗄 때 한 번 저장한다.
  const [live, setLive] = useState<Partial<Note> | null>(null);
  const [paletteOpen, setPaletteOpen] = useState(false);
  const [inkOpen, setInkOpen] = useState(false);
  const textRef = useRef<HTMLDivElement | null>(null);
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

  const isList = note.kind === "list";

  return (
    <div
      className={`note note-${note.color}${live ? " moving" : ""}`}
      style={{ left: shown.x, top: shown.y, width: shown.w, height: shown.h, zIndex: note.z }}
      onPointerDownCapture={() => onRaise(note.id)}
    >
      <div className="note-bar" title="끌어서 옮기기" onPointerDown={beginDrag("move")}>
        <div className="note-tools">
          <button
            className={`note-tool note-swatch note-${note.color}`}
            title="색 바꾸기"
            aria-label="색 바꾸기"
            aria-expanded={paletteOpen}
            onClick={() => setPaletteOpen((open) => !open)}
          />
          <button
            className="note-tool"
            title={isList ? "글 메모로 바꾸기" : "체크리스트로 바꾸기"}
            aria-label={isList ? "글 메모로 바꾸기" : "체크리스트로 바꾸기"}
            onClick={() => onSwitchKind(note.id)}
          >
            {isList ? <TextIcon /> : <ChecklistIcon />}
          </button>
          <button
            className="note-tool note-font"
            title="글자 작게"
            aria-label="글자 작게"
            disabled={note.fontSize <= NOTE_FONT_MIN}
            onClick={() => onChange(note.id, { fontSize: note.fontSize - 1 })}
          >
            A−
          </button>
          <button
            className="note-tool note-font"
            title="글자 크게"
            aria-label="글자 크게"
            disabled={note.fontSize >= NOTE_FONT_MAX}
            onClick={() => onChange(note.id, { fontSize: note.fontSize + 1 })}
          >
            A+
          </button>
          <span className="note-tool-gap" />
          <button className="note-tool note-bold" title="굵게 (선택한 글자)" aria-label="굵게" onMouseDown={keepSelection} onClick={format.bold}>
            B
          </button>
          <button className="note-tool note-italic" title="기울임 (선택한 글자)" aria-label="기울임" onMouseDown={keepSelection} onClick={format.italic}>
            I
          </button>
          <button
            className="note-tool note-ink"
            title="글자 색 (선택한 글자)"
            aria-label="글자 색"
            aria-expanded={inkOpen}
            onMouseDown={keepSelection}
            onClick={() => setInkOpen((open) => !open)}
          >
            A
          </button>
        </div>
        <button className="note-tool note-remove" title="메모 떼기" aria-label="메모 떼기" onClick={() => onRemove(note)}>
          <CloseIcon size={12} />
        </button>
      </div>
      {paletteOpen && (
        <div className="note-palette">
          {NOTE_COLORS.map((color) => (
            <button
              key={color}
              className={`note-color note-${color}${color === note.color ? " current" : ""}`}
              title={`${COLOR_NAMES[color]}색으로`}
              aria-label={`${COLOR_NAMES[color]}색으로`}
              onClick={() => {
                onChange(note.id, { color });
                setPaletteOpen(false);
              }}
            />
          ))}
        </div>
      )}
      {inkOpen && (
        <div className="note-palette">
          <button
            className="ink-color"
            style={{ background: "#2b2a26" }}
            title="기본 글자색으로"
            aria-label="기본 글자색으로"
            onMouseDown={keepSelection}
            onClick={() => {
              format.color("");
              setInkOpen(false);
            }}
          />
          {INK_COLORS.map((color) => (
            <button
              key={color}
              className="ink-color"
              style={{ background: INK_HEX[color] }}
              title={`${INK_NAMES[color]} 글자로`}
              aria-label={`${INK_NAMES[color]} 글자로`}
              onMouseDown={keepSelection}
              onClick={() => {
                format.color(color);
                setInkOpen(false);
              }}
            />
          ))}
        </div>
      )}
      {isList ? (
        <Checklist note={note} onChange={onChange} />
      ) : (
        <RichText
          className="note-text"
          value={note}
          placeholder="메모…"
          style={{ fontSize: note.fontSize }}
          elementRef={(el) => (textRef.current = el)}
          onChange={(rich) => onChange(note.id, rich)}
        />
      )}
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
  onSwitchKind: (id: string) => void;
  onRemove: (note: Note) => void;
  onRaise: (id: string) => void;
}

export function NoteBoard(props: BoardProps) {
  const { notes, focusId, boardRef, onAdd, onChange, onSwitchKind, onRemove, onRaise } = props;
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
          onSwitchKind={onSwitchKind}
          onRemove={onRemove}
          onRaise={onRaise}
        />
      ))}
    </div>
  );
}
