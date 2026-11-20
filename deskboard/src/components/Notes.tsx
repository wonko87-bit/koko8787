import { memo, useEffect, useRef, useState } from "react";
import type { KeyboardEvent, PointerEvent as ReactPointerEvent, RefObject } from "react";
import { newBox, newItem, snap } from "../board";
import { INK_HEX } from "../richText";
import { BOX_MIN_W, INK_COLORS, NOTE_COLORS, NOTE_FONT_MAX, NOTE_FONT_MIN, NOTE_MIN_H, NOTE_MIN_W } from "../types";
import type { InkColor, Note, NoteItem, Page, TextBox } from "../types";
import { PlusIcon } from "./icons";
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

/** 글상자가 글을 쓸 수 있으려면 오른쪽에 최소한 이만큼의 폭은 있어야 한다. */
const BOX_MIN_ROOM = 90;
/** 다른 글상자의 왼쪽 끝·윗줄과 이만큼 안으로 가까우면 거기에 맞춰 붙인다. */
const SNAP_WITHIN = 6;

/** 메모의 지금 내용을 보고 바꿀 것을 정한다. 연달아 일어나는 변경(글상자를 비우고 바로 다른 곳을 클릭 등)이 서로를 덮어쓰지 않게 한다. */
type UpdateNote = (id: string, change: (note: Note) => Partial<Note>) => void;

/**
 * 글 메모의 내용. 메모지 안 빈 곳을 클릭하면 그 자리에 글상자가 생기고,
 * 아무것도 적지 않은(공백뿐인) 글상자는 다른 곳을 누르는 순간 사라진다.
 */
function NoteCanvas({ note, focusFirst, onUpdate }: { note: Note; focusFirst: boolean; onUpdate: UpdateNote }) {
  const editors = useRef(new Map<string, HTMLDivElement>());
  const [focusId, setFocusId] = useState<string | null>(focusFirst ? (note.boxes[0]?.id ?? null) : null);
  // 끄는 동안에는 화면에서만 움직이고, 손을 뗄 때 한 번 저장한다.
  const [moving, setMoving] = useState<{ id: string; x: number; y: number } | null>(null);
  const [sizing, setSizing] = useState<{ id: string; w: number } | null>(null);
  /** 줄을 맞추고 있는 자리. 끄는 동안 가는 선으로 보여 준다. */
  const [guides, setGuides] = useState<{ x: number | null; y: number | null }>({ x: null, y: null });
  /** 빈 곳에서 누르기 시작했는지. 글상자 안에서 글자를 끌어 선택하다 빈 곳에서 뗀 것은 클릭으로 치지 않는다. */
  const pressedOnEmpty = useRef(false);

  useEffect(() => {
    if (!focusId) return;
    editors.current.get(focusId)?.focus();
    setFocusId(null);
  }, [focusId]);

  const editBoxes = (edit: (boxes: TextBox[]) => TextBox[]) => onUpdate(note.id, (n) => ({ boxes: edit(n.boxes) }));

  const beginMove = (box: TextBox) => (e: ReactPointerEvent<HTMLElement>) => {
    if (e.button !== 0) return;
    e.preventDefault();
    const handle = e.currentTarget;
    handle.setPointerCapture(e.pointerId);
    const start = { px: e.clientX, py: e.clientY };
    let latest: { x: number; y: number } | null = null;
    // 다른 글상자들의 왼쪽 끝과 윗줄에 맞춰 붙는다.
    const others = note.boxes.filter((b) => b.id !== box.id);
    const move = (ev: PointerEvent) => {
      const x = snap(Math.max(0, box.x + ev.clientX - start.px), others.map((b) => b.x), SNAP_WITHIN);
      const y = snap(Math.max(0, box.y + ev.clientY - start.py), others.map((b) => b.y), SNAP_WITHIN);
      latest = { x: x.value, y: y.value };
      setMoving({ id: box.id, ...latest });
      setGuides({ x: x.guide, y: y.guide });
    };
    const end = () => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", end);
      handle.removeEventListener("pointercancel", end);
      const to = latest;
      if (to) editBoxes((boxes) => boxes.map((b) => (b.id === box.id ? { ...b, x: Math.round(to.x), y: Math.round(to.y) } : b)));
      setMoving(null);
      setGuides({ x: null, y: null });
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", end);
    handle.addEventListener("pointercancel", end);
  };

  const beginSize = (box: TextBox) => (e: ReactPointerEvent<HTMLElement>) => {
    if (e.button !== 0) return;
    e.preventDefault();
    const handle = e.currentTarget;
    handle.setPointerCapture(e.pointerId);
    const startX = e.clientX;
    // 폭을 정한 적이 없으면 지금 보이는 폭에서 시작한다.
    const startWidth = box.w > 0 ? box.w : (handle.parentElement?.offsetWidth ?? BOX_MIN_W);
    let latest: number | null = null;
    const move = (ev: PointerEvent) => {
      latest = Math.max(BOX_MIN_W, Math.round(startWidth + ev.clientX - startX));
      setSizing({ id: box.id, w: latest });
    };
    const end = () => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", end);
      handle.removeEventListener("pointercancel", end);
      const w = latest;
      if (w !== null) editBoxes((boxes) => boxes.map((b) => (b.id === box.id ? { ...b, w } : b)));
      setSizing(null);
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", end);
    handle.addEventListener("pointercancel", end);
  };

  return (
    <div
      className="note-canvas"
      style={{ fontSize: note.fontSize }}
      onMouseDown={(e) => (pressedOnEmpty.current = e.button === 0 && e.target === e.currentTarget)}
      onClick={(e) => {
        if (!pressedOnEmpty.current || e.target !== e.currentTarget) return;
        pressedOnEmpty.current = false;
        const el = e.currentTarget;
        const area = el.getBoundingClientRect();
        // 클릭한 자리에 첫 줄이 오게 한다. 오른쪽 끝에 너무 붙으면 글 쓸 폭이 없으니 조금 당겨 놓는다.
        // 이미 있는 글상자의 왼쪽 끝·윗줄 근처를 눌렀으면 거기에 맞춰 줄을 세운다.
        const box = newBox({
          x: snap(Math.min(e.clientX - area.left - 4, el.clientWidth - BOX_MIN_ROOM) + el.scrollLeft, note.boxes.map((b) => b.x), SNAP_WITHIN).value,
          y: snap(e.clientY - area.top + el.scrollTop - note.fontSize * 0.8, note.boxes.map((b) => b.y), SNAP_WITHIN).value,
        });
        editBoxes((boxes) => [...boxes, box]);
        setFocusId(box.id);
      }}
    >
      {note.boxes.length === 0 && <p className="canvas-hint">아무 데나 클릭해서 적으세요</p>}
      {guides.x !== null && <div className="guide guide-x" style={{ left: guides.x }} />}
      {guides.y !== null && <div className="guide guide-y" style={{ top: guides.y }} />}
      {note.boxes.map((box) => {
        const at = moving?.id === box.id ? moving : box;
        const width = sizing?.id === box.id ? sizing.w : box.w;
        return (
          <div
            key={box.id}
            className="box"
            // 폭을 정하지 않은 글상자는 메모지 오른쪽 끝에서 줄을 바꾸되, 끝에 붙어 있어도 최소한의 폭은 갖는다
            // (넘치면 메모지 안에서 스크롤된다).
            style={{
              left: at.x,
              top: at.y,
              ...(width > 0 ? { width } : { maxWidth: `max(${BOX_MIN_ROOM}px, calc(100% - ${at.x}px - 4px))` }),
            }}
          >
            <div className="box-grip" title="끌어서 옮기기" onPointerDown={beginMove(box)} />
            <RichText
              value={box}
              elementRef={(el) => {
                if (el) editors.current.set(box.id, el);
                else editors.current.delete(box.id);
              }}
              onChange={(rich) => editBoxes((boxes) => boxes.map((b) => (b.id === box.id ? { ...b, ...rich } : b)))}
              onBlur={() => editBoxes((boxes) => boxes.filter((b) => b.id !== box.id || b.text.trim() !== ""))}
            />
            <div
              className="box-width"
              title="끌어서 폭 조절 (더블클릭하면 내용에 맞춤)"
              onPointerDown={beginSize(box)}
              onDoubleClick={() => editBoxes((boxes) => boxes.map((b) => (b.id === box.id ? { ...b, w: 0 } : b)))}
            />
          </div>
        );
      })}
    </div>
  );
}

interface CardProps {
  note: Note;
  focusOnMount: boolean;
  onChange: (id: string, patch: Partial<Note>) => void;
  onUpdate: UpdateNote;
  onSwitchKind: (id: string) => void;
  onRemove: (note: Note) => void;
  onRaise: (id: string) => void;
  /** 띠를 우클릭했을 때 (다른 페이지로 보내기 등) */
  onBarMenu: (at: { x: number; y: number }, note: Note) => void;
}

const NoteCard = memo(function NoteCard(props: CardProps) {
  const { note, focusOnMount, onChange, onUpdate, onSwitchKind, onRemove, onRaise, onBarMenu } = props;
  // 끄는 동안에는 화면에서만 움직이고, 손을 뗄 때 한 번 저장한다.
  const [live, setLive] = useState<Partial<Note> | null>(null);
  const [paletteOpen, setPaletteOpen] = useState(false);
  const shown = live ? { ...note, ...live } : note;

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
      <div
        className="note-bar"
        title="끌어서 옮기기"
        onPointerDown={beginDrag("move")}
        onContextMenu={(e) => {
          e.preventDefault();
          onBarMenu({ x: e.clientX, y: e.clientY }, note);
        }}
      >
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
      {isList ? (
        <Checklist note={note} onChange={onChange} />
      ) : (
        <NoteCanvas note={note} focusFirst={focusOnMount} onUpdate={onUpdate} />
      )}
      {/* 글을 편집하는 동안에만 보이는 서식 줄. 선택한 글자에 적용된다. */}
      <div className="note-format">
        <button className="note-tool note-bold" title="굵게 (Ctrl+B)" aria-label="굵게" onMouseDown={keepSelection} onClick={format.bold}>
          B
        </button>
        <button className="note-tool note-italic" title="기울임 (Ctrl+I)" aria-label="기울임" onMouseDown={keepSelection} onClick={format.italic}>
          I
        </button>
        <button
          className="ink-color"
          style={{ background: "#2b2a26" }}
          title="기본 글자색으로"
          aria-label="기본 글자색으로"
          onMouseDown={keepSelection}
          onClick={() => format.color("")}
        />
        {INK_COLORS.map((color) => (
          <button
            key={color}
            className="ink-color"
            style={{ background: INK_HEX[color] }}
            title={`${INK_NAMES[color]} 글자로`}
            aria-label={`${INK_NAMES[color]} 글자로`}
            onMouseDown={keepSelection}
            onClick={() => format.color(color)}
          />
        ))}
      </div>
      <div className="note-grip" title="끌어서 크기 조절" onPointerDown={beginDrag("size")} />
    </div>
  );
});

interface BoardProps {
  /** 지금 페이지에 붙은 메모들 */
  notes: Note[];
  pages: Page[];
  activePage: string;
  /** 방금 만든 메모. 바로 글을 쓸 수 있게 커서를 넣는다. */
  focusId: string | null;
  boardRef: RefObject<HTMLDivElement>;
  onAdd: (at: { x: number; y: number }) => void;
  onChange: (id: string, patch: Partial<Note>) => void;
  onUpdate: UpdateNote;
  onSwitchKind: (id: string) => void;
  onRemove: (note: Note) => void;
  onRaise: (id: string) => void;
  onBarMenu: (at: { x: number; y: number }, note: Note) => void;
  onSelectPage: (id: string) => void;
  onAddPage: () => void;
  onPageMenu: (at: { x: number; y: number }, page: Page) => void;
}

export function NoteBoard(props: BoardProps) {
  const { notes, pages, activePage, focusId, boardRef, onAdd, onChange, onUpdate, onSwitchKind, onRemove, onRaise } = props;
  const { onBarMenu, onSelectPage, onAddPage, onPageMenu } = props;
  return (
    <div className="notes-pane">
      <div className="pages" role="tablist">
        {pages.map((page) => (
          <button
            key={page.id}
            role="tab"
            aria-selected={page.id === activePage}
            className={page.id === activePage ? "page current" : "page"}
            title={`${page.name} (우클릭: 이름 바꾸기 · 순서 · 삭제)`}
            onClick={() => onSelectPage(page.id)}
            onContextMenu={(e) => {
              e.preventDefault();
              onPageMenu({ x: e.clientX, y: e.clientY }, page);
            }}
          >
            {page.name}
          </button>
        ))}
        <button className="page-add" title="페이지 추가" aria-label="페이지 추가" onClick={onAddPage}>
          <PlusIcon />
        </button>
      </div>
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
            onUpdate={onUpdate}
            onSwitchKind={onSwitchKind}
            onRemove={onRemove}
            onRaise={onRaise}
            onBarMenu={onBarMenu}
          />
        ))}
      </div>
    </div>
  );
}
