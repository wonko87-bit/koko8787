import { useLayoutEffect, useRef } from "react";
import type { ClipboardEvent, CSSProperties, KeyboardEvent } from "react";
import { INK_HEX, parseRich, renderRich, sameRich } from "../richText";
import type { InkColor, Rich } from "../types";

/** 메모 종이 위의 기본 글자색. 색을 "기본"으로 되돌릴 때 이 색을 입힌다. */
const DEFAULT_INK = "#2b2a26";

/** 지금 선택한 글자에 서식을 입힌다. 편집 칸 안에 선택이 있을 때만 뜻이 있다. */
export const format = {
  bold: () => document.execCommand("bold"),
  italic: () => document.execCommand("italic"),
  color: (color: InkColor | "") => document.execCommand("foreColor", false, color ? INK_HEX[color] : DEFAULT_INK),
};

interface Props {
  value: Rich;
  /** 한 줄짜리 칸(체크리스트 항목). Enter 는 줄을 바꾸지 않고 onKeyDown 으로만 전해진다. */
  singleLine?: boolean;
  placeholder?: string;
  className?: string;
  style?: CSSProperties;
  onChange: (value: Rich) => void;
  onKeyDown?: (e: KeyboardEvent<HTMLDivElement>) => void;
  onBlur?: () => void;
  /** 편집 칸 요소를 받아 간다 (포커스를 옮기는 데 쓴다). */
  elementRef?: (el: HTMLDivElement | null) => void;
}

/**
 * 굵게·기울임·글자색을 입힐 수 있는 편집 칸.
 *
 * 글자를 치는 동안에는 브라우저가 화면을 직접 고치고, 우리는 그 결과를 읽어 올리기만 한다.
 * 화면을 다시 그리는 것은 값이 밖에서 바뀌었을 때(불러오기, 체크리스트 전환 등)뿐이다.
 * 칠 때마다 다시 그리면 커서가 맨 앞으로 튀고 한글 조합이 끊긴다.
 */
export function RichText(props: Props) {
  const { value, singleLine = false, placeholder, className, style, onChange, onKeyDown, onBlur, elementRef } = props;
  const ref = useRef<HTMLDivElement | null>(null);
  /** 화면에 지금 그려져 있는 값 */
  const shown = useRef<Rich | null>(null);

  useLayoutEffect(() => {
    if (!ref.current || (shown.current && sameRich(shown.current, value))) return;
    renderRich(ref.current, value);
    shown.current = value;
  }, [value]);

  const readBack = () => {
    const el = ref.current;
    if (!el) return;
    const next = parseRich(el, singleLine);
    // 다 지우면 브라우저가 <br> 하나를 남겨 두는데, 그러면 안내 글귀(placeholder)가 안 보인다.
    if (next.text === "" && el.childNodes.length > 0) el.textContent = "";
    shown.current = next;
    onChange(next);
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    onKeyDown?.(e);
    if (e.defaultPrevented || e.nativeEvent.isComposing) return;
    if (e.key === "Enter") {
      e.preventDefault();
      // 기본 동작은 줄마다 <div> 를 만든다. <br> 로 줄을 바꿔 구조를 단순하게 둔다.
      if (!singleLine) document.execCommand("insertLineBreak");
    }
  };

  // 다른 곳에서 복사해 온 글은 서식을 떼고 글자만 넣는다.
  const handlePaste = (e: ClipboardEvent<HTMLDivElement>) => {
    e.preventDefault();
    const lines = e.clipboardData.getData("text/plain").replace(/\r/g, "").split("\n");
    if (singleLine) return void document.execCommand("insertText", false, lines.join(" "));
    lines.forEach((line, i) => {
      if (i > 0) document.execCommand("insertLineBreak");
      if (line) document.execCommand("insertText", false, line);
    });
  };

  return (
    <div
      ref={(el) => {
        ref.current = el;
        elementRef?.(el);
      }}
      className={className ? `rich ${className}` : "rich"}
      style={style}
      contentEditable
      suppressContentEditableWarning
      role="textbox"
      aria-multiline={!singleLine}
      aria-label={placeholder}
      data-placeholder={placeholder}
      spellCheck={false}
      onInput={readBack}
      onBlur={onBlur}
      onKeyDown={handleKeyDown}
      onPaste={handlePaste}
      onDrop={(e) => e.preventDefault()}
    />
  );
}
