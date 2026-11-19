// 서식 있는 글(글자 + 구간별 서식)과 편집 화면(DOM) 사이를 오가는 함수들.
//
// 저장하는 것은 HTML 이 아니라 "글자와, 어디부터 어디까지가 굵은지" 다. 그래서 편집기가
// 어떤 태그를 만들어 내든 화면에 실제로 보이는 모양(굵기·기울임·색)만 읽어 오면 된다.

import type { InkColor, Mark, Rich } from "./types";

export const INK_HEX: Record<InkColor, string> = {
  red: "#d92d20",
  orange: "#c2570c",
  green: "#1a7f37",
  blue: "#1d5fd1",
  purple: "#7a3fd0",
};

/** 화면이 알려 주는 색 표기("rgb(217, 45, 32)")에서 색 이름을 찾는다. */
const INK_BY_RGB = new Map<string, InkColor>(
  (Object.entries(INK_HEX) as [InkColor, string][]).map(([name, hex]) => {
    const n = parseInt(hex.slice(1), 16);
    return [`rgb(${n >> 16}, ${(n >> 8) & 255}, ${n & 255})`, name];
  }),
);

type Style = Pick<Mark, "bold" | "italic" | "color">;
const PLAIN: Style = { bold: false, italic: false, color: "" };
const sameStyle = (a: Style, b: Style) => a.bold === b.bold && a.italic === b.italic && a.color === b.color;
const isPlain = (s: Style) => !s.bold && !s.italic && !s.color;

export function sameRich(a: Rich, b: Rich): boolean {
  return (
    a.text === b.text &&
    a.marks.length === b.marks.length &&
    a.marks.every((m, i) => m.start === b.marks[i].start && m.end === b.marks[i].end && sameStyle(m, b.marks[i]))
  );
}

/** 글의 한 자리에 걸린 서식 */
function styleAt(marks: Mark[], index: number): Style {
  return marks.find((m) => m.start <= index && index < m.end) ?? PLAIN;
}

/** 글을 편집 화면에 그린다. */
export function renderRich(root: HTMLElement, value: Rich): void {
  root.textContent = "";
  const { text, marks } = value;
  let start = 0;
  while (start < text.length) {
    const style = styleAt(marks, start);
    let end = start + 1;
    while (end < text.length && sameStyle(styleAt(marks, end), style)) end++;

    // 서식은 안쪽부터 기울임 → 굵게 → 색 순서로 감싼다.
    let holder: Node = root;
    if (style.color) {
      const span = document.createElement("span");
      span.style.color = INK_HEX[style.color];
      holder = holder.appendChild(span);
    }
    if (style.bold) holder = holder.appendChild(document.createElement("b"));
    if (style.italic) holder = holder.appendChild(document.createElement("i"));

    text
      .slice(start, end)
      .split("\n")
      .forEach((line, i) => {
        if (i > 0) holder.appendChild(document.createElement("br"));
        if (line) holder.appendChild(document.createTextNode(line));
      });
    start = end;
  }
  // 줄바꿈으로 끝나면, 브라우저는 뒤에 <br> 이 하나 더 있어야 빈 마지막 줄을 보여 준다.
  if (text.endsWith("\n")) root.appendChild(document.createElement("br"));
}

/** 편집 화면에 지금 보이는 글과 서식을 읽어 온다. */
export function parseRich(root: HTMLElement, singleLine: boolean): Rich {
  let text = "";
  const marks: Mark[] = [];

  const add = (piece: string, style: Style) => {
    if (!piece) return;
    const last = marks[marks.length - 1];
    if (!isPlain(style)) {
      if (last && last.end === text.length && sameStyle(last, style)) last.end += piece.length;
      else marks.push({ start: text.length, end: text.length + piece.length, ...style });
    }
    text += piece;
  };

  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const looks = getComputedStyle(node.parentElement as HTMLElement);
      const piece = singleLine ? (node.nodeValue ?? "").replace(/\n/g, " ") : (node.nodeValue ?? "");
      add(piece, {
        bold: Number(looks.fontWeight) >= 600,
        italic: looks.fontStyle === "italic",
        color: INK_BY_RGB.get(looks.color) ?? "",
      });
      return;
    }
    if (!(node instanceof HTMLElement)) return;
    if (node.tagName === "BR") {
      if (!singleLine) add("\n", PLAIN);
      return;
    }
    // 줄 단위 요소(div, p)가 끼어들었으면 그 앞에서 줄을 바꾼다.
    const block = node !== root && (node.tagName === "DIV" || node.tagName === "P");
    if (block && !singleLine && text && !text.endsWith("\n")) add("\n", PLAIN);
    node.childNodes.forEach(walk);
  };
  walk(root);

  // 맨 끝의 <br> 하나는 빈 마지막 줄을 보여 주기 위한 것이지 줄바꿈이 아니다 (renderRich 참고).
  let last: Node | null = root.lastChild;
  while (last && last.lastChild) last = last.lastChild;
  if (!singleLine && last instanceof HTMLElement && last.tagName === "BR" && text.endsWith("\n")) {
    text = text.slice(0, -1);
  }
  return { text, marks: marks.filter((m) => m.start < text.length).map((m) => ({ ...m, end: Math.min(m.end, text.length) })) };
}

/** 글의 일부분을 서식과 함께 잘라 낸다. */
export function sliceRich(value: Rich, start: number, end: number): Rich {
  return {
    text: value.text.slice(start, end),
    marks: value.marks
      .filter((m) => m.start < end && m.end > start)
      .map((m) => ({ ...m, start: Math.max(m.start, start) - start, end: Math.min(m.end, end) - start })),
  };
}

/** 서식 있는 글 여럿을 사이에 구분 글자를 넣어 하나로 잇는다. */
export function joinRich(parts: Rich[], separator: string): Rich {
  let text = "";
  const marks: Mark[] = [];
  parts.forEach((part, i) => {
    if (i > 0) text += separator;
    for (const m of part.marks) marks.push({ ...m, start: m.start + text.length, end: m.end + text.length });
    text += part.text;
  });
  return { text, marks };
}
