import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./styles.css";

// 브라우저 미리보기에서는 투명 창 대신 바탕화면 비슷한 배경을 깔아 준다.
if (!("__TAURI_INTERNALS__" in window)) {
  document.documentElement.classList.add("preview");
}

// WebView 기본 우클릭 메뉴(새로고침, 뒤로 등)는 막고, 입력창의 복사/붙여넣기 메뉴만 남긴다.
window.addEventListener("contextmenu", (e) => {
  const target = e.target as HTMLElement | null;
  if (import.meta.env.DEV || target?.closest("input, textarea")) return;
  e.preventDefault();
});

// 파일 놓기는 백엔드가 받는다. 혹시 웹뷰가 먼저 받게 되더라도 그 파일을 화면에 열어 버리지 않게 한다.
for (const type of ["dragover", "drop"]) {
  window.addEventListener(type, (e) => {
    if ((e as DragEvent).dataTransfer?.types.includes("Files")) e.preventDefault();
  });
}

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
