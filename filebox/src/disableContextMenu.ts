// WebView 기본 우클릭 메뉴(새로고침, 뒤로 등) 차단.
// 개발 모드에서는 디버깅을 위해 그대로 두고, 입력창에서는 복사/붙여넣기 메뉴를 허용한다.
export function installContextMenuBlocker() {
  if (import.meta.env.DEV) return;
  window.addEventListener("contextmenu", (e) => {
    const target = e.target as HTMLElement | null;
    if (target?.closest("input, textarea")) return;
    e.preventDefault();
  });
}
