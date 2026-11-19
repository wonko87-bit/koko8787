import { useEffect, useLayoutEffect, useRef, useState } from "react";

export interface MenuItem {
  label: string;
  onSelect: () => void;
  danger?: boolean;
  disabled?: boolean;
}

export interface MenuState {
  x: number;
  y: number;
  items: MenuItem[];
}

/** 누른 자리에 뜨는 메뉴. 위젯 밖으로 나가지 않게 안쪽으로 당겨 놓는다. */
export function ContextMenu({ menu, onClose }: { menu: MenuState; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({ left: menu.x, top: menu.y });

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    setPosition({
      left: Math.max(4, Math.min(menu.x, window.innerWidth - el.offsetWidth - 4)),
      top: Math.max(4, Math.min(menu.y, window.innerHeight - el.offsetHeight - 4)),
    });
  }, [menu]);

  useEffect(() => {
    const close = (e: Event) => {
      if (e instanceof KeyboardEvent && e.key !== "Escape") return;
      if (e instanceof MouseEvent && ref.current?.contains(e.target as Node)) return;
      onClose();
    };
    window.addEventListener("mousedown", close, true);
    window.addEventListener("keydown", close);
    window.addEventListener("blur", close);
    window.addEventListener("resize", close);
    return () => {
      window.removeEventListener("mousedown", close, true);
      window.removeEventListener("keydown", close);
      window.removeEventListener("blur", close);
      window.removeEventListener("resize", close);
    };
  }, [onClose]);

  return (
    <div ref={ref} className="menu" role="menu" style={position}>
      {menu.items.map((item) => (
        <button
          key={item.label}
          role="menuitem"
          className={item.danger ? "menu-item danger" : "menu-item"}
          disabled={item.disabled}
          onClick={() => {
            onClose();
            item.onSelect();
          }}
        >
          {item.label}
        </button>
      ))}
    </div>
  );
}
