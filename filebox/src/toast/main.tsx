import React, { useEffect, useRef, useState } from "react";
import ReactDOM from "react-dom/client";
import { listen } from "@tauri-apps/api/event";
import { api } from "../api";
import { installContextMenuBlocker } from "../disableContextMenu";
import { categoryIcon } from "../types";
import type { FileEntry, Suggestion } from "../types";
import "../styles.css";

installContextMenuBlocker();

interface ToastItem {
  entry: FileEntry;
  suggestions: Suggestion[];
}

const AUTO_HIDE_MS = 12_000;
const MAX_VISIBLE = 3;

function ToastApp() {
  const [items, setItems] = useState<ToastItem[]>([]);
  const itemsRef = useRef(items);
  itemsRef.current = items;

  useEffect(() => {
    const un = listen<ToastItem>("toast-file", (event) => {
      setItems((prev) => {
        const next = [event.payload, ...prev];
        return next.slice(0, MAX_VISIBLE);
      });
    });
    return () => {
      un.then((u) => u());
    };
  }, []);

  // 항목이 비면 창 숨김, 남아 있으면 일정 시간 후 자동 숨김
  useEffect(() => {
    if (items.length === 0) {
      api.hideToast();
      return;
    }
    const timer = setTimeout(() => {
      setItems([]);
    }, AUTO_HIDE_MS);
    return () => clearTimeout(timer);
  }, [items]);

  const dismiss = (id: string) =>
    setItems((prev) => prev.filter((i) => i.entry.id !== id));

  const send = async (item: ToastItem, favoriteId: string) => {
    try {
      await api.sendToFavorite(item.entry.id, favoriteId);
    } catch (e) {
      console.error(e);
    }
    dismiss(item.entry.id);
  };

  return (
    <div className="toast-stack">
      {items.map((item) => (
        <div className="toast-card" key={item.entry.id}>
          <div className="t-title">📦 FileBox — 새 파일 수집됨</div>
          <div className="t-name">
            {categoryIcon(item.entry.category)} {item.entry.file_name}
          </div>
          <div className="t-actions">
            {item.suggestions.map((s) => (
              <button
                key={s.favorite.id}
                className="suggest"
                title={s.favorite.path}
                onClick={() => send(item, s.favorite.id)}
              >
                ⭐ {s.favorite.name}
              </button>
            ))}
            <button
              onClick={() => {
                api.showMainWindow();
                dismiss(item.entry.id);
              }}
            >
              열어서 정리
            </button>
            <button onClick={() => dismiss(item.entry.id)}>나중에</button>
          </div>
        </div>
      ))}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <ToastApp />
  </React.StrictMode>,
);
