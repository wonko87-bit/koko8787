import { useEffect, useRef, useState } from "react";
import { api } from "../api";
import type { Settings, TileSize } from "../types";

const TILE_SIZES: { value: TileSize; label: string }[] = [
  { value: "s", label: "작게" },
  { value: "m", label: "보통" },
  { value: "l", label: "크게" },
];

interface Props {
  settings: Settings;
  onChange: (patch: Partial<Settings>) => void;
  onResetWindow: () => void;
  onQuit: () => void;
  onClose: () => void;
}

export function SettingsPanel({ settings, onChange, onResetWindow, onQuit, onClose }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const [version, setVersion] = useState("");

  useEffect(() => {
    api.version().then(setVersion, () => {});
  }, []);

  useEffect(() => {
    const close = (e: Event) => {
      if (e instanceof KeyboardEvent && e.key !== "Escape") return;
      if (e instanceof MouseEvent) {
        const target = e.target as HTMLElement;
        if (ref.current?.contains(target) || target.closest("[data-settings-toggle]")) return;
      }
      onClose();
    };
    window.addEventListener("mousedown", close, true);
    window.addEventListener("keydown", close);
    return () => {
      window.removeEventListener("mousedown", close, true);
      window.removeEventListener("keydown", close);
    };
  }, [onClose]);

  return (
    <div ref={ref} className="settings" role="dialog" aria-label="설정">
      <div className="settings-row">
        <span>타일 크기</span>
        <div className="segmented">
          {TILE_SIZES.map(({ value, label }) => (
            <button
              key={value}
              className={settings.tileSize === value ? "active" : ""}
              aria-pressed={settings.tileSize === value}
              onClick={() => onChange({ tileSize: value })}
            >
              {label}
            </button>
          ))}
        </div>
      </div>
      <label className="settings-row">
        <span>배경 불투명도</span>
        <input
          type="range"
          min={20}
          max={100}
          step={2}
          value={Math.round(settings.opacity * 100)}
          onChange={(e) => onChange({ opacity: Number(e.target.value) / 100 })}
        />
        <output>{Math.round(settings.opacity * 100)}%</output>
      </label>
      <div className="settings-actions">
        <button onClick={onResetWindow}>위치·크기 초기화</button>
        <button className="danger" onClick={onQuit}>
          DeskBoard 종료
        </button>
      </div>
      {version && <p className="settings-version">DeskBoard {version}</p>}
    </div>
  );
}
