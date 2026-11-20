import { useEffect, useRef, useState } from "react";
import type { KeyboardEvent } from "react";
import { api } from "../api";
import { THEMES } from "../types";
import type { Settings, TileSize } from "../types";

const TILE_SIZES: { value: TileSize; label: string }[] = [
  { value: "s", label: "작게" },
  { value: "m", label: "보통" },
  { value: "l", label: "크게" },
];

const MODIFIER_KEYS = ["Control", "Alt", "Shift", "Meta"];

/**
 * 누른 키 조합을 "Ctrl+Alt+D" 꼴로 만든다. 보조 키만 누른 상태면 null.
 * 글자 하나만으로는 평소 타자와 겹치므로 Ctrl·Alt·Win 중 하나는 있어야 한다 (F1~F24 는 예외).
 */
export function hotkeyFromEvent(e: Pick<KeyboardEvent, "key" | "code" | "ctrlKey" | "altKey" | "shiftKey" | "metaKey">):
  | { hotkey: string }
  | { error: string }
  | null {
  if (MODIFIER_KEYS.includes(e.key)) return null;
  const key = e.code.replace(/^Key/, "").replace(/^Digit/, "");
  const isFunctionKey = /^F\d{1,2}$/.test(key);
  if (!e.ctrlKey && !e.altKey && !e.metaKey && !isFunctionKey) {
    return { error: "Ctrl, Alt, Win 중 하나와 함께 눌러 주세요." };
  }
  const parts = [e.ctrlKey && "Ctrl", e.altKey && "Alt", e.shiftKey && "Shift", e.metaKey && "Super", key];
  return { hotkey: parts.filter(Boolean).join("+") };
}

interface Props {
  settings: Settings;
  /** 설치할 수 있는 새 버전 (없으면 null) */
  updateVersion: string | null;
  onChange: (patch: Partial<Settings>) => void;
  onCheckUpdate: () => Promise<string>;
  onInstallUpdate: () => void;
  onExport: () => void;
  onImport: () => void;
  onOpenBackups: () => void;
  onResetWindow: () => void;
  onQuit: () => void;
  onClose: () => void;
}

export function SettingsPanel(props: Props) {
  const { settings, updateVersion, onChange, onCheckUpdate, onInstallUpdate } = props;
  const { onExport, onImport, onOpenBackups, onResetWindow, onQuit, onClose } = props;
  const ref = useRef<HTMLDivElement>(null);
  const [version, setVersion] = useState("");
  const [fonts, setFonts] = useState<string[]>([]);
  const [autostart, setAutostart] = useState<boolean | null>(null);
  const [capturing, setCapturing] = useState(false);
  const [hotkeyError, setHotkeyError] = useState<string | null>(null);
  const [updateStatus, setUpdateStatus] = useState<string | null>(null);

  useEffect(() => {
    api.version().then(setVersion, () => {});
    api.listFonts().then(setFonts, () => {});
    api.getAutostart().then(setAutostart, () => {});
  }, []);

  useEffect(() => {
    const close = (e: Event) => {
      if (e instanceof KeyboardEvent && (e.key !== "Escape" || capturing)) return;
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
  }, [onClose, capturing]);

  const applyHotkey = async (hotkey: string) => {
    try {
      await api.setHotkey(hotkey);
      onChange({ hotkey });
      setHotkeyError(null);
    } catch (e) {
      setHotkeyError(String(e));
    }
  };

  const captureHotkey = (e: KeyboardEvent<HTMLButtonElement>) => {
    if (!capturing) return;
    e.preventDefault();
    e.stopPropagation();
    if (e.key === "Escape") return setCapturing(false);
    const result = hotkeyFromEvent(e);
    if (!result) return;
    if ("error" in result) return setHotkeyError(result.error);
    setCapturing(false);
    applyHotkey(result.hotkey);
  };

  const checkNow = async () => {
    setUpdateStatus("확인하는 중…");
    setUpdateStatus(await onCheckUpdate());
  };

  // 지금 쓰는 글꼴이 목록에 없어도(다른 PC 에서 가져온 보드 등) 선택 상자에는 보이게 한다.
  const fontChoices = settings.noteFont && !fonts.includes(settings.noteFont) ? [settings.noteFont, ...fonts] : fonts;

  return (
    <div ref={ref} className="settings" role="dialog" aria-label="설정">
      <h4>모양</h4>
      <div className="settings-row">
        <span>테마</span>
        <div className="theme-choices">
          {THEMES.map(({ id, name }) => (
            <button
              key={id}
              className={settings.theme === id ? "theme-choice current" : "theme-choice"}
              data-theme={id}
              title={name}
              aria-label={`${name} 테마`}
              aria-pressed={settings.theme === id}
              onClick={() => onChange({ theme: id })}
            >
              <i />
            </button>
          ))}
        </div>
      </div>
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
      <label className="settings-row">
        <span>메모 글꼴</span>
        <select value={settings.noteFont} onChange={(e) => onChange({ noteFont: e.target.value })}>
          <option value="">기본</option>
          {fontChoices.map((font) => (
            <option key={font} value={font}>
              {font}
            </option>
          ))}
        </select>
      </label>

      <h4>동작</h4>
      <div className="settings-row">
        <span>불러오기 단축키</span>
        <button
          className={capturing ? "hotkey capturing" : "hotkey"}
          title="눌러서 새 단축키 지정"
          onClick={() => {
            setCapturing(true);
            setHotkeyError(null);
          }}
          onKeyDown={captureHotkey}
          onBlur={() => setCapturing(false)}
        >
          {capturing ? "키를 누르세요…" : settings.hotkey || "없음"}
        </button>
        {settings.hotkey && !capturing && (
          <button className="link-button" onClick={() => applyHotkey("")}>
            지우기
          </button>
        )}
      </div>
      {hotkeyError && <p className="settings-error">{hotkeyError}</p>}
      <label className="settings-check">
        <input
          type="checkbox"
          checked={autostart === true}
          disabled={autostart === null}
          onChange={(e) => api.setAutostart(e.target.checked).then(setAutostart, () => {})}
        />
        Windows 시작 시 자동 실행
      </label>
      <div className="settings-row">
        <label className="settings-check">
          <input
            type="checkbox"
            checked={settings.checkUpdates}
            onChange={(e) => onChange({ checkUpdates: e.target.checked })}
          />
          시작할 때 업데이트 확인
        </label>
        <button className="link-button" onClick={checkNow}>
          지금 확인
        </button>
      </div>
      {updateVersion ? (
        <p className="settings-note">
          새 버전 v{updateVersion} 이 있습니다.{" "}
          <button className="link-button" onClick={onInstallUpdate}>
            설치
          </button>
        </p>
      ) : (
        updateStatus && <p className="settings-note">{updateStatus}</p>
      )}

      <h4>보드</h4>
      <div className="settings-actions">
        <button onClick={onExport}>내보내기…</button>
        <button onClick={onImport}>가져오기…</button>
      </div>
      <p className="settings-note">
        하루 한 번 자동으로 백업해 2주 치를 보관합니다.{" "}
        <button className="link-button" onClick={onOpenBackups}>
          백업 폴더 열기
        </button>
      </p>
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
