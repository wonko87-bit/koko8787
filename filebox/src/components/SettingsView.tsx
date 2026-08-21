import { useEffect, useState } from "react";
import { open } from "@tauri-apps/plugin-dialog";
import { disable, enable, isEnabled } from "@tauri-apps/plugin-autostart";
import { getVersion } from "@tauri-apps/api/app";
import { api } from "../api";
import { openFolder } from "../openFolder";
import { checkForUpdate } from "../updater";
import type { Settings } from "../types";

export default function SettingsView({
  settings,
  onChange,
}: {
  settings: Settings;
  onChange: (s: Settings) => void;
}) {
  const [autostart, setAutostart] = useState<boolean | null>(null);
  const [version, setVersion] = useState("");
  const [checking, setChecking] = useState(false);

  useEffect(() => {
    isEnabled().then(setAutostart).catch(() => setAutostart(null));
    getVersion().then(setVersion).catch(() => {});
  }, []);

  const runUpdateCheck = async () => {
    setChecking(true);
    await checkForUpdate(true);
    setChecking(false);
  };

  const toggleAutostart = async (on: boolean) => {
    try {
      if (on) await enable();
      else await disable();
      setAutostart(await isEnabled());
    } catch (e) {
      alert(String(e));
    }
  };

  const save = async (next: Settings) => {
    onChange(await api.updateSettings(next));
  };

  const addWatchDir = async () => {
    const dir = await open({ directory: true, title: "감시할 폴더 선택" });
    if (typeof dir === "string" && !settings.watch_dirs.includes(dir)) {
      await save({ ...settings, watch_dirs: [...settings.watch_dirs, dir] });
    }
  };

  const removeWatchDir = async (dir: string) => {
    await save({
      ...settings,
      watch_dirs: settings.watch_dirs.filter((d) => d !== dir),
    });
  };

  const changeInbox = async () => {
    const dir = await open({ directory: true, title: "관리함 폴더 선택" });
    if (typeof dir === "string") {
      await save({ ...settings, inbox_dir: dir });
    }
  };

  return (
    <>
      <div className="panel">
        <h2>감시 폴더</h2>
        {settings.watch_dirs.map((d) => (
          <div className="row" key={d}>
            <div className="grow path">{d}</div>
            <button onClick={() => openFolder(d)}>열기</button>
            <button className="ghost" onClick={() => removeWatchDir(d)}>
              제거
            </button>
          </div>
        ))}
        <div className="form-actions">
          <button onClick={addWatchDir}>감시 폴더 추가…</button>
        </div>
        <div className="hint">
          이 폴더들에 새로 생기는 파일이 관리함으로 자동 수집돼요. (앱 실행 중에
          생긴 파일만 대상 · 기존 파일은 상단의 "기존 파일 수집"으로)
        </div>
      </div>

      <div className="panel">
        <h2>관리함 폴더</h2>
        <div className="row">
          <div className="grow path">{settings.inbox_dir}</div>
          <button onClick={() => openFolder(settings.inbox_dir)}>열기</button>
          <button onClick={changeInbox}>변경…</button>
        </div>
        <div className="hint">
          수집된 파일이 실제로 보관되는 폴더예요. 분류는 폴더 이동 없이 앱
          안에서 가상으로 관리됩니다.
        </div>
      </div>

      <div className="panel">
        <h2>버전</h2>
        <div className="row">
          <div className="grow">
            FileBox {version ? `v${version}` : ""}
            <div className="path">
              업데이트가 있으면 시작할 때 알려드려요.
            </div>
          </div>
          <button onClick={runUpdateCheck} disabled={checking}>
            {checking ? "확인 중…" : "업데이트 확인"}
          </button>
        </div>
      </div>

      <div className="panel">
        <h2>동작</h2>
        <div className="toggle-row">
          <div>
            <div>윈도우 시작 시 자동 실행</div>
            <div className="desc">
              FileBox가 켜져 있어야 새 파일이 자동 수집돼요. 트레이에 조용히
              상주합니다.
            </div>
          </div>
          <input
            type="checkbox"
            checked={autostart === true}
            disabled={autostart === null}
            onChange={(e) => toggleAutostart(e.target.checked)}
          />
        </div>
        <div className="toggle-row">
          <div>
            <div>자동 수집</div>
            <div className="desc">
              끄면 감시는 하지 않고, "기존 파일 수집"으로만 가져와요.
            </div>
          </div>
          <input
            type="checkbox"
            checked={settings.auto_collect}
            onChange={(e) => save({ ...settings, auto_collect: e.target.checked })}
          />
        </div>
        <div className="toggle-row">
          <div>
            <div>새 파일 토스트 알림</div>
            <div className="desc">
              끄면 화면 우하단 토스트 대신 조용한 OS 알림만 표시돼요.
            </div>
          </div>
          <input
            type="checkbox"
            checked={settings.toast_enabled}
            onChange={(e) => save({ ...settings, toast_enabled: e.target.checked })}
          />
        </div>
        <div className="toggle-row">
          <div>
            <div>감시 일시정지</div>
            <div className="desc">트레이 메뉴에서도 전환할 수 있어요.</div>
          </div>
          <input
            type="checkbox"
            checked={settings.paused}
            onChange={(e) => save({ ...settings, paused: e.target.checked })}
          />
        </div>
      </div>
    </>
  );
}
