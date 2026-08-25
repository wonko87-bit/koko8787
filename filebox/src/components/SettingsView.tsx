import { useEffect, useState } from "react";
import { open } from "@tauri-apps/plugin-dialog";
import { disable, enable, isEnabled } from "@tauri-apps/plugin-autostart";
import { ask, message } from "@tauri-apps/plugin-dialog";
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
  const [trashCount, setTrashCount] = useState<number | null>(null);
  const [emptying, setEmptying] = useState(false);

  const refreshTrash = () => {
    api.trashCount().then(setTrashCount).catch(() => setTrashCount(null));
  };

  useEffect(() => {
    isEnabled().then(setAutostart).catch(() => setAutostart(null));
    getVersion().then(setVersion).catch(() => {});
    refreshTrash();
  }, []);

  const emptyTrash = async () => {
    const ok = await ask(
      "윈도우 휴지통을 완전히 비웁니다.\n\n" +
        "FileBox에서 보낸 파일뿐 아니라 휴지통에 있는 모든 항목이 사라지며, " +
        "되돌릴 수 없습니다.\n\n정말 비울까요?",
      { title: "휴지통 비우기", kind: "warning", okLabel: "비우기", cancelLabel: "취소" },
    );
    if (!ok) return;
    setEmptying(true);
    try {
      const removed = await api.emptyTrash();
      await message(`휴지통을 비웠어요. (${removed}개 항목)`, { title: "휴지통 비우기" });
    } catch (e) {
      await message(String(e), { title: "휴지통 비우기 실패", kind: "error" });
    } finally {
      setEmptying(false);
      refreshTrash();
    }
  };

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

      <div className="panel danger-panel">
        <h2>휴지통</h2>
        <div className="row">
          <div className="grow">
            <div>
              윈도우 휴지통
              {trashCount !== null ? ` — ${trashCount}개 항목` : ""}
            </div>
            <div className="path">
              관리함에서 "휴지통으로" 보낸 파일은 여기에 있어요. 되돌리려면
              휴지통을 열어 복구하면 됩니다.
            </div>
          </div>
          <button onClick={() => api.openTrash().catch((e) => alert(String(e)))}>
            휴지통 열기
          </button>
          <button className="danger" onClick={emptyTrash} disabled={emptying}>
            {emptying ? "비우는 중…" : "휴지통 비우기"}
          </button>
        </div>
        <p className="danger-note">
          ⚠ 휴지통 비우기는 <strong>윈도우 휴지통 전체</strong>를 지웁니다.
          FileBox가 보낸 파일뿐 아니라 다른 프로그램에서 버린 파일까지 모두
          사라지며, 되돌릴 수 없습니다.
        </p>
      </div>

      <div className="panel">
        <h2>버전</h2>
        <div className="row">
          <div className="grow">
            FileBox {version ? `v${version}` : ""}
          </div>
          <button onClick={runUpdateCheck} disabled={checking}>
            {checking ? "확인 중…" : "업데이트 확인"}
          </button>
        </div>
        <div className="toggle-row">
          <div>
            <div>시작할 때 자동으로 업데이트 확인</div>
            <div className="desc">
              꺼두면 앱이 알아서 확인하지 않아요. 위의 "업데이트 확인" 버튼으로
              원할 때만 확인할 수 있고, 설치 여부는 언제나 직접 결정합니다.
            </div>
          </div>
          <input
            type="checkbox"
            checked={settings.auto_update_check}
            onChange={(e) =>
              save({ ...settings, auto_update_check: e.target.checked })
            }
          />
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
