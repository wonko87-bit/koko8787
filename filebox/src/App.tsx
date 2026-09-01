import { useCallback, useEffect, useRef, useState } from "react";
import { listen } from "@tauri-apps/api/event";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { api } from "./api";
import { checkForUpdate } from "./updater";
import type { FileEntry, Favorite, Rule, Settings } from "./types";
import InboxView from "./components/InboxView";
import HistoryView from "./components/HistoryView";
import FavoritesView from "./components/FavoritesView";
import RulesView from "./components/RulesView";
import SettingsView from "./components/SettingsView";
import CollectModal from "./components/CollectModal";

const TABS = ["관리함", "기록", "즐겨찾기", "규칙", "설정"] as const;
type Tab = (typeof TABS)[number];

export default function App() {
  const [tab, setTab] = useState<Tab>("관리함");
  const [entries, setEntries] = useState<FileEntry[]>([]);
  const [favorites, setFavorites] = useState<Favorite[]>([]);
  const [rules, setRules] = useState<Rule[]>([]);
  const [settings, setSettings] = useState<Settings | null>(null);
  const [collectOpen, setCollectOpen] = useState(false);
  const [dragOver, setDragOver] = useState(false);

  const reloadEntries = useCallback(() => {
    api.listEntries().then(setEntries).catch(console.error);
  }, []);
  const reloadFavorites = useCallback(() => {
    api.listFavorites().then(setFavorites).catch(console.error);
  }, []);
  const reloadRules = useCallback(() => {
    api.listRules().then(setRules).catch(console.error);
  }, []);
  const reloadSettings = useCallback(() => {
    api.getSettings().then(setSettings).catch(console.error);
  }, []);

  useEffect(() => {
    reloadEntries();
    reloadFavorites();
    reloadRules();
    reloadSettings();
    const unsubs = [
      listen("entries-changed", reloadEntries),
      listen("favorites-changed", reloadFavorites),
      listen("rules-changed", reloadRules),
      listen("settings-changed", reloadSettings),
    ];
    return () => {
      unsubs.forEach((p) => p.then((un) => un()));
    };
  }, [reloadEntries, reloadFavorites, reloadRules, reloadSettings]);

  // 창에 파일을 끌어다 놓으면 감시 폴더 밖의 파일도 관리함으로 수집한다.
  useEffect(() => {
    const unlisten = getCurrentWebview().onDragDropEvent((event) => {
      if (event.payload.type === "over") {
        setDragOver(true);
      } else if (event.payload.type === "drop") {
        setDragOver(false);
        const paths = event.payload.paths;
        if (paths.length > 0) {
          api
            .collectPaths(paths)
            .then((result) => {
              if (result.errors.length === 0) return;
              // "수집할 수 있는 파일이 없다"고만 하면, 실은 다른 프로그램이
              // 붙잡고 있어서 못 옮긴 경우와 구분이 안 된다.
              alert(
                (result.moved > 0 ? `${result.moved}개 수집. ` : "") +
                  `${result.errors.length}개는 가져오지 못했어요.\n\n` +
                  result.errors.slice(0, 10).join("\n") +
                  "\n\n파일을 열어 둔 프로그램이 있으면 닫고 다시 시도해 보세요.",
              );
            })
            .catch((e) => alert(String(e)));
        }
      } else {
        setDragOver(false);
      }
    });
    return () => {
      unlisten.then((un) => un());
    };
  }, []);

  // 설정이 읽힌 뒤 한 번만, 그리고 사용자가 켜둔 경우에만 조용히 업데이트를 확인한다.
  const updateChecked = useRef(false);
  useEffect(() => {
    if (!settings || updateChecked.current) return;
    updateChecked.current = true;
    if (settings.auto_update_check) checkForUpdate(false);
  }, [settings]);

  const inboxCount = entries.filter((e) => e.status === "inbox").length;

  const togglePause = async () => {
    if (!settings) return;
    const next = { ...settings, paused: !settings.paused };
    setSettings(await api.updateSettings(next));
  };

  return (
    <>
      <div className="header">
        <h1>📦 FileBox</h1>
        {settings &&
          (settings.paused ? (
            <span className="badge paused">
              <span className="dot" /> 감시 일시정지
            </span>
          ) : (
            <span className="badge">
              <span className="dot" /> 감시 중
            </span>
          ))}
        <div className="spacer" />
        <button onClick={() => setCollectOpen(true)}>기존 파일 수집…</button>
        {settings && (
          <button onClick={togglePause}>
            {settings.paused ? "감시 재개" : "일시정지"}
          </button>
        )}
      </div>

      <div className="tabs">
        {TABS.map((t) => (
          <button
            key={t}
            className={t === tab ? "active" : ""}
            onClick={() => setTab(t)}
          >
            {t}
            {t === "관리함" && inboxCount > 0 ? ` (${inboxCount})` : ""}
          </button>
        ))}
      </div>

      <div className="content">
        {tab === "관리함" && (
          <InboxView entries={entries} favorites={favorites} />
        )}
        {tab === "기록" && <HistoryView entries={entries} />}
        {tab === "즐겨찾기" && <FavoritesView favorites={favorites} />}
        {tab === "규칙" && <RulesView rules={rules} favorites={favorites} />}
        {tab === "설정" && settings && (
          <SettingsView settings={settings} onChange={setSettings} />
        )}
      </div>

      {dragOver && (
        <div className="drop-overlay">
          <div className="drop-card">
            <div className="drop-icon">📥</div>
            여기에 놓으면 관리함으로 수집됩니다
          </div>
        </div>
      )}

      {collectOpen && (
        <CollectModal
          onClose={() => {
            setCollectOpen(false);
            reloadEntries();
          }}
        />
      )}
    </>
  );
}
