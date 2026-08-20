import { useState } from "react";
import { open } from "@tauri-apps/plugin-dialog";
import { api } from "../api";
import { openFolder } from "../openFolder";
import type { Favorite } from "../types";

export default function FavoritesView({ favorites }: { favorites: Favorite[] }) {
  const [name, setName] = useState("");
  const [path, setPath] = useState("");

  const pickFolder = async () => {
    const dir = await open({ directory: true, title: "즐겨찾기 폴더 선택" });
    if (typeof dir === "string") {
      setPath(dir);
      if (!name) {
        const base = dir.split(/[\\/]/).filter(Boolean).pop() ?? dir;
        setName(base);
      }
    }
  };

  const add = async () => {
    if (!name.trim() || !path) return;
    try {
      await api.addFavorite(name.trim(), path);
      setName("");
      setPath("");
    } catch (e) {
      alert(String(e));
    }
  };

  return (
    <>
      <div className="panel">
        <h2>즐겨찾기 추가</h2>
        <div className="form-grid">
          <label>폴더</label>
          <div style={{ display: "flex", gap: 8 }}>
            <input
              style={{ flex: 1 }}
              value={path}
              readOnly
              placeholder="폴더를 선택하세요"
            />
            <button onClick={pickFolder}>폴더 선택…</button>
          </div>
          <label>표시 이름</label>
          <input
            value={name}
            onChange={(e) => setName(e.target.value)}
            placeholder="예: 회의자료"
          />
        </div>
        <div className="form-actions">
          <button className="primary" onClick={add} disabled={!name.trim() || !path}>
            추가
          </button>
        </div>
        <div className="hint">
          즐겨찾기 폴더는 파일의 최종 저장처예요. 파일 카드와 알림 토스트의
          원클릭 이동 버튼으로 사용됩니다.
        </div>
      </div>

      <div className="panel">
        <h2>등록된 즐겨찾기 ({favorites.length})</h2>
        {favorites.length === 0 && (
          <div className="hint">아직 즐겨찾기가 없어요. 위에서 추가해보세요.</div>
        )}
        {favorites.map((f) => (
          <div className="row" key={f.id}>
            <div className="grow">
              <div>⭐ {f.name}</div>
              <div className="path">{f.path}</div>
            </div>
            <button onClick={() => openFolder(f.path)}>열기</button>
            <button className="ghost" onClick={() => api.removeFavorite(f.id)}>
              삭제
            </button>
          </div>
        ))}
      </div>
    </>
  );
}
