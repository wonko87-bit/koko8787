import { useState } from "react";
import { api } from "../api";
import type { Favorite, Rule } from "../types";

const EMPTY: Rule = {
  id: "",
  name: "",
  extensions: [],
  keywords: [],
  category: null,
  favorite_id: null,
};

export default function RulesView({
  rules,
  favorites,
}: {
  rules: Rule[];
  favorites: Favorite[];
}) {
  const [draft, setDraft] = useState<Rule>(EMPTY);
  const [extText, setExtText] = useState("");
  const [kwText, setKwText] = useState("");

  const editing = draft.id !== "";

  const startEdit = (rule: Rule) => {
    setDraft(rule);
    setExtText(rule.extensions.join(", "));
    setKwText(rule.keywords.join(", "));
  };

  const reset = () => {
    setDraft(EMPTY);
    setExtText("");
    setKwText("");
  };

  const save = async () => {
    const rule: Rule = {
      ...draft,
      name: draft.name.trim() || "이름 없는 규칙",
      extensions: extText.split(",").map((s) => s.trim()).filter(Boolean),
      keywords: kwText.split(",").map((s) => s.trim()).filter(Boolean),
      category: draft.category?.trim() ? draft.category.trim() : null,
      favorite_id: draft.favorite_id || null,
    };
    if (rule.extensions.length === 0 && rule.keywords.length === 0) {
      alert("확장자 또는 키워드를 하나 이상 입력하세요.");
      return;
    }
    await api.upsertRule(rule);
    reset();
  };

  const favName = (id: string | null) =>
    favorites.find((f) => f.id === id)?.name ?? null;

  return (
    <>
      <div className="panel">
        <h2>{editing ? "규칙 수정" : "새 규칙"}</h2>
        <div className="form-grid">
          <label>규칙 이름</label>
          <input
            value={draft.name}
            onChange={(e) => setDraft({ ...draft, name: e.target.value })}
            placeholder="예: 청구서 PDF"
          />
          <label>확장자</label>
          <input
            value={extText}
            onChange={(e) => setExtText(e.target.value)}
            placeholder="쉼표로 구분 · 예: pdf, hwp"
          />
          <label>파일명 키워드</label>
          <input
            value={kwText}
            onChange={(e) => setKwText(e.target.value)}
            placeholder="쉼표로 구분 · 예: 청구서, invoice"
          />
          <label>카테고리</label>
          <input
            value={draft.category ?? ""}
            onChange={(e) => setDraft({ ...draft, category: e.target.value })}
            placeholder="비우면 확장자 기본 카테고리 사용"
          />
          <label>추천 즐겨찾기</label>
          <select
            value={draft.favorite_id ?? ""}
            onChange={(e) =>
              setDraft({ ...draft, favorite_id: e.target.value || null })
            }
          >
            <option value="">(없음)</option>
            {favorites.map((f) => (
              <option key={f.id} value={f.id}>
                {f.name}
              </option>
            ))}
          </select>
        </div>
        <div className="form-actions">
          {editing && <button onClick={reset}>취소</button>}
          <button className="primary" onClick={save}>
            {editing ? "저장" : "규칙 추가"}
          </button>
        </div>
        <div className="hint">
          확장자와 키워드를 모두 입력하면 둘 다 만족해야 매칭돼요. 매칭된 파일은
          지정한 카테고리로 분류되고, 지정한 즐겨찾기가 최우선 추천으로 떠요.
        </div>
      </div>

      <div className="panel">
        <h2>등록된 규칙 ({rules.length})</h2>
        {rules.length === 0 && (
          <div className="hint">
            규칙이 없어도 확장자 기본 분류(문서/이미지/영상…)는 자동으로
            적용돼요.
          </div>
        )}
        {rules.map((r) => (
          <div className="row" key={r.id}>
            <div className="grow">
              <div>{r.name}</div>
              <div className="path">
                {r.extensions.length > 0 && `확장자: ${r.extensions.join(", ")}`}
                {r.extensions.length > 0 && r.keywords.length > 0 && " · "}
                {r.keywords.length > 0 && `키워드: ${r.keywords.join(", ")}`}
                {r.category && ` → 카테고리: ${r.category}`}
                {favName(r.favorite_id) && ` → ⭐ ${favName(r.favorite_id)}`}
              </div>
            </div>
            <button onClick={() => startEdit(r)}>수정</button>
            <button className="ghost" onClick={() => api.removeRule(r.id)}>
              삭제
            </button>
          </div>
        ))}
      </div>
    </>
  );
}
