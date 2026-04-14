/* TaskBridge – frontend logic */

// ---- State ----
const state = { google: false, todoist: false };

// ---- DOM refs ----
const textarea     = document.getElementById("input");
const submitBtn    = document.getElementById("submit-btn");
const targetPreview = document.getElementById("target-preview");
const historyList  = document.getElementById("history-list");
const toast        = document.getElementById("toast");
const badgeGoogle  = document.getElementById("badge-google");
const badgeTodoist = document.getElementById("badge-todoist");

let toastTimer = null;

// ---- Init ----
document.addEventListener("DOMContentLoaded", () => {
  fetchStatus();
  textarea.addEventListener("input", updatePreview);
  textarea.addEventListener("keydown", handleKeydown);
  document.getElementById("submit-btn").addEventListener("click", handleSubmit);

  // Hint chips
  document.querySelectorAll(".hint").forEach(chip => {
    chip.addEventListener("click", () => {
      const prefix = chip.dataset.prefix;
      if (!prefix) return;
      const cur = textarea.value.trimStart();
      // Remove existing prefix if present
      const clean = cur.replace(/^\/[ct]\s*/, "");
      textarea.value = prefix ? `${prefix} ${clean}` : clean;
      textarea.focus();
      updatePreview();
    });
  });

  // Check URL params for post-auth feedback
  const params = new URLSearchParams(window.location.search);
  if (params.get("connected") === "google")  showToast("✅ Google Calendar 연결 완료!", "success");
  if (params.get("connected") === "todoist") showToast("✅ Todoist 연결 완료!", "success");
  if (params.get("error"))                   showToast("⚠️ 연결에 실패했습니다.", "warn");
  // Clean URL
  if (params.toString()) window.history.replaceState({}, "", "/");
});

// ---- Auth status ----
async function fetchStatus() {
  try {
    const res  = await fetch("/api/status");
    const data = await res.json();
    state.google  = !!data.google;
    state.todoist = !!data.todoist;
    updateBadges();
    updateAuthPanel();
  } catch (e) {
    console.error("Status fetch failed", e);
  }
}

function updateBadges() {
  setBadge(badgeGoogle,  state.google,  "Google Cal");
  setBadge(badgeTodoist, state.todoist, "Todoist");
}

function setBadge(el, connected, label) {
  if (!el) return;
  el.innerHTML = `<span class="dot"></span>${label}`;
  el.className = "badge" + (connected ? " connected" : "");
}

function updateAuthPanel() {
  const panel = document.getElementById("auth-panel");
  if (!panel) return;

  const gBtn = document.getElementById("btn-google");
  const tBtn = document.getElementById("btn-todoist");

  if (state.google) {
    gBtn.classList.add("connected-btn");
    gBtn.innerHTML = "✓ Google Calendar 연결됨";
  }
  if (state.todoist) {
    tBtn.classList.add("connected-btn");
    tBtn.innerHTML = "✓ Todoist 연결됨";
  }

  if (state.google && state.todoist) {
    panel.style.display = "none";
  }
}

// ---- Input parsing preview ----
function parseTarget(text) {
  const s = text.trimStart();
  if (s.startsWith("/c ") || s === "/c") return "calendar";
  if (s.startsWith("/t ") || s === "/t") return "todoist";
  return "both";
}

function targetLabel(t) {
  if (t === "calendar") return "<span>📅 Google Calendar</span>에만 저장";
  if (t === "todoist")  return "<span>✅ Todoist</span>에만 저장";
  return "<span>📅 Google Calendar</span> + <span>✅ Todoist</span> 모두 저장";
}

function updatePreview() {
  const val = textarea.value;
  targetPreview.innerHTML = val.trim() ? targetLabel(parseTarget(val)) : "";
}

// ---- Keyboard shortcut: Cmd/Ctrl+Enter ----
function handleKeydown(e) {
  if ((e.metaKey || e.ctrlKey) && e.key === "Enter") {
    e.preventDefault();
    handleSubmit();
  }
}

// ---- Submit ----
async function handleSubmit() {
  const text = textarea.value.trim();
  if (!text) {
    showToast("⚠️ 내용을 입력해주세요.", "warn");
    textarea.focus();
    return;
  }

  const target = parseTarget(textarea.value);

  // Quick auth check
  if (target === "calendar" && !state.google) {
    showToast("⚠️ Google Calendar 연결이 필요합니다.", "warn");
    return;
  }
  if (target === "todoist" && !state.todoist) {
    showToast("⚠️ Todoist 연결이 필요합니다.", "warn");
    return;
  }
  if (target === "both" && !state.google && !state.todoist) {
    showToast("⚠️ 먼저 서비스를 연결해주세요.", "warn");
    return;
  }

  setLoading(true);

  try {
    const res  = await fetch("/api/save", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: textarea.value }),
    });
    const data = await res.json();

    if (!res.ok) {
      showToast(`❌ ${data.error || "저장 실패"}`, "error");
      return;
    }

    // Show partial warnings
    const warnings = data.errors || [];
    if (warnings.length) {
      showToast(`⚠️ 일부 저장 실패: ${warnings.join(", ")}`, "warn");
    } else {
      showToast("✅ 저장되었습니다!", "success");
    }

    // Add to history
    addHistory(text, data.target, data.results);

    textarea.value = "";
    updatePreview();

  } catch (e) {
    showToast("❌ 네트워크 오류가 발생했습니다.", "error");
  } finally {
    setLoading(false);
  }
}

// ---- Loading state ----
function setLoading(on) {
  submitBtn.classList.toggle("loading", on);
  submitBtn.disabled = on;
}

// ---- History ----
const recentHistory = [];

function addHistory(text, target, results) {
  recentHistory.unshift({ text, target, results, time: new Date() });
  if (recentHistory.length > 10) recentHistory.pop();
  renderHistory();
}

function renderHistory() {
  if (!historyList) return;

  const section = document.getElementById("history-section");
  if (recentHistory.length > 0 && section) section.style.display = "";

  historyList.innerHTML = recentHistory.map(entry => {
    const icon  = entry.target === "calendar" ? "📅" :
                  entry.target === "todoist"  ? "✅" : "🔗";
    const tags  = buildTags(entry.target, entry.results);
    const timeStr = entry.time.toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit" });

    return `
      <div class="history-item">
        <div class="hi-icon">${icon}</div>
        <div class="hi-body">
          <div class="hi-text">${escHtml(entry.text.replace(/^\/[ct]\s*/, ""))}</div>
          <div class="hi-meta">
            ${tags}
            <span>${timeStr}</span>
          </div>
        </div>
      </div>
    `;
  }).join("");
}

function buildTags(target, results) {
  const tags = [];
  if (results.todoist)  tags.push(`<span class="hi-tag">Todoist ✓</span>`);
  if (results.calendar) tags.push(`<span class="hi-tag">Calendar ✓</span>`);
  return tags.join("");
}

function escHtml(str) {
  return str.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

// ---- Toast ----
function showToast(msg, type = "success") {
  toast.className = `show ${type}`;
  toast.innerHTML = `<span class="toast-icon">${toastIcon(type)}</span><span>${msg}</span>`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { toast.className = ""; }, 3200);
}

function toastIcon(type) {
  if (type === "success") return "✅";
  if (type === "error")   return "❌";
  return "⚠️";
}
