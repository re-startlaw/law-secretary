// 記録ビューア フロントエンド。
// フェーズ1-B: 文書DB表（符号タブ・全列ソート・自然順・和暦・絞り込み・件数・
// 未索引バッジ・OCR低信頼マーク）。行クリックでインラインビューア（1-C）を展開。

import { mountInlineViewer } from "/static/viewer.js";

// ---- 状態 --------------------------------------------------------------
const state = {
  caseId: null,
  user: "",
  documents: [],     // サーバ由来＋表示用 id 付与
  evidenceTab: "all",
  sortCol: "evidence_no",
  sortDir: 1,        // 1=昇順, -1=降順
  filter: "",
  annoOnly: false,
  wareki: false,
  openSha: null,     // インライン展開中の文書
};

const EVIDENCE_TABS = [
  { key: "all", label: "全件" },
  { key: "甲", label: "甲号証" },
  { key: "乙", label: "乙号証" },
  { key: "弁", label: "弁号証" },
  { key: "訴訟書類", label: "訴訟書類" },
  { key: "資料", label: "資料" },
  { key: "none", label: "符号無し" },
];

const COLUMNS = [
  { key: "id", label: "ID", sortable: true, cls: "col-id" },
  { key: "evidence_no", label: "符号", sortable: true, cls: "col-ev" },
  { key: "title", label: "タイトル", sortable: true, cls: "col-title" },
  { key: "document_date", label: "日付", sortable: true, cls: "col-date" },
  { key: "author", label: "作成者", sortable: true, cls: "col-author" },
  { key: "memo", label: "メモ", sortable: true, cls: "col-memo" },
  { key: "page_count", label: "page", sortable: true, cls: "col-page" },
];

// ---- 符号の自然順キー（server の evidence_sort_key と一致） -------------
const KIND_ORDER = { "甲": 0, "乙": 1, "弁": 2 };
function toHalf(s) {
  return (s || "").replace(/[０-９]/g, (c) =>
    String.fromCharCode(c.charCodeAt(0) - 0xfee0)
  );
}
function evidenceSortKey(ev) {
  if (!ev) return [9, "", 0, 0];
  const s = toHalf(ev);
  const kind = s[0] || "";
  const kindOrder = kind in KIND_ORDER ? KIND_ORDER[kind] : 8;
  const m = s.slice(1).match(/^([A-Za-zＡ-Ｚａ-ｚ]?)(\d+)(?:[-_－―の](\d+))?/);
  if (!m) return [kindOrder, s.slice(1), 0, 0];
  return [kindOrder, m[1] || "", parseInt(m[2] || "0", 10), parseInt(m[3] || "0", 10)];
}

function cmpArr(a, b) {
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) {
    if (a[i] < b[i]) return -1;
    if (a[i] > b[i]) return 1;
  }
  return a.length - b.length;
}

// ---- 和暦変換 ----------------------------------------------------------
function formatDate(iso, wareki) {
  if (!iso) return "";
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!m) return iso;
  const [, y, mo, d] = m;
  const year = parseInt(y, 10);
  if (!wareki) return `${year}/${parseInt(mo, 10)}/${parseInt(d, 10)}`;
  if (year >= 2019) {
    const r = year - 2018;
    const ry = r === 1 ? "元" : String(r);
    return `令和${ry}年${parseInt(mo, 10)}月${parseInt(d, 10)}日`;
  }
  return iso;
}

// ---- 絞り込み・タブ判定 -------------------------------------------------
function matchEvidenceTab(doc, key) {
  if (key === "all") return true;
  if (key === "甲" || key === "乙" || key === "弁") {
    return toHalf(doc.evidence_no).startsWith(key);
  }
  if (key === "none") {
    // 符号無し（訴訟書類/資料 のカテゴリ上書きが無いもの）
    return !doc.evidence_no && !doc.category;
  }
  // 訴訟書類 / 資料 は viewer_meta の手動上書きのみ（フェーズ2）
  return doc.category === key;
}

function visibleDocuments() {
  const f = state.filter.trim().toLowerCase();
  return state.documents.filter((d) => {
    if (!matchEvidenceTab(d, state.evidenceTab)) return false;
    if (state.annoOnly && !d.has_annotations) return false;
    if (f) {
      const hay = `${d.evidence_no} ${d.title} ${d.author} ${d.memo}`.toLowerCase();
      if (!hay.includes(f)) return false;
    }
    return true;
  });
}

function sortDocuments(docs) {
  const { sortCol, sortDir } = state;
  const keyed = docs.map((d) => {
    let k;
    if (sortCol === "evidence_no") k = evidenceSortKey(d.evidence_no);
    else if (sortCol === "page_count" || sortCol === "id")
      k = [d[sortCol] == null ? -1 : d[sortCol]];
    else k = [String(d[sortCol] ?? "")];
    return { d, k };
  });
  keyed.sort((a, b) => sortDir * cmpArr(a.k, b.k));
  return keyed.map((x) => x.d);
}

// ---- レンダリング ------------------------------------------------------
function renderEvidenceTabs() {
  const wrap = document.getElementById("evidence-tabs");
  const counts = {};
  for (const t of EVIDENCE_TABS) {
    counts[t.key] = state.documents.filter((d) => matchEvidenceTab(d, t.key)).length;
  }
  wrap.innerHTML = "";
  for (const t of EVIDENCE_TABS) {
    const b = document.createElement("button");
    b.className = "evidence-tab" + (state.evidenceTab === t.key ? " active" : "");
    b.textContent = `${t.label}（${counts[t.key]}）`;
    b.onclick = () => {
      state.evidenceTab = t.key;
      renderEvidenceTabs();
      renderTable();
    };
    wrap.appendChild(b);
  }
}

function renderTable() {
  const docs = sortDocuments(visibleDocuments());
  document.getElementById("doc-count").textContent = `${docs.length} 件`;

  // ヘッダ
  const thead = document.querySelector("#docs-table thead");
  thead.innerHTML = "";
  const tr = document.createElement("tr");
  for (const c of COLUMNS) {
    const th = document.createElement("th");
    th.className = c.cls;
    th.textContent = c.label;
    if (c.sortable) {
      th.classList.add("sortable");
      if (state.sortCol === c.key) {
        th.classList.add("sorted");
        th.textContent += state.sortDir === 1 ? " ▲" : " ▼";
      }
      th.onclick = () => {
        if (state.sortCol === c.key) state.sortDir *= -1;
        else {
          state.sortCol = c.key;
          state.sortDir = 1;
        }
        renderTable();
      };
    }
    tr.appendChild(th);
  }
  thead.appendChild(tr);

  // 本体
  const tbody = document.querySelector("#docs-table tbody");
  tbody.innerHTML = "";
  for (const d of docs) {
    const row = document.createElement("tr");
    row.className = "doc-row" + (state.openSha === d.sha256 ? " open" : "");
    row.dataset.sha = d.sha256;

    row.appendChild(td(String(d.id), "col-id"));

    const evTd = td("", "col-ev");
    evTd.appendChild(textSpan(d.evidence_no || "—"));
    if (!d.indexed)
      evTd.appendChild(badge(d.kind === "media" ? "メディア" : "未索引", "badge-unindexed"));
    if (d.ocr_low_confidence) evTd.appendChild(badge("OCR低信頼", "badge-ocr"));
    row.appendChild(evTd);

    const titleTd = td("", "col-title");
    titleTd.appendChild(textSpan(d.title));
    if (d.has_annotations) titleTd.appendChild(badge("注釈", "badge-anno"));
    row.appendChild(titleTd);

    row.appendChild(td(formatDate(d.document_date, state.wareki), "col-date"));
    row.appendChild(td(d.author || "", "col-author"));
    row.appendChild(td(d.memo || "", "col-memo"));
    row.appendChild(td(d.page_count == null ? "" : String(d.page_count), "col-page"));

    row.onclick = () => toggleRow(d, row);
    tbody.appendChild(row);

    if (state.openSha === d.sha256) {
      const exp = document.createElement("tr");
      exp.className = "viewer-row";
      const cell = document.createElement("td");
      cell.colSpan = COLUMNS.length;
      cell.className = "viewer-cell";
      exp.appendChild(cell);
      tbody.appendChild(exp);
      mountInlineViewer(d, cell, {
        caseId: state.caseId,
        documents: docs,
        onClose: () => {
          state.openSha = null;
          renderTable();
        },
        onNavigate: (sha) => {
          state.openSha = sha;
          renderTable();
        },
      });
    }
  }
}

function toggleRow(d, row) {
  state.openSha = state.openSha === d.sha256 ? null : d.sha256;
  renderTable();
  if (state.openSha) {
    const el = document.querySelector(`tr.doc-row[data-sha="${d.sha256}"]`);
    if (el) el.scrollIntoView({ block: "nearest" });
  }
}

function td(text, cls) {
  const el = document.createElement("td");
  if (cls) el.className = cls;
  if (text) el.textContent = text;
  return el;
}
function textSpan(text) {
  const s = document.createElement("span");
  s.textContent = text;
  return s;
}
function badge(text, cls) {
  const b = document.createElement("span");
  b.className = "badge " + cls;
  b.textContent = text;
  return b;
}

// ---- データ取得 --------------------------------------------------------
async function loadDocuments() {
  const res = await fetch(`/api/cases/${encodeURIComponent(state.caseId)}/documents`);
  const data = await res.json();
  // 自然順で安定 ID を付与（フィルタに依らない通し番号）。
  const docs = (data.documents || []).slice();
  docs.sort((a, b) => cmpArr(evidenceSortKey(a.evidence_no), evidenceSortKey(b.evidence_no)));
  docs.forEach((d, i) => (d.id = i + 1));
  state.documents = docs;
  state.openSha = null;
  renderEvidenceTabs();
  renderTable();
}

// ---- タブ・コントロール ------------------------------------------------
function setupMainTabs() {
  document.querySelectorAll(".main-tab").forEach((btn) => {
    btn.onclick = () => {
      document.querySelectorAll(".main-tab").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
      if (btn.dataset.tab === "search" && window.__initSearchTab) window.__initSearchTab();
    };
  });
}

function setupControls() {
  const fi = document.getElementById("filter-input");
  fi.oninput = () => {
    state.filter = fi.value;
    renderTable();
  };
  document.getElementById("anno-only").onchange = (e) => {
    state.annoOnly = e.target.checked;
    renderTable();
  };
  document.getElementById("wareki").onchange = (e) => {
    state.wareki = e.target.checked;
    renderTable();
  };
}

async function boot() {
  setupMainTabs();
  setupControls();
  try {
    const res = await fetch("/api/cases");
    const data = await res.json();
    state.user = data.user_display || data.user || "";
    document.getElementById("user-badge").textContent = state.user ? `👤 ${state.user}` : "";
    const cases = data.cases || [];
    const picker = document.getElementById("case-picker");
    picker.innerHTML = "";
    for (const c of cases) {
      const opt = document.createElement("option");
      opt.value = c.id;
      opt.textContent = c.name;
      picker.appendChild(opt);
    }
    picker.onchange = () => {
      state.caseId = picker.value;
      loadDocuments();
    };
    if (cases.length) {
      state.caseId = cases[0].id;
      await loadDocuments();
    } else {
      document.querySelector("#docs-table tbody").innerHTML =
        '<tr><td class="placeholder">cases.json に事件が登録されていません。</td></tr>';
    }
  } catch (e) {
    document.querySelector("#docs-table tbody").innerHTML =
      `<tr><td class="placeholder">読み込みエラー: ${e}</td></tr>`;
  }
}

// 検索タブ（1-C）が状態を参照できるよう公開。
export const appState = state;
export { loadDocuments };
window.__appState = state;
window.__openDoc = (sha) => {
  const d = state.documents.find((x) => x.sha256 === sha);
  if (!d) return;
  document.querySelector('.main-tab[data-tab="docs"]').click();
  state.openSha = sha;
  state.evidenceTab = "all";
  renderEvidenceTabs();
  renderTable();
};

boot();
