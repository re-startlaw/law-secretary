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
  openShas: [],      // インライン展開中の文書（複数同時オープン＝タブ表示）
  openPages: {},     // sha -> 初期ページ（検索から開いた場合）
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
    row.className = "doc-row" + (state.openShas.includes(d.sha256) ? " open" : "");
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
    row.appendChild(buildMemoCell(d));
    row.appendChild(td(d.page_count == null ? "" : String(d.page_count), "col-page"));

    row.onclick = (e) => {
      if (e.target.closest(".memo-edit")) return;  // メモ編集中はトグルしない
      toggleRow(d, row);
    };
    tbody.appendChild(row);

    if (state.openShas.includes(d.sha256)) {
      const exp = document.createElement("tr");
      exp.className = "viewer-row";
      const cell = document.createElement("td");
      cell.colSpan = COLUMNS.length;
      cell.className = "viewer-cell";
      exp.appendChild(cell);
      tbody.appendChild(exp);
      const initialPage = state.openPages[d.sha256];
      delete state.openPages[d.sha256];
      mountInlineViewer(d, cell, {
        caseId: state.caseId,
        documents: docs,
        initialPage,
        onClose: () => { closeDoc(d.sha256); },
        onNavigate: (sha) => {
          const i = state.openShas.indexOf(d.sha256);
          if (i >= 0) state.openShas[i] = sha; else state.openShas.push(sha);
          renderTabs();
          renderTable();
        },
        onAnnotationsSaved: (sha, has) => {
          const doc = state.documents.find((x) => x.sha256 === sha);
          if (doc) { doc.has_annotations = has; renderEvidenceTabs(); updateRowBadges(); }
        },
        onMetaSaved: (sha, fields) => {
          const doc = state.documents.find((x) => x.sha256 === sha);
          if (doc) Object.assign(doc, fields);
        },
      });
    }
  }
}

function toggleRow(d, row) {
  const i = state.openShas.indexOf(d.sha256);
  if (i >= 0) state.openShas.splice(i, 1);
  else state.openShas.push(d.sha256);
  renderTabs();
  renderTable();
  if (state.openShas.includes(d.sha256)) {
    const el = document.querySelector(`tr.doc-row[data-sha="${d.sha256}"]`);
    if (el) el.scrollIntoView({ block: "nearest" });
  }
}

function closeDoc(sha) {
  const i = state.openShas.indexOf(sha);
  if (i >= 0) state.openShas.splice(i, 1);
  renderTabs();
  renderTable();
}

// 開いている文書のタブ表示（複数同時オープン）
function renderTabs() {
  let bar = document.getElementById("open-tabs");
  if (!bar) {
    bar = document.createElement("div");
    bar.id = "open-tabs";
    bar.className = "open-tabs";
    const wrap = document.getElementById("docs-table-wrap");
    wrap.parentNode.insertBefore(bar, wrap);
  }
  bar.innerHTML = "";
  if (!state.openShas.length) { bar.style.display = "none"; return; }
  bar.style.display = "";
  const lbl = document.createElement("span");
  lbl.className = "tabs-label";
  lbl.textContent = "タブ表示:";
  bar.appendChild(lbl);
  for (const sha of state.openShas) {
    const d = state.documents.find((x) => x.sha256 === sha);
    if (!d) continue;
    const tab = document.createElement("span");
    tab.className = "open-tab";
    tab.textContent = `${d.evidence_no || "—"} ${d.title}`;
    tab.title = "クリックで該当文書へスクロール";
    tab.onclick = () => {
      const el = document.querySelector(`tr.doc-row[data-sha="${sha}"]`);
      if (el) el.scrollIntoView({ block: "center" });
    };
    const x = document.createElement("button");
    x.className = "tab-close";
    x.textContent = "×";
    x.onclick = (e) => { e.stopPropagation(); closeDoc(sha); };
    tab.appendChild(x);
    bar.appendChild(tab);
  }
}

// メモ列インライン編集（ダブルクリックで編集→PUT /meta）
function buildMemoCell(d) {
  const cell = td("", "col-memo");
  const span = document.createElement("span");
  span.className = "memo-text";
  span.textContent = d.memo || "";
  cell.appendChild(span);
  cell.title = "ダブルクリックで編集";
  cell.ondblclick = (e) => {
    e.stopPropagation();
    const input = document.createElement("input");
    input.className = "memo-edit";
    input.value = d.memo || "";
    cell.innerHTML = "";
    cell.appendChild(input);
    input.focus();
    const commit = async () => {
      const val = input.value;
      d.memo = val;
      cell.innerHTML = "";
      const s = document.createElement("span");
      s.className = "memo-text";
      s.textContent = val;
      cell.appendChild(s);
      await saveMeta(d.sha256, { memo: val });
    };
    input.onblur = commit;
    input.onkeydown = (ev) => {
      if (ev.key === "Enter") input.blur();
      else if (ev.key === "Escape") { cell.innerHTML = ""; cell.appendChild(span); }
    };
  };
  return cell;
}

async function saveMeta(sha, fields) {
  try {
    await fetch(`/api/cases/${encodeURIComponent(state.caseId)}/meta/${sha}`, {
      method: "PUT",
      headers: { "X-Kiroku-Viewer": "1", "Content-Type": "application/json" },
      body: JSON.stringify(fields),
    });
  } catch (e) { /* noop */ }
}

// 注釈バッジだけを最小更新（フル再描画を避ける）
function updateRowBadges() {
  for (const d of state.documents) {
    const row = document.querySelector(`tr.doc-row[data-sha="${d.sha256}"] .col-title`);
    if (!row) continue;
    const has = row.querySelector(".badge-anno");
    if (d.has_annotations && !has) row.appendChild(badge("注釈", "badge-anno"));
    else if (!d.has_annotations && has) has.remove();
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
  state.openShas = [];
  state.openPages = {};

  // 現事件の reindex/has_index を取得してボタン・案内バーを更新。
  const picker = document.getElementById("case-picker");
  const cRes = await fetch("/api/cases");
  const cData = await cRes.json();
  const currentCase = (cData.cases || []).find((c) => c.id === state.caseId);
  setupReindexBtn(currentCase);
  showNoIndexBar(currentCase);

  // リロード後に索引作成中だった場合はポーリングを再開。
  if (currentCase && currentCase.reindex) {
    try {
      const rRes = await fetch(`/api/cases/${encodeURIComponent(state.caseId)}/reindex`);
      const rData = await rRes.json();
      if (rData.running && !_reindexPollTimer) {
        showReindexStatus("索引作成中…（ブラウザを閉じても構いません）", "");
        updateReindexBtn(true);
        _reindexPollTimer = setTimeout(pollReindex, 3000);
      }
    } catch (_e) { /* noop */ }
  }

  renderTabs();
  renderEvidenceTabs();
  renderTable();
  checkOrphans();
}

// 孤児注釈（documents に無い sha の注釈）を検出して案内バーを出す
async function checkOrphans() {
  let bar = document.getElementById("orphan-bar");
  try {
    const res = await fetch(`/api/cases/${encodeURIComponent(state.caseId)}/orphan-annotations`);
    const data = await res.json();
    const orphans = data.orphans || [];
    if (!orphans.length) { if (bar) bar.remove(); return; }
    if (!bar) {
      bar = document.createElement("div");
      bar.id = "orphan-bar";
      bar.className = "orphan-bar";
      const wrap = document.getElementById("docs-table-wrap");
      wrap.parentNode.insertBefore(bar, wrap);
    }
    const total = orphans.reduce((s, o) => s + o.count, 0);
    bar.innerHTML = `⚠ 紐付け先が見つからない注釈が ${orphans.length} 文書分（計${total}件）あります。` +
      `ファイル名変更や差し替えの可能性。`;
    const btn = document.createElement("button");
    btn.textContent = "紐付け直す…";
    btn.onclick = () => openRelinkDialog(orphans);
    bar.appendChild(btn);
  } catch (e) { if (bar) bar.remove(); }
}

function openRelinkDialog(orphans) {
  const o = orphans[0];
  const docList = state.documents
    .map((d, i) => `${i + 1}: ${d.evidence_no || "—"} ${d.title}`)
    .join("\n");
  const ans = prompt(
    `孤児注釈（${o.count}件, sha=${o.sha256.slice(0, 12)}…）の紐付け先を番号で選択:\n\n${docList}`
  );
  const idx = parseInt(ans, 10);
  if (!idx || idx < 1 || idx > state.documents.length) return;
  const target = state.documents[idx - 1];
  relink(o.sha256, target.sha256);
}

async function relink(sha, targetSha) {
  try {
    const res = await fetch(`/api/cases/${encodeURIComponent(state.caseId)}/annotations/${sha}/relink`, {
      method: "POST",
      headers: { "X-Kiroku-Viewer": "1", "Content-Type": "application/json" },
      body: JSON.stringify({ target_sha: targetSha }),
    });
    if (!res.ok) { alert("紐付けに失敗しました"); return; }
    await loadDocuments();
  } catch (e) { alert("紐付けエラー: " + e); }
}

// ---- モーダルシステム --------------------------------------------------
let _activeModal = null;
function openModal(content) {
  closeModal();
  const backdrop = document.createElement("div");
  backdrop.className = "modal-backdrop";
  backdrop.onclick = (e) => { if (e.target === backdrop) closeModal(); };
  const modal = document.createElement("div");
  modal.className = "modal";
  modal.appendChild(content);
  backdrop.appendChild(modal);
  document.getElementById("modal-root").appendChild(backdrop);
  _activeModal = backdrop;
}
function closeModal() {
  if (_activeModal) { _activeModal.remove(); _activeModal = null; }
}

// ---- 事件ピッカー再構築 -----------------------------------------------
function rebuildPicker(cases) {
  const picker = document.getElementById("case-picker");
  picker.innerHTML = "";
  for (const c of cases) {
    const opt = document.createElement("option");
    opt.value = c.id;
    opt.textContent = c.name;
    picker.appendChild(opt);
  }
}

// ---- 再索引ボタン・ポーリング ------------------------------------------
let _reindexPollTimer = null;
let _currentCaseMeta = null;  // {reindex, has_index} of current case

function showReindexStatus(msg, kind) {
  let el = document.getElementById("reindex-status");
  if (!el) {
    el = document.createElement("span");
    el.id = "reindex-status";
    el.className = "reindex-status";
    const toolbar = document.querySelector(".docs-toolbar");
    if (toolbar) toolbar.appendChild(el);
  }
  el.textContent = msg;
  el.className = "reindex-status" + (kind ? " " + kind : "");
  if (kind === "success" || kind === "error") {
    setTimeout(() => { if (el) el.remove(); }, 5000);
  }
}

function clearReindexStatus() {
  const el = document.getElementById("reindex-status");
  if (el) el.remove();
}

async function pollReindex() {
  try {
    const res = await fetch(`/api/cases/${encodeURIComponent(state.caseId)}/reindex`);
    const data = await res.json();
    if (!data.running) {
      clearTimeout(_reindexPollTimer);
      _reindexPollTimer = null;
      updateReindexBtn(false);
      if (data.ran) {
        if (data.returncode === 0) {
          showReindexStatus("索引作成完了", "success");
          if (_currentCaseMeta) _currentCaseMeta.has_index = true;
          removeNoIndexBar();
          await loadDocuments();
        } else {
          // 失敗: ログを取得して表示
          try {
            const lr = await fetch(`/api/cases/${encodeURIComponent(state.caseId)}/reindex/log`);
            const ld = await lr.json();
            showReindexStatus(`索引作成に失敗しました（終了コード ${data.returncode}）:\n${ld.log || ""}`, "error");
          } catch (_e) {
            showReindexStatus(`索引作成に失敗しました（終了コード ${data.returncode}）`, "error");
          }
        }
      }
      return;
    }
    showReindexStatus("索引作成中…（ブラウザを閉じても構いません）", "");
    updateReindexBtn(true);
    _reindexPollTimer = setTimeout(pollReindex, 3000);
  } catch (_e) {
    _reindexPollTimer = setTimeout(pollReindex, 3000);
  }
}

function updateReindexBtn(running) {
  const btn = document.getElementById("reindex-btn");
  if (!btn) return;
  btn.disabled = running;
  btn.textContent = running ? "索引作成中…" : "再索引";
}

async function startReindex() {
  try {
    const res = await fetch(`/api/cases/${encodeURIComponent(state.caseId)}/reindex`, {
      method: "POST", headers: { "X-Kiroku-Viewer": "1" },
    });
    if (res.status === 409) {
      // 既に実行中 → ポーリングのみ
    } else if (!res.ok) {
      const d = await res.json().catch(() => ({}));
      alert(d.detail || "索引作成に失敗しました");
      return;
    }
    showReindexStatus("索引作成中…（ブラウザを閉じても構いません）", "");
    updateReindexBtn(true);
    if (!_reindexPollTimer) _reindexPollTimer = setTimeout(pollReindex, 3000);
  } catch (e) {
    alert("索引作成エラー: " + e);
  }
}

function setupReindexBtn(caseMeta) {
  _currentCaseMeta = caseMeta;
  // 既存ボタンを削除してから再生成。
  const old = document.getElementById("reindex-btn");
  if (old) old.remove();
  if (!caseMeta || !caseMeta.reindex) return;
  const btn = document.createElement("button");
  btn.id = "reindex-btn";
  btn.className = "tb-btn";
  btn.textContent = "再索引";
  btn.onclick = () => startReindex();
  const toolbar = document.querySelector(".docs-toolbar");
  const countEl = document.getElementById("doc-count");
  if (toolbar && countEl) toolbar.insertBefore(btn, countEl);
}

function removeNoIndexBar() {
  const bar = document.getElementById("no-index-bar");
  if (bar) bar.remove();
}

function showNoIndexBar(caseMeta) {
  removeNoIndexBar();
  if (!caseMeta || caseMeta.has_index) return;
  const bar = document.createElement("div");
  bar.id = "no-index-bar";
  bar.className = "no-index-bar";
  bar.textContent = "この事件は未索引です（検索不可・閲覧のみ）。";
  if (caseMeta.reindex) {
    const btn = document.createElement("button");
    btn.textContent = "索引を作成";
    btn.onclick = () => startReindex();
    bar.appendChild(btn);
  }
  const wrap = document.getElementById("docs-table-wrap");
  if (wrap) wrap.parentNode.insertBefore(bar, wrap);
}

// ---- 事件追加モーダル --------------------------------------------------
async function openAddCaseModal() {
  const frag = document.createDocumentFragment();

  const h2 = document.createElement("h2");
  h2.textContent = "事件を追加";
  frag.appendChild(h2);

  // 事件名
  const nameRow = document.createElement("div");
  nameRow.className = "modal-row";
  const nameLbl = document.createElement("label");
  nameLbl.textContent = "事件名";
  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.placeholder = "空欄の場合はフォルダ名から自動設定";
  nameRow.append(nameLbl, nameInput);
  frag.appendChild(nameRow);

  // フォルダパス
  const pathRow = document.createElement("div");
  pathRow.className = "modal-row";
  const pathLbl = document.createElement("label");
  pathLbl.textContent = "フォルダパス";
  const pathInput = document.createElement("input");
  pathInput.type = "text";
  pathInput.placeholder = "/path/to/__Document__";
  const pickBtn = document.createElement("button");
  pickBtn.className = "tb-btn";
  pickBtn.textContent = "フォルダを選択…";
  pathRow.append(pathLbl, pathInput, pickBtn);
  frag.appendChild(pathRow);

  const pickHint = document.createElement("div");
  pickHint.className = "modal-hint";
  pickHint.textContent = "";
  frag.appendChild(pickHint);

  // ☑ 即時索引
  const idxRow = document.createElement("div");
  idxRow.className = "modal-row";
  const idxChk = document.createElement("input");
  idxChk.type = "checkbox";
  idxChk.checked = true;
  const idxLbl = document.createElement("label");
  idxLbl.style.flex = "1";
  idxLbl.textContent = "追加後すぐ索引を作成（OCR・件数によっては数十分〜数時間かかります）";
  idxLbl.prepend(idxChk);
  idxRow.appendChild(idxLbl);
  frag.appendChild(idxRow);

  // ☑ reindex 許可
  const reindexRow = document.createElement("div");
  reindexRow.className = "modal-row";
  const reindexChk = document.createElement("input");
  reindexChk.type = "checkbox";
  reindexChk.checked = true;
  const reindexLbl = document.createElement("label");
  reindexLbl.style.flex = "1";
  reindexLbl.textContent = "このMacで索引の作成・更新を許可する（通常は米谷のMacのみON。佐藤さんのMacではOFF）";
  reindexLbl.prepend(reindexChk);
  reindexRow.appendChild(reindexLbl);
  frag.appendChild(reindexRow);

  const errEl = document.createElement("div");
  errEl.className = "modal-error";
  errEl.style.display = "none";
  frag.appendChild(errEl);

  const actions = document.createElement("div");
  actions.className = "modal-actions";
  const cancelBtn = document.createElement("button");
  cancelBtn.textContent = "キャンセル";
  cancelBtn.onclick = () => closeModal();
  const addBtn = document.createElement("button");
  addBtn.className = "primary";
  addBtn.textContent = "追加";
  actions.append(cancelBtn, addBtn);
  frag.appendChild(actions);

  // フォルダ選択ボタン
  pickBtn.onclick = async () => {
    pickBtn.disabled = true;
    pickHint.textContent = "Finderのダイアログを確認してください（他のウィンドウの背面に出ることがあります）";
    try {
      const res = await fetch("/api/pick-folder", {
        method: "POST", headers: { "X-Kiroku-Viewer": "1" },
      });
      if (res.status === 409) {
        pickHint.textContent = "フォルダ選択ダイアログが既に開いています。";
        return;
      }
      const data = await res.json();
      if (data.cancelled) {
        pickHint.textContent = "フォルダ選択がキャンセルされました。";
      } else {
        pathInput.value = data.path || "";
        pickHint.textContent = "";
      }
    } catch (e) {
      pickHint.textContent = "エラー: " + e;
    } finally {
      pickBtn.disabled = false;
    }
  };

  // 追加ボタン
  addBtn.onclick = async () => {
    errEl.style.display = "none";
    addBtn.disabled = true;
    try {
      const res = await fetch("/api/cases", {
        method: "POST",
        headers: { "X-Kiroku-Viewer": "1", "Content-Type": "application/json" },
        body: JSON.stringify({
          name: nameInput.value.trim(),
          path: pathInput.value.trim(),
          reindex: reindexChk.checked,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        errEl.textContent = data.detail || "エラーが発生しました";
        errEl.style.display = "";
        return;
      }
      closeModal();
      // ピッカー再構築 → 新事件に切替
      const cRes = await fetch("/api/cases");
      const cData = await cRes.json();
      rebuildPicker(cData.cases || []);
      const picker = document.getElementById("case-picker");
      picker.value = data.id;
      state.caseId = data.id;
      await loadDocuments();
      // 索引チェック & 即時索引
      if (idxChk.checked && !data.has_index && data.reindex) {
        await startReindex();
      }
    } catch (e) {
      errEl.textContent = "通信エラー: " + e;
      errEl.style.display = "";
    } finally {
      addBtn.disabled = false;
    }
  };

  const wrapper = document.createElement("div");
  wrapper.appendChild(frag);
  openModal(wrapper);
  nameInput.focus();
}

// ---- 事件管理モーダル --------------------------------------------------
async function openManageCasesModal() {
  let casesData = [];
  try {
    const res = await fetch("/api/cases");
    const data = await res.json();
    casesData = data.cases || [];
  } catch (e) {
    alert("事件一覧の取得に失敗しました: " + e);
    return;
  }

  const renderManageModal = () => {
    const wrapper = document.createElement("div");

    const h2 = document.createElement("h2");
    h2.textContent = "事件の管理";
    wrapper.appendChild(h2);

    if (!casesData.length) {
      const p = document.createElement("div");
      p.className = "modal-hint";
      p.textContent = "登録されている事件はありません。";
      wrapper.appendChild(p);
    } else {
      const tbl = document.createElement("table");
      tbl.className = "case-list";
      const thead = tbl.createTHead();
      const hRow = thead.insertRow();
      for (const h of ["事件名", "パス", "索引", "索引許可", "操作"]) {
        const th = document.createElement("th");
        th.textContent = h;
        hRow.appendChild(th);
      }
      const tbody = tbl.createTBody();
      for (const c of casesData) {
        const tr = tbody.insertRow();
        tr.insertCell().textContent = c.name;
        const pathCell = tr.insertCell();
        pathCell.textContent = c.path;
        pathCell.style.cssText = "max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px;color:var(--muted)";
        tr.insertCell().textContent = c.has_index ? "あり" : "なし";
        tr.insertCell().textContent = c.reindex ? "ON" : "OFF";
        const opCell = tr.insertCell();
        const delBtn = document.createElement("button");
        delBtn.textContent = "削除";
        delBtn.onclick = () => confirmDelete(c);
        opCell.appendChild(delBtn);
      }
      wrapper.appendChild(tbl);
    }

    const note = document.createElement("div");
    note.className = "modal-hint";
    note.textContent = "削除はアプリからの登録解除のみです。フォルダ・索引・注釈ファイルは削除されません。";
    wrapper.appendChild(note);

    const actions = document.createElement("div");
    actions.className = "modal-actions";
    const closeBtn = document.createElement("button");
    closeBtn.textContent = "閉じる";
    closeBtn.onclick = () => closeModal();
    actions.appendChild(closeBtn);
    wrapper.appendChild(actions);

    return wrapper;
  };

  const confirmDelete = (c) => {
    const wrapper = document.createElement("div");
    const h2 = document.createElement("h2");
    h2.textContent = "事件の登録を解除";
    wrapper.appendChild(h2);
    const msg = document.createElement("div");
    msg.innerHTML = `「<strong>${escapeHtml(c.name)}</strong>」の登録を解除します。<br>` +
      "<strong>フォルダ・索引・注釈ファイルは削除されません。</strong>";
    wrapper.appendChild(msg);

    const errEl = document.createElement("div");
    errEl.className = "modal-error";
    errEl.style.display = "none";
    wrapper.appendChild(errEl);

    const actions = document.createElement("div");
    actions.className = "modal-actions";
    const backBtn = document.createElement("button");
    backBtn.textContent = "戻る";
    backBtn.onclick = () => openModal(renderManageModal());
    const okBtn = document.createElement("button");
    okBtn.className = "primary";
    okBtn.textContent = "登録解除";
    okBtn.onclick = async () => {
      okBtn.disabled = true;
      try {
        const res = await fetch(`/api/cases/${encodeURIComponent(c.id)}`, {
          method: "DELETE", headers: { "X-Kiroku-Viewer": "1" },
        });
        const data = await res.json();
        if (!res.ok) {
          errEl.textContent = data.detail || "削除に失敗しました";
          errEl.style.display = "";
          okBtn.disabled = false;
          return;
        }
        // ピッカー再構築
        const cRes = await fetch("/api/cases");
        const cData = await cRes.json();
        const newCases = cData.cases || [];
        rebuildPicker(newCases);
        casesData = newCases;
        if (!newCases.length) {
          closeModal();
          openAddCaseModal();
        } else {
          const picker = document.getElementById("case-picker");
          // 削除した事件が表示中なら先頭へ切替
          if (state.caseId === c.id) {
            state.caseId = newCases[0].id;
            picker.value = state.caseId;
            await loadDocuments();
          }
          openModal(renderManageModal());
        }
      } catch (e) {
        errEl.textContent = "通信エラー: " + e;
        errEl.style.display = "";
        okBtn.disabled = false;
      }
    };
    actions.append(backBtn, okBtn);
    wrapper.appendChild(actions);
    openModal(wrapper);
  };

  openModal(renderManageModal());
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

  // ＋追加・⚙管理 ボタン
  document.getElementById("add-case-btn").onclick = () => openAddCaseModal();
  document.getElementById("manage-case-btn").onclick = () => openManageCasesModal();

  const picker = document.getElementById("case-picker");
  picker.onchange = () => {
    state.caseId = picker.value;
    // ポーリングは事件切替でリセット
    if (_reindexPollTimer) { clearTimeout(_reindexPollTimer); _reindexPollTimer = null; }
    clearReindexStatus();
    loadDocuments();
  };

  try {
    const res = await fetch("/api/cases");
    const data = await res.json();
    state.user = data.user_display || data.user || "";
    document.getElementById("user-badge").textContent = state.user ? `👤 ${state.user}` : "";
    const cases = data.cases || [];
    rebuildPicker(cases);
    if (cases.length) {
      state.caseId = cases[0].id;
      await loadDocuments();
    } else {
      // 0件 → 追加モーダルを自動で開く
      document.querySelector("#docs-table tbody").innerHTML =
        '<tr><td class="placeholder">事件が登録されていません。「＋ 追加」から登録してください。</td></tr>';
      openAddCaseModal();
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
window.__openDoc = (sha, page) => {
  const d = state.documents.find((x) => x.sha256 === sha);
  if (!d) return;
  document.querySelector('.main-tab[data-tab="docs"]').click();
  if (!state.openShas.includes(sha)) state.openShas.push(sha);
  if (page) state.openPages[sha] = page;
  state.evidenceTab = "all";
  renderTabs();
  renderEvidenceTabs();
  renderTable();
};

// ---- テキスト検索タブ -------------------------------------------------
let searchTabBuilt = false;
function buildSearchTab() {
  const panel = document.getElementById("tab-search");
  panel.innerHTML = "";
  panel.classList.add("search-panel");

  const bar = document.createElement("div");
  bar.className = "search-bar";
  const q = document.createElement("input");
  q.type = "search";
  q.id = "ts-q";
  q.placeholder = "本文キーワード（スペース区切り＝OR検索）";
  q.onkeydown = (e) => { if (e.key === "Enter") runSearch(); };
  const filt = document.createElement("input");
  filt.type = "search";
  filt.id = "ts-filter";
  filt.placeholder = "絞り込み（タイトル・符号）";
  const run = document.createElement("button");
  run.textContent = "検索実行";
  run.className = "primary";
  run.onclick = runSearch;
  const clear = document.createElement("button");
  clear.textContent = "× クリア";
  clear.onclick = () => {
    q.value = ""; filt.value = "";
    document.getElementById("ts-results").innerHTML = "";
    document.getElementById("ts-status").textContent = "";
  };
  bar.append(q, filt, run, clear);
  panel.appendChild(bar);

  const help = document.createElement("div");
  help.className = "search-help";
  help.textContent = "スペース区切りはOR検索。スニペットは正規化テキスト由来です（原文はビューアで確認）。";
  panel.appendChild(help);

  const status = document.createElement("div");
  status.id = "ts-status";
  status.className = "search-status";
  panel.appendChild(status);

  const results = document.createElement("div");
  results.id = "ts-results";
  results.className = "search-results";
  panel.appendChild(results);
}

async function runSearch() {
  const q = document.getElementById("ts-q").value.trim();
  const filter = document.getElementById("ts-filter").value.trim();
  const status = document.getElementById("ts-status");
  const results = document.getElementById("ts-results");
  results.innerHTML = "";
  if (!q) { status.textContent = "キーワードを入力してください。"; return; }
  status.textContent = "検索中…";
  try {
    const url = `/api/cases/${encodeURIComponent(state.caseId)}/search?` +
      new URLSearchParams({ q, filter });
    const res = await fetch(url);
    const data = await res.json();
    const hits = data.hits || [];
    if (!data.indexed) {
      status.textContent = "この事件は未索引です。索引を生成してください。";
      return;
    }
    if (!hits.length) {
      const lowOcr = state.documents.some((d) => d.ocr_low_confidence);
      status.textContent = "0 件" + (lowOcr ? "（OCR低信頼の文書があります。原文の表記揺れで未ヒットの可能性）" : "");
      return;
    }
    status.textContent = `${hits.length} 件`;
    for (const h of hits) {
      const card = document.createElement("div");
      card.className = "result-card";
      card.innerHTML =
        `<div class="rc-head"><strong>${escapeHtml(h.evidence_no || "—")}</strong> ` +
        `${escapeHtml(h.title)} <span class="rc-page">p.${h.page_no}</span></div>` +
        `<div class="rc-snippet">${highlightSnippet(h.snippet, q)}</div>`;
      card.onclick = () => window.__openDoc(h.sha256, h.page_no);
      results.appendChild(card);
    }
  } catch (e) {
    status.textContent = "検索エラー: " + e;
  }
}

function highlightSnippet(snippet, query) {
  let s = escapeHtml(snippet || "");
  if (s.includes("[") && s.includes("]")) {
    // FTS snippet のマーカー [..] を強調表示に変換。
    s = s.replace(/\[/g, "<mark>").replace(/\]/g, "</mark>");
  } else {
    // LIKE フォールバックはクエリ語を強調。
    for (const tok of query.split(/\s+/).filter(Boolean)) {
      const re = new RegExp("(" + tok.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")", "gi");
      s = s.replace(re, "<mark>$1</mark>");
    }
  }
  return s;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

window.__initSearchTab = () => {
  if (!searchTabBuilt) { buildSearchTab(); searchTabBuilt = true; }
  document.getElementById("ts-q").focus();
};

boot();
