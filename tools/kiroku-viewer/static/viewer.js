// インラインPDFビューア（フェーズ1-C＋フェーズ2）。
// 1-C: 仮想化レンダリング・textLayer・FILE/TXT・ズーム・回転・ページ送り/ジャンプ・
//      文書内検索（D10）・DL・前後文書・とじる。
// 2:   注釈オーバーレイ（付箋★/矩形/ペン/直線/テキスト/コメント・undo/redo/削除、
//      D8座標・自動保存D9）、他ユーザー注釈の色分け表示（選択不可）、丁数オフセット、
//      回転状態の永続化、注釈込みPDF書き出し（墨消し警告D13）。

import * as pdfjsLib from "/static/vendor/pdfjs/build/pdf.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = "/static/vendor/pdfjs/build/pdf.worker.mjs";
const PDFJS_BASE = "/static/vendor/pdfjs/";
const RENDER_BUFFER = 2;
const MAX_CONCURRENT = 2;
const SAVE_DEBOUNCE_MS = 2000;   // 注釈PUTのデバウンス（D9: 2秒以上）
const CSRF = { "X-Kiroku-Viewer": "1" };

const TOOLS = [
  { key: "hand", label: "ハンド" },
  { key: "note", label: "付箋★" },
  { key: "rect", label: "矩形" },
  { key: "pen", label: "ペン" },
  { key: "line", label: "直線" },
  { key: "text", label: "テキスト" },
  { key: "comment", label: "コメント" },
];
const COLORS = ["#dc2626", "#2563eb", "#16a34a", "#d97706", "#000000"];

function normalize(s) { return (s || "").normalize("NFKC").toLowerCase(); }
function uuid() {
  return (crypto.randomUUID && crypto.randomUUID()) ||
    "id-" + Math.floor(performance.now() * 1000).toString(36) + Math.floor(Math.random() * 1e6).toString(36);
}

export function mountInlineViewer(doc, container, ctx) {
  container.innerHTML = "";
  const root = document.createElement("div");
  root.className = "inline-viewer";
  container.appendChild(root);

  if (doc.kind === "media") {
    root.appendChild(buildBar(doc, ctx));
    const body = document.createElement("div");
    body.className = "viewer-body media-body";
    body.innerHTML = `<p>メディアファイル（${escapeHtml(doc.file_name)}）。ブラウザでは再生しません。</p>`;
    const open = document.createElement("button");
    open.textContent = "QuickTime等で開く";
    open.onclick = () => openExternal(ctx.caseId, doc.sha256);
    body.appendChild(open);
    root.appendChild(body);
    return null;
  }
  return new ViewerInstance(doc, root, ctx);
}

function buildBar(doc, ctx) {
  const bar = document.createElement("div");
  bar.className = "viewer-bar";
  const title = document.createElement("span");
  title.className = "vb-title";
  title.innerHTML = `<strong>${escapeHtml(doc.evidence_no || "—")}</strong> ${escapeHtml(doc.title)}`;
  bar.appendChild(title);
  if (!doc.indexed) {
    const b = document.createElement("span");
    b.className = "badge badge-unindexed";
    b.textContent = "未索引（検索不可・閲覧可）";
    bar.appendChild(b);
  }
  const close = document.createElement("button");
  close.textContent = "とじる";
  close.className = "viewer-close";
  close.onclick = () => ctx.onClose();
  bar.appendChild(close);
  return bar;
}

class ViewerInstance {
  constructor(doc, root, ctx) {
    this.doc = doc;
    this.root = root;
    this.ctx = ctx;
    this.scale = 1.0;
    this.fitMode = true;             // true = ウィンドウ幅追従モード
    this._lastFitWidth = 0;          // 無限ループ防止用の前回フィット幅
    this._fitDebounceTimer = null;
    this._resizeObserver = null;
    this.rotation = doc.rotation || 0;     // 保存済み回転状態を復元
    this.choOffset = doc.cho_offset || 0;  // 丁数オフセット
    this.currentPage = 1;
    this.numPages = 0;
    this.pages = [];
    this.inFlight = 0;
    this.renderQueue = [];
    this.mode = "file";
    this.textPages = null;
    this.searchHits = [];
    this.searchIdx = -1;
    this.searchQuery = null;
    this.destroyed = false;
    // 注釈
    this.annoMode = false;
    this.tool = "hand";
    this.color = COLORS[0];
    this.ownAnnotations = [];
    this.otherAnnotations = [];
    this.ownUpdatedAt = "";
    this.selectedId = null;
    this.undoStack = [];
    this.redoStack = [];
    this.saveTimer = null;
    this.build();
    this.load();
  }

  build() {
    this.root.appendChild(buildBar(this.doc, this.ctx));
    const layout = document.createElement("div");
    layout.className = "viewer-layout";

    const side = document.createElement("div");
    side.className = "viewer-side";
    this.fileBtn = sideBtn("FILE", () => this.setMode("file"));
    this.txtBtn = sideBtn("TXT", () => this.setMode("txt"));
    this.fileBtn.classList.add("active");
    side.append(this.fileBtn, this.txtBtn);
    if (!this.ctx.readOnly) {
      this.annoBtn = sideBtn("注釈モード", () => this.toggleAnnoMode());
      side.appendChild(this.annoBtn);
      side.appendChild(sideBtn("書き出し", () => this.exportPdf()));
    }
    side.appendChild(sideBtn("DL", () => this.download()));
    side.appendChild(sideBtn("↑ 前の文書", () => this.navDoc(-1)));
    side.appendChild(sideBtn("↓ 次の文書", () => this.navDoc(1)));
    side.appendChild(sideBtn("とじる", () => this.ctx.onClose()));
    layout.appendChild(side);

    const right = document.createElement("div");
    right.className = "viewer-right";
    right.appendChild(this.buildToolbar());
    if (!this.ctx.readOnly) {
      right.appendChild(this.buildAnnoToolbar());
    }

    this.scrollEl = document.createElement("div");
    this.scrollEl.className = "pdf-scroll";
    this.scrollEl.addEventListener("scroll", () => this.onScroll());
    right.appendChild(this.scrollEl);

    this.txtEl = document.createElement("div");
    this.txtEl.className = "txt-view";
    this.txtEl.style.display = "none";
    right.appendChild(this.txtEl);

    layout.appendChild(right);
    this.root.appendChild(layout);

    this.root.tabIndex = 0;
    this.root.addEventListener("keydown", (e) => this.onKey(e));

    // ResizeObserver でウィンドウ幅追従（課題①）
    if (typeof ResizeObserver !== "undefined") {
      this._resizeObserver = new ResizeObserver(() => {
        if (!this.fitMode || this.numPages === 0) return;
        clearTimeout(this._fitDebounceTimer);
        this._fitDebounceTimer = setTimeout(() => {
          const avail = this.scrollEl.clientWidth - 32;
          if (Math.abs(avail - this._lastFitWidth) <= 2) return; // 2px以下の変化は無視
          this._lastFitWidth = avail;
          this.fitWidth(true);  // silent=true: reflow なし
          this.reflow();
          this.jumpTo(this.currentPage);
        }, 200);
      });
      this._resizeObserver.observe(this.scrollEl);
    }
  }

  buildToolbar() {
    const tb = document.createElement("div");
    tb.className = "viewer-toolbar";

    this.pageInput = document.createElement("input");
    this.pageInput.type = "text";
    this.pageInput.className = "page-input";
    this.pageInput.value = "1";
    this.pageInput.onchange = () => this.jumpTo(parseInt(this.pageInput.value, 10));
    this.pageTotal = document.createElement("span");
    this.pageTotal.className = "page-total";
    const pageGrp = document.createElement("span");
    pageGrp.className = "tb-grp";
    pageGrp.append(tbBtn("‹", () => this.jumpTo(this.currentPage - 1)), this.pageInput,
      this.pageTotal, tbBtn("›", () => this.jumpTo(this.currentPage + 1)));
    tb.appendChild(pageGrp);

    // 丁数オフセット
    this.choInput = document.createElement("input");
    this.choInput.type = "number";
    this.choInput.className = "cho-input";
    this.choInput.value = String(this.choOffset);
    this.choInput.title = "丁数オフセット（PDF頁＋この値＝丁数）";
    this.choInput.onchange = () => this.setChoOffset(parseInt(this.choInput.value, 10) || 0);
    const choGrp = document.createElement("span");
    choGrp.className = "tb-grp";
    choGrp.append(label("丁数+"), this.choInput);
    tb.appendChild(choGrp);

    this.zoomLabel = document.createElement("span");
    this.zoomLabel.className = "zoom-label";
    const zoomGrp = document.createElement("span");
    zoomGrp.className = "tb-grp";
    zoomGrp.append(tbBtn("−", () => this.zoom(-0.2)), this.zoomLabel,
      tbBtn("＋", () => this.zoom(0.2)), tbBtn("100%", () => this.setScale(1.0)),
      tbBtn("幅", () => this.fitWidth()));
    tb.appendChild(zoomGrp);

    tb.appendChild(tbBtn("⟳ 90°", () => this.rotate()));

    this.searchInput = document.createElement("input");
    this.searchInput.type = "search";
    this.searchInput.className = "doc-search";
    this.searchInput.placeholder = this.doc.indexed ? "文書内検索" : "未索引（検索不可）";
    this.searchInput.disabled = !this.doc.indexed;
    this.searchInput.onkeydown = (e) => { if (e.key === "Enter") this.runDocSearch(); };
    this.searchCount = document.createElement("span");
    this.searchCount.className = "search-count";
    const searchGrp = document.createElement("span");
    searchGrp.className = "tb-grp tb-search";
    searchGrp.append(this.searchInput, tbBtn("検索", () => this.runDocSearch()),
      tbBtn("‹", () => this.cycleHit(-1)), tbBtn("›", () => this.cycleHit(1)), this.searchCount);
    tb.appendChild(searchGrp);
    return tb;
  }

  buildAnnoToolbar() {
    const tb = document.createElement("div");
    tb.className = "anno-toolbar";
    tb.style.display = "none";
    this.annoToolbar = tb;

    this.toolBtns = {};
    for (const t of TOOLS) {
      const b = tbBtn(t.label, () => this.setTool(t.key));
      b.classList.add("tool-btn");
      if (t.key === "hand") b.classList.add("active");
      this.toolBtns[t.key] = b;
      tb.appendChild(b);
    }
    // 色
    const colorWrap = document.createElement("span");
    colorWrap.className = "tb-grp color-wrap";
    this.colorBtns = [];
    for (const col of COLORS) {
      const sw = document.createElement("button");
      sw.className = "color-sw";
      sw.style.background = col;
      if (col === this.color) sw.classList.add("active");
      sw.onclick = () => this.setColor(col);
      this.colorBtns.push(sw);
      colorWrap.appendChild(sw);
    }
    tb.appendChild(colorWrap);

    tb.appendChild(tbBtn("元に戻す", () => this.undo()));
    tb.appendChild(tbBtn("やり直し", () => this.redo()));
    tb.appendChild(tbBtn("削除", () => this.deleteSelected()));
    this.annoStatus = document.createElement("span");
    this.annoStatus.className = "anno-status";
    tb.appendChild(this.annoStatus);
    return tb;
  }

  async load() {
    const url = `/api/cases/${encodeURIComponent(this.ctx.caseId)}/pdf/${this.doc.sha256}`;
    try {
      this.loadingTask = pdfjsLib.getDocument({
        url, cMapUrl: PDFJS_BASE + "cmaps/", cMapPacked: true,
        standardFontDataUrl: PDFJS_BASE + "standard_fonts/",
      });
      this.pdf = await this.loadingTask.promise;
    } catch (e) {
      this.scrollEl.innerHTML = `<p class="placeholder">PDFを開けません: ${escapeHtml(String(e))}</p>`;
      return;
    }
    if (this.destroyed) return;
    this.numPages = this.pdf.numPages;
    this.updatePageTotal();
    const first = await this.pdf.getPage(1);
    const base = first.getViewport({ scale: 1, rotation: 0 });
    this.defaultBaseW = base.width;
    this.defaultBaseH = base.height;
    this.fitWidth(true);
    this.buildPlaceholders();
    this.updateVisible();
    if (this.ctx.initialPage && this.ctx.initialPage > 1) this.jumpTo(this.ctx.initialPage);
    if (this.doc.indexed) this.fetchText().then(() => {
      // 検索タブから開いた際に初期クエリを適用してハイライト
      if (this.ctx.initialQuery && this.searchInput) {
        this.searchInput.value = this.ctx.initialQuery;
        this.runDocSearch();
      }
    });
    if (!this.ctx.readOnly) this.loadAnnotations();
  }

  updatePageTotal() {
    const cho = this.choOffset ? `（${this.currentPage + this.choOffset}丁）` : "";
    this.pageTotal.textContent = `/ ${this.numPages}${cho}`;
  }
  choLabel(n) {
    return this.choOffset ? `${n} / ${this.numPages}（${n + this.choOffset}丁）`
      : `${n} / ${this.numPages}`;
  }

  // ---- プレースホルダ・寸法 ----
  buildPlaceholders() {
    this.scrollEl.innerHTML = "";
    this.pages = [];
    for (let i = 1; i <= this.numPages; i++) {
      const div = document.createElement("div");
      div.className = "pdf-page";
      div.dataset.page = i;
      const p = { div, canvas: null, textDiv: null, overlay: null, viewport: null,
        rendered: false, rendering: false, renderTask: null,
        baseW: this.defaultBaseW, baseH: this.defaultBaseH };
      this.setPlaceholderHeight(p);
      const lab = document.createElement("div");
      lab.className = "page-label";
      lab.textContent = i;
      div.appendChild(lab);
      this.scrollEl.appendChild(div);
      this.pages.push(p);
    }
  }
  displayDims(p) {
    const rot = this.rotation % 180 !== 0;
    return { w: (rot ? p.baseH : p.baseW) * this.scale, h: (rot ? p.baseW : p.baseH) * this.scale };
  }
  setPlaceholderHeight(p) {
    const { w, h } = this.displayDims(p);
    p.div.style.width = w + "px";
    p.div.style.height = h + "px";
  }

  // ---- スクロール・可視判定 ----
  onScroll() {
    if (this._scrollRaf) return;
    this._scrollRaf = requestAnimationFrame(() => {
      this._scrollRaf = null;
      this.updateVisible();
      this.updateCurrentPage();
    });
  }
  visibleRange() {
    const top = this.scrollEl.scrollTop, bottom = top + this.scrollEl.clientHeight;
    let first = this.numPages, last = 1;
    for (let i = 0; i < this.pages.length; i++) {
      const div = this.pages[i].div;
      const dTop = div.offsetTop - this.scrollEl.offsetTop, dBottom = dTop + div.offsetHeight;
      if (dBottom >= top && dTop <= bottom) { first = Math.min(first, i + 1); last = Math.max(last, i + 1); }
    }
    if (first > last) { first = this.currentPage; last = this.currentPage; }
    return { first, last };
  }
  updateVisible() {
    const { first, last } = this.visibleRange();
    const lo = Math.max(1, first - RENDER_BUFFER), hi = Math.min(this.numPages, last + RENDER_BUFFER);
    for (let i = 1; i <= this.numPages; i++) {
      const p = this.pages[i - 1];
      if (i >= lo && i <= hi) { if (!p.rendered && !p.rendering) this.enqueue(i); }
      else if (p.rendered || p.rendering) this.unload(i);
    }
  }
  updateCurrentPage() {
    const top = this.scrollEl.scrollTop + this.scrollEl.clientHeight * 0.3;
    let cur = 1;
    for (let i = 0; i < this.pages.length; i++) {
      const dTop = this.pages[i].div.offsetTop - this.scrollEl.offsetTop;
      if (dTop <= top) cur = i + 1; else break;
    }
    if (cur !== this.currentPage) {
      this.currentPage = cur;
      this.pageInput.value = String(cur);
      this.updatePageTotal();
      if (this.mode === "txt") this.renderTxt();
    }
  }

  // ---- レンダリング ----
  enqueue(i) {
    const p = this.pages[i - 1];
    if (p.rendered || p.rendering || p.queued) return;
    p.queued = true; this.renderQueue.push(i); this.pump();
  }
  pump() {
    while (this.inFlight < MAX_CONCURRENT && this.renderQueue.length) {
      const i = this.renderQueue.shift();
      this.pages[i - 1].queued = false;
      this.renderPage(i);
    }
  }
  async renderPage(i) {
    const p = this.pages[i - 1];
    if (p.rendered || p.rendering || this.destroyed) return;
    p.rendering = true; this.inFlight++;
    try {
      const page = await this.pdf.getPage(i);
      if (this.destroyed) return;
      const base = page.getViewport({ scale: 1, rotation: 0 });
      p.baseW = base.width; p.baseH = base.height;
      this.setPlaceholderHeight(p);
      const viewport = page.getViewport({ scale: this.scale, rotation: this.rotation });
      p.viewport = viewport;
      const dpr = window.devicePixelRatio || 1;

      const canvas = document.createElement("canvas");
      canvas.className = "pdf-canvas";
      canvas.width = Math.floor(viewport.width * dpr);
      canvas.height = Math.floor(viewport.height * dpr);
      canvas.style.width = viewport.width + "px";
      canvas.style.height = viewport.height + "px";
      const cctx = canvas.getContext("2d");

      const textDiv = document.createElement("div");
      textDiv.className = "textLayer";
      textDiv.style.width = viewport.width + "px";
      textDiv.style.height = viewport.height + "px";
      textDiv.style.setProperty("--scale-factor", String(viewport.scale));

      const overlay = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      overlay.setAttribute("class", "anno-layer");
      overlay.setAttribute("width", viewport.width);
      overlay.setAttribute("height", viewport.height);
      overlay.style.width = viewport.width + "px";
      overlay.style.height = viewport.height + "px";
      overlay.style.pointerEvents = this.annoMode ? "auto" : "none";
      this.attachOverlayHandlers(overlay, i);

      p.canvas = canvas; p.textDiv = textDiv; p.overlay = overlay;
      p.div.querySelector(".page-label")?.remove();
      p.div.append(canvas, textDiv, overlay);
      this.drawOverlay(i);  // 注釈は描画完了を待たず即時表示

      p.renderTask = page.render({
        canvasContext: cctx, viewport,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      });
      await p.renderTask.promise;
      if (this.destroyed) return;
      const tl = new pdfjsLib.TextLayer({
        textContentSource: page.streamTextContent(), container: textDiv, viewport,
      });
      await tl.render();
      p.rendered = true;
      this.drawOverlay(i);
      this.applyHighlight(p, i);
    } catch (e) {
      if (!(e && e.name === "RenderingCancelledException")) console.warn("render failed", i, e);
    } finally {
      p.rendering = false; p.renderTask = null; this.inFlight--; this.pump();
    }
  }
  unload(i) {
    const p = this.pages[i - 1];
    if (p.renderTask) { try { p.renderTask.cancel(); } catch {} }
    p.canvas?.remove(); p.textDiv?.remove(); p.overlay?.remove();
    p.canvas = p.textDiv = p.overlay = p.viewport = null;
    p.rendered = false; p.rendering = false;
    if (!p.div.querySelector(".page-label")) {
      const lab = document.createElement("div");
      lab.className = "page-label"; lab.textContent = i;
      p.div.appendChild(lab);
    }
  }
  reflow() {
    for (let i = 1; i <= this.numPages; i++) { this.unload(i); this.setPlaceholderHeight(this.pages[i - 1]); }
    this.renderQueue = []; this.updateVisible();
  }

  // ---- ズーム・回転・ジャンプ ----
  setScale(s) {
    this.fitMode = false;  // 手動ズームで追従モードを解除
    this.scale = Math.max(0.2, Math.min(5, s));
    this.zoomLabel.textContent = Math.round(this.scale * 100) + "%";
    this.reflow();
  }
  zoom(d) { this.setScale(this.scale + d); }
  fitWidth(silent) {
    const avail = this.scrollEl.clientWidth - 32;
    const rot = this.rotation % 180 !== 0;
    const baseW = rot ? this.defaultBaseH : this.defaultBaseW;
    const s = avail > 0 && baseW ? avail / baseW : 1.0;
    this.fitMode = true;   // 「幅」ボタンで追従モードを復帰
    this._lastFitWidth = avail;
    // setScale は fitMode=false にするので直接設定する
    this.scale = Math.max(0.2, Math.min(5, s));
    this.zoomLabel.textContent = Math.round(this.scale * 100) + "%";
    if (!silent) this.reflow();
  }
  rotate() {
    this.rotation = (this.rotation + 90) % 360;
    this.reflow();
    this.saveMeta({ rotation: this.rotation });  // 回転状態を永続化
  }
  jumpTo(n) {
    if (!n || isNaN(n)) return;
    n = Math.max(1, Math.min(this.numPages, n));
    const div = this.pages[n - 1].div;
    this.scrollEl.scrollTop = div.offsetTop - this.scrollEl.offsetTop;
    this.currentPage = n; this.pageInput.value = String(n);
    this.updatePageTotal();
    this.updateVisible();
    if (this.mode === "txt") this.renderTxt();
  }
  setChoOffset(v) {
    this.choOffset = v;
    this.updatePageTotal();
    if (this.mode === "txt") this.renderTxt();
    this.saveMeta({ cho_offset: v });
  }

  // ---- FILE/TXT ----
  setMode(mode) {
    this.mode = mode;
    const file = mode === "file";
    this.scrollEl.style.display = file ? "" : "none";
    this.txtEl.style.display = file ? "none" : "";
    this.fileBtn.classList.toggle("active", file);
    this.txtBtn.classList.toggle("active", !file);
    if (!file) this.renderTxt();
  }
  renderTxt() {
    if (!this.doc.indexed) { this.txtEl.innerHTML = '<p class="placeholder">未索引のためテキストはありません。</p>'; return; }
    if (!this.textPages) { this.txtEl.innerHTML = '<p class="placeholder">テキスト取得中…</p>'; return; }
    const pg = this.textPages.find((t) => t.page_no === this.currentPage);
    const head = `<div class="txt-head">${this.choLabel(this.currentPage)}</div>`;
    let body = escapeHtml(pg ? pg.text : "");
    // ヒット語ハイライト（検索タブ・文書内検索共用）
    if (this.searchQuery && pg && pg.text) {
      const escaped = this.searchQuery.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      body = body.replace(new RegExp(escaped, "gi"), (m) => `<mark>${m}</mark>`);
    }
    this.txtEl.innerHTML = head + `<pre class="txt-body">${body}</pre>`;
  }
  async fetchText() {
    try {
      const res = await fetch(`/api/cases/${encodeURIComponent(this.ctx.caseId)}/text/${this.doc.sha256}`);
      this.textPages = (await res.json()).pages || [];
      if (this.mode === "txt") this.renderTxt();
    } catch {}
  }

  // ---- 文書内検索（D10） ----
  async runDocSearch() {
    const q = normalize(this.searchInput.value.trim());
    if (!q) { this.searchHits = []; this.searchIdx = -1; this.searchCount.textContent = ""; this.searchQuery = null; this.reHighlightVisible(); return; }
    if (!this.textPages) await this.fetchText();
    const hits = [];
    for (const t of (this.textPages || [])) if (normalize(t.text).includes(q)) hits.push(t.page_no);
    this.searchHits = hits; this.searchQuery = q;
    this.searchIdx = hits.length ? 0 : -1;
    this.searchCount.textContent = hits.length ? `1/${hits.length}` : "0件";
    this.reHighlightVisible();
    if (hits.length) this.jumpTo(hits[0]);
  }
  cycleHit(d) {
    if (!this.searchHits.length) return;
    this.searchIdx = (this.searchIdx + d + this.searchHits.length) % this.searchHits.length;
    this.searchCount.textContent = `${this.searchIdx + 1}/${this.searchHits.length}`;
    this.jumpTo(this.searchHits[this.searchIdx]);
  }
  reHighlightVisible() {
    for (let i = 1; i <= this.numPages; i++) { const p = this.pages[i - 1]; if (p.rendered) this.applyHighlight(p, i); }
  }
  applyHighlight(p, i) {
    if (!p.div) return;
    p.div.classList.toggle("search-hit-page", !!this.searchQuery && this.searchHits.includes(i));
  }

  // ===================== 注釈（フェーズ2） =====================
  async loadAnnotations() {
    try {
      const res = await fetch(`/api/cases/${encodeURIComponent(this.ctx.caseId)}/annotations/${this.doc.sha256}`);
      const data = await res.json();
      const all = data.annotations || [];
      this.ownAnnotations = all.filter((a) => a._editable).map(stripMeta);
      this.otherAnnotations = all.filter((a) => !a._editable);
      this.ownUpdatedAt = data.own_updated_at || "";
      this.redrawAllOverlays();
    } catch {}
  }
  toggleAnnoMode() {
    this.annoMode = !this.annoMode;
    this.annoBtn.classList.toggle("active", this.annoMode);
    this.annoToolbar.style.display = this.annoMode ? "" : "none";
    for (const p of this.pages) if (p.overlay) p.overlay.style.pointerEvents = this.annoMode ? "auto" : "none";
    this.scrollEl.classList.toggle("anno-active", this.annoMode);
  }
  setTool(key) {
    this.tool = key;
    for (const k in this.toolBtns) this.toolBtns[k].classList.toggle("active", k === key);
  }
  setColor(c) {
    this.color = c;
    this.colorBtns.forEach((b, i) => b.classList.toggle("active", COLORS[i] === c));
  }

  attachOverlayHandlers(svg, pageNo) {
    svg.addEventListener("pointerdown", (e) => this.onOverlayDown(e, pageNo));
  }
  pdfPointFromEvent(e, pageNo) {
    const p = this.pages[pageNo - 1];
    const rect = p.overlay.getBoundingClientRect();
    const vx = e.clientX - rect.left, vy = e.clientY - rect.top;
    const [px, py] = p.viewport.convertToPdfPoint(vx, vy);  // D8: PDF座標へ
    return [px, py];
  }
  onOverlayDown(e, pageNo) {
    if (!this.annoMode) return;
    const p = this.pages[pageNo - 1];
    if (!p.viewport) return;
    if (this.tool === "hand") { this.selectAt(e, pageNo); return; }
    e.preventDefault();
    const start = this.pdfPointFromEvent(e, pageNo);

    if (this.tool === "note" || this.tool === "text" || this.tool === "comment") {
      let text = "";
      if (this.tool !== "note") { text = prompt(this.tool === "text" ? "テキスト" : "コメント") || ""; if (!text && this.tool === "comment") return; }
      else { text = prompt("付箋メモ（任意）") || ""; }
      this.addAnnotation({ id: uuid(), type: this.tool, page: pageNo, point: start, text, color: this.color });
      return;
    }
    // ドラッグ系（rect/line/pen）
    const points = [start];
    const move = (ev) => { points.push(this.pdfPointFromEvent(ev, pageNo)); this.drawDraft(pageNo, points); };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      this.clearDraft(pageNo);
      if (this.tool === "rect") {
        const [x0, y0] = points[0], [x1, y1] = points[points.length - 1];
        if (Math.abs(x1 - x0) < 2 && Math.abs(y1 - y0) < 2) return;
        this.addAnnotation({ id: uuid(), type: "rect", page: pageNo, rect: [x0, y0, x1, y1], color: this.color });
      } else if (this.tool === "line") {
        this.addAnnotation({ id: uuid(), type: "line", page: pageNo, points: [points[0], points[points.length - 1]], color: this.color });
      } else if (this.tool === "pen") {
        if (points.length >= 2) this.addAnnotation({ id: uuid(), type: "pen", page: pageNo, points, color: this.color });
      }
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  }

  addAnnotation(a) {
    this.pushUndo();
    this.ownAnnotations.push(a);
    this.redoStack = [];
    this.drawOverlay(a.page);
    this.scheduleSave();
  }
  pushUndo() { this.undoStack.push(JSON.stringify(this.ownAnnotations)); if (this.undoStack.length > 50) this.undoStack.shift(); }
  undo() {
    if (!this.undoStack.length) return;
    this.redoStack.push(JSON.stringify(this.ownAnnotations));
    this.ownAnnotations = JSON.parse(this.undoStack.pop());
    this.selectedId = null; this.redrawAllOverlays(); this.scheduleSave();
  }
  redo() {
    if (!this.redoStack.length) return;
    this.undoStack.push(JSON.stringify(this.ownAnnotations));
    this.ownAnnotations = JSON.parse(this.redoStack.pop());
    this.selectedId = null; this.redrawAllOverlays(); this.scheduleSave();
  }
  deleteSelected() {
    if (!this.selectedId) return;
    const a = this.ownAnnotations.find((x) => x.id === this.selectedId);
    this.pushUndo();
    this.ownAnnotations = this.ownAnnotations.filter((x) => x.id !== this.selectedId);
    this.selectedId = null;
    if (a) this.drawOverlay(a.page);
    this.scheduleSave();
  }
  onKey(e) {
    if ((e.key === "Delete" || e.key === "Backspace") && this.selectedId && this.annoMode) { e.preventDefault(); this.deleteSelected(); }
    else if ((e.metaKey || e.ctrlKey) && e.key === "z") { e.preventDefault(); e.shiftKey ? this.redo() : this.undo(); }
  }

  redrawAllOverlays() {
    for (let i = 1; i <= this.numPages; i++) {
      const p = this.pages[i - 1];
      if (p && p.overlay && p.viewport) this.drawOverlay(i);
    }
  }

  drawOverlay(pageNo) {
    const p = this.pages[pageNo - 1];
    if (!p || !p.overlay || !p.viewport) return;
    const svg = p.overlay;
    svg.innerHTML = "";
    const vp = p.viewport;
    const all = [...this.otherAnnotations.filter((a) => a.page === pageNo).map((a) => ({ a, own: false })),
                 ...this.ownAnnotations.filter((a) => a.page === pageNo).map((a) => ({ a, own: true }))];
    for (const { a, own } of all) this.drawOne(svg, vp, a, own);
  }
  drawOne(svg, vp, a, own) {
    const NS = "http://www.w3.org/2000/svg";
    const col = a.color || (own ? "#2563eb" : "#9333ea");
    const sel = own && a.id === this.selectedId;
    const v = (x, y) => vp.convertToViewportPoint(x, y);  // D8: 表示座標へ復元
    let el;
    if (a.type === "rect") {
      const [x0, y0] = v(a.rect[0], a.rect[1]), [x1, y1] = v(a.rect[2], a.rect[3]);
      el = document.createElementNS(NS, "rect");
      el.setAttribute("x", Math.min(x0, x1)); el.setAttribute("y", Math.min(y0, y1));
      el.setAttribute("width", Math.abs(x1 - x0)); el.setAttribute("height", Math.abs(y1 - y0));
      el.setAttribute("fill", "none"); el.setAttribute("stroke", col); el.setAttribute("stroke-width", sel ? 3 : 2);
    } else if (a.type === "line") {
      const [x0, y0] = v(a.points[0][0], a.points[0][1]), [x1, y1] = v(a.points[1][0], a.points[1][1]);
      el = document.createElementNS(NS, "line");
      el.setAttribute("x1", x0); el.setAttribute("y1", y0); el.setAttribute("x2", x1); el.setAttribute("y2", y1);
      el.setAttribute("stroke", col); el.setAttribute("stroke-width", sel ? 3 : 2);
    } else if (a.type === "pen") {
      el = document.createElementNS(NS, "polyline");
      el.setAttribute("points", a.points.map((pt) => v(pt[0], pt[1]).join(",")).join(" "));
      el.setAttribute("fill", "none"); el.setAttribute("stroke", col); el.setAttribute("stroke-width", sel ? 3 : 2);
    } else if (a.type === "note" || a.type === "text" || a.type === "comment") {
      const [x, y] = v(a.point[0], a.point[1]);
      const g = document.createElementNS(NS, "g");
      if (a.type === "note") {
        const star = document.createElementNS(NS, "text");
        star.setAttribute("x", x); star.setAttribute("y", y); star.setAttribute("fill", col);
        star.setAttribute("font-size", "18"); star.textContent = "★";
        g.appendChild(star);
      } else if (a.type === "comment") {
        const c = document.createElementNS(NS, "circle");
        c.setAttribute("cx", x); c.setAttribute("cy", y); c.setAttribute("r", 7);
        c.setAttribute("fill", col); g.appendChild(c);
      }
      if (a.text) {
        const t = document.createElementNS(NS, "text");
        t.setAttribute("x", x + (a.type === "comment" ? 12 : 16)); t.setAttribute("y", y);
        t.setAttribute("fill", col); t.setAttribute("font-size", "13");
        t.setAttribute("class", "anno-text-label"); t.textContent = a.text;
        g.appendChild(t);
      }
      el = g;
      if (sel) el.classList.add("selected");
    }
    if (!el) return;
    el.dataset.id = a.id;
    el.classList.add("anno-shape");
    if (own) el.classList.add("own"); else el.style.pointerEvents = "none"; // 他ユーザーは選択不可（D9）
    if (sel) el.classList.add("selected");
    svg.appendChild(el);
  }

  selectAt(e, pageNo) {
    const id = e.target?.dataset?.id || e.target?.parentNode?.dataset?.id;
    const a = id && this.ownAnnotations.find((x) => x.id === id);
    this.selectedId = a ? id : null;
    this.drawOverlay(pageNo);
  }

  drawDraft(pageNo, points) {
    const p = this.pages[pageNo - 1];
    if (!p.overlay) return;
    this.clearDraft(pageNo);
    const NS = "http://www.w3.org/2000/svg";
    const vp = p.viewport;
    const v = (x, y) => vp.convertToViewportPoint(x, y);
    let el;
    if (this.tool === "rect") {
      const [x0, y0] = v(points[0][0], points[0][1]), [x1, y1] = v(points[points.length - 1][0], points[points.length - 1][1]);
      el = document.createElementNS(NS, "rect");
      el.setAttribute("x", Math.min(x0, x1)); el.setAttribute("y", Math.min(y0, y1));
      el.setAttribute("width", Math.abs(x1 - x0)); el.setAttribute("height", Math.abs(y1 - y0));
      el.setAttribute("fill", "none");
    } else if (this.tool === "line") {
      const [x0, y0] = v(points[0][0], points[0][1]), [x1, y1] = v(points[points.length - 1][0], points[points.length - 1][1]);
      el = document.createElementNS(NS, "line");
      el.setAttribute("x1", x0); el.setAttribute("y1", y0); el.setAttribute("x2", x1); el.setAttribute("y2", y1);
    } else {
      el = document.createElementNS(NS, "polyline");
      el.setAttribute("points", points.map((pt) => v(pt[0], pt[1]).join(",")).join(" "));
      el.setAttribute("fill", "none");
    }
    el.setAttribute("stroke", this.color); el.setAttribute("stroke-width", 2);
    el.setAttribute("class", "anno-draft"); el.style.pointerEvents = "none";
    p.overlay.appendChild(el);
  }
  clearDraft(pageNo) {
    const p = this.pages[pageNo - 1];
    p.overlay?.querySelectorAll(".anno-draft").forEach((n) => n.remove());
  }

  scheduleSave() {
    this.annoStatus.textContent = "保存待ち…";
    clearTimeout(this.saveTimer);
    this.saveTimer = setTimeout(() => this.saveAnnotations(), SAVE_DEBOUNCE_MS);
  }
  async saveAnnotations() {
    this.annoStatus.textContent = "保存中…";
    try {
      const res = await fetch(`/api/cases/${encodeURIComponent(this.ctx.caseId)}/annotations/${this.doc.sha256}`, {
        method: "PUT", headers: { ...CSRF, "Content-Type": "application/json" },
        body: JSON.stringify({ base_updated_at: this.ownUpdatedAt, annotations: this.ownAnnotations.map(stripMeta) }),
      });
      if (res.status === 409) {
        this.annoStatus.textContent = "競合: 再読込しました";
        await this.loadAnnotations();
        return;
      }
      const data = await res.json();
      this.ownUpdatedAt = data.updated_at;
      this.annoStatus.textContent = `保存済（${data.count}件）`;
      this.ctx.onAnnotationsSaved?.(this.doc.sha256, data.count > 0);
    } catch (e) {
      this.annoStatus.textContent = "保存失敗";
    }
  }

  // ---- メタ（丁数・回転・カテゴリ等） ----
  async saveMeta(fields) {
    try {
      await fetch(`/api/cases/${encodeURIComponent(this.ctx.caseId)}/meta/${this.doc.sha256}`, {
        method: "PUT", headers: { ...CSRF, "Content-Type": "application/json" },
        body: JSON.stringify(fields),
      });
      this.ctx.onMetaSaved?.(this.doc.sha256, fields);
    } catch {}
  }

  // ---- 書き出し（墨消し警告 D13） ----
  exportPdf() {
    const ok = confirm(
      "注釈込みPDFを書き出します。\n\n" +
      "【重要】矩形注釈はマスキング（黒塗り）ではありません。第三者提出用の墨消しは別途" +
      "墨消しツールで行ってください。\n\n書き出しますか？"
    );
    if (!ok) return;
    const a = document.createElement("a");
    a.href = `/api/cases/${encodeURIComponent(this.ctx.caseId)}/export/${this.doc.sha256}`;
    document.body.appendChild(a); a.click(); a.remove();
  }
  download() {
    const a = document.createElement("a");
    a.href = `/api/cases/${encodeURIComponent(this.ctx.caseId)}/pdf/${this.doc.sha256}`;
    a.download = this.doc.file_name || "document.pdf";
    document.body.appendChild(a); a.click(); a.remove();
  }
  navDoc(dir) {
    const docs = this.ctx.documents || [];
    const idx = docs.findIndex((d) => d.sha256 === this.doc.sha256);
    if (idx < 0) return;
    const next = docs[idx + dir];
    if (next) { this.destroy(); this.ctx.onNavigate(next.sha256); }
  }
  destroy() {
    this.destroyed = true;
    clearTimeout(this.saveTimer);
    clearTimeout(this._fitDebounceTimer);
    if (this._resizeObserver) { try { this._resizeObserver.disconnect(); } catch {} this._resizeObserver = null; }
    for (const p of this.pages) if (p.renderTask) { try { p.renderTask.cancel(); } catch {} }
    try { this.loadingTask?.destroy(); } catch {}
  }
}

// ---- ユーティリティ ----
function stripMeta(a) { const o = {}; for (const k in a) if (!k.startsWith("_")) o[k] = a[k]; return o; }
function sideBtn(label, onClick) { const b = document.createElement("button"); b.className = "side-btn"; b.textContent = label; b.onclick = onClick; return b; }
function tbBtn(label, onClick) { const b = document.createElement("button"); b.className = "tb-btn"; b.textContent = label; b.onclick = onClick; return b; }
function label(t) { const s = document.createElement("span"); s.className = "tb-label"; s.textContent = t; return s; }
async function openExternal(caseId, sha) {
  try { await fetch(`/api/cases/${encodeURIComponent(caseId)}/open-file/${sha}`, { method: "POST", headers: CSRF }); }
  catch (e) { alert("ファイルを開けませんでした: " + e); }
}
function escapeHtml(s) { return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
