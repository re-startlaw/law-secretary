// インラインPDFビューア（フェーズ1-C）。
// PDF.js 仮想化レンダリング・textLayer（選択コピー）・FILE/TXT切替・ズーム・
// ページ回転90°・ページ送り/ジャンプ・文書内検索（D10）・DL・前後文書↑↓・とじる。
// 注釈オーバーレイ・タブ表示はフェーズ2。

import * as pdfjsLib from "/static/vendor/pdfjs/build/pdf.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = "/static/vendor/pdfjs/build/pdf.worker.mjs";
const PDFJS_BASE = "/static/vendor/pdfjs/";
const RENDER_BUFFER = 2;       // 表示中±2ページのみ canvas 保持（D7）
const MAX_CONCURRENT = 2;      // 同時レンダリング2枚まで（D7）

function normalize(s) {
  return (s || "").normalize("NFKC").toLowerCase();
}

export function mountInlineViewer(doc, container, ctx) {
  container.innerHTML = "";
  const root = document.createElement("div");
  root.className = "inline-viewer";
  container.appendChild(root);

  // メディアファイルは閲覧不可。open-file で外部アプリ起動。
  if (doc.kind === "media") {
    root.appendChild(buildBar(doc, ctx));
    const body = document.createElement("div");
    body.className = "viewer-body media-body";
    body.innerHTML = `<p>メディアファイル（${doc.file_name}）。ブラウザでは再生しません。</p>`;
    const open = document.createElement("button");
    open.textContent = "QuickTime等で開く";
    open.onclick = () => openExternal(ctx.caseId, doc.sha256);
    body.appendChild(open);
    root.appendChild(body);
    return;
  }

  new ViewerInstance(doc, root, ctx);
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
    this.rotation = 0;
    this.currentPage = 1;
    this.numPages = 0;
    this.pages = [];            // {div, canvas, textDiv, rendered, rendering, renderTask, baseW, baseH}
    this.inFlight = 0;
    this.renderQueue = [];
    this.mode = "file";         // file | txt
    this.textPages = null;      // [{page_no, text}]（索引済のみ）
    this.searchHits = [];       // ヒットページ番号
    this.searchIdx = -1;
    this.destroyed = false;
    this.build();
    this.load();
  }

  build() {
    this.root.appendChild(buildBar(this.doc, this.ctx));

    const layout = document.createElement("div");
    layout.className = "viewer-layout";

    // 左サイドバー
    const side = document.createElement("div");
    side.className = "viewer-side";
    this.fileBtn = sideBtn("FILE", () => this.setMode("file"));
    this.txtBtn = sideBtn("TXT", () => this.setMode("txt"));
    this.fileBtn.classList.add("active");
    side.append(this.fileBtn, this.txtBtn);
    side.appendChild(sideBtn("DL", () => this.download()));
    side.appendChild(sideBtn("↑ 前の文書", () => this.navDoc(-1)));
    side.appendChild(sideBtn("↓ 次の文書", () => this.navDoc(1)));
    side.appendChild(sideBtn("とじる", () => this.ctx.onClose()));
    layout.appendChild(side);

    // 右本体
    const right = document.createElement("div");
    right.className = "viewer-right";
    right.appendChild(this.buildToolbar());

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
  }

  buildToolbar() {
    const tb = document.createElement("div");
    tb.className = "viewer-toolbar";

    // ページジャンプ
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

    // ズーム
    this.zoomLabel = document.createElement("span");
    this.zoomLabel.className = "zoom-label";
    const zoomGrp = document.createElement("span");
    zoomGrp.className = "tb-grp";
    zoomGrp.append(tbBtn("−", () => this.zoom(-0.2)), this.zoomLabel,
      tbBtn("＋", () => this.zoom(0.2)), tbBtn("100%", () => this.setScale(1.0)),
      tbBtn("幅", () => this.fitWidth()));
    tb.appendChild(zoomGrp);

    // 回転
    tb.appendChild(tbBtn("⟳ 90°", () => this.rotate()));

    // 文書内検索（D10）
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
      tbBtn("‹", () => this.cycleHit(-1)), tbBtn("›", () => this.cycleHit(1)),
      this.searchCount);
    tb.appendChild(searchGrp);

    return tb;
  }

  async load() {
    const url = `/api/cases/${encodeURIComponent(this.ctx.caseId)}/pdf/${this.doc.sha256}`;
    try {
      this.loadingTask = pdfjsLib.getDocument({
        url,
        cMapUrl: PDFJS_BASE + "cmaps/",
        cMapPacked: true,
        standardFontDataUrl: PDFJS_BASE + "standard_fonts/",
      });
      this.pdf = await this.loadingTask.promise;
    } catch (e) {
      this.scrollEl.innerHTML = `<p class="placeholder">PDFを開けません: ${escapeHtml(String(e))}</p>`;
      return;
    }
    if (this.destroyed) return;
    this.numPages = this.pdf.numPages;
    this.pageTotal.textContent = `/ ${this.numPages}`;

    // 1ページ目の寸法を基準として全プレースホルダ高さを推定（D7）。
    const first = await this.pdf.getPage(1);
    const base = first.getViewport({ scale: 1, rotation: 0 });
    this.defaultBaseW = base.width;
    this.defaultBaseH = base.height;

    this.fitWidth(true);  // 既定は幅フィット
    this.buildPlaceholders();
    this.updateVisible();
    // 検索結果から開いた場合は該当ページへジャンプ。
    if (this.ctx.initialPage && this.ctx.initialPage > 1) this.jumpTo(this.ctx.initialPage);
    // テキスト（TXT表示・文書内検索用）を遅延取得。
    if (this.doc.indexed) this.fetchText();
  }

  buildPlaceholders() {
    this.scrollEl.innerHTML = "";
    this.pages = [];
    for (let i = 1; i <= this.numPages; i++) {
      const div = document.createElement("div");
      div.className = "pdf-page";
      div.dataset.page = i;
      const p = { div, canvas: null, textDiv: null, rendered: false, rendering: false,
        renderTask: null, baseW: this.defaultBaseW, baseH: this.defaultBaseH };
      this.setPlaceholderHeight(p);
      const label = document.createElement("div");
      label.className = "page-label";
      label.textContent = i;
      div.appendChild(label);
      this.scrollEl.appendChild(div);
      this.pages.push(p);
    }
  }

  displayDims(p) {
    const rot = this.rotation % 180 !== 0;
    const w = (rot ? p.baseH : p.baseW) * this.scale;
    const h = (rot ? p.baseW : p.baseH) * this.scale;
    return { w, h };
  }

  setPlaceholderHeight(p) {
    const { w, h } = this.displayDims(p);
    p.div.style.width = w + "px";
    p.div.style.height = h + "px";
  }

  onScroll() {
    if (this._scrollRaf) return;
    this._scrollRaf = requestAnimationFrame(() => {
      this._scrollRaf = null;
      this.updateVisible();
      this.updateCurrentPage();
    });
  }

  visibleRange() {
    const top = this.scrollEl.scrollTop;
    const bottom = top + this.scrollEl.clientHeight;
    let first = this.numPages, last = 1;
    for (let i = 0; i < this.pages.length; i++) {
      const div = this.pages[i].div;
      const dTop = div.offsetTop - this.scrollEl.offsetTop;
      const dBottom = dTop + div.offsetHeight;
      if (dBottom >= top && dTop <= bottom) {
        first = Math.min(first, i + 1);
        last = Math.max(last, i + 1);
      }
    }
    if (first > last) { first = this.currentPage; last = this.currentPage; }
    return { first, last };
  }

  updateVisible() {
    const { first, last } = this.visibleRange();
    const lo = Math.max(1, first - RENDER_BUFFER);
    const hi = Math.min(this.numPages, last + RENDER_BUFFER);
    for (let i = 1; i <= this.numPages; i++) {
      const p = this.pages[i - 1];
      if (i >= lo && i <= hi) {
        if (!p.rendered && !p.rendering) this.enqueue(i);
      } else if (p.rendered || p.rendering) {
        this.unload(i);
      }
    }
  }

  updateCurrentPage() {
    const top = this.scrollEl.scrollTop + this.scrollEl.clientHeight * 0.3;
    let cur = 1;
    for (let i = 0; i < this.pages.length; i++) {
      const div = this.pages[i].div;
      const dTop = div.offsetTop - this.scrollEl.offsetTop;
      if (dTop <= top) cur = i + 1; else break;
    }
    if (cur !== this.currentPage) {
      this.currentPage = cur;
      this.pageInput.value = String(cur);
      if (this.mode === "txt") this.renderTxt();
    }
  }

  enqueue(i) {
    const p = this.pages[i - 1];
    if (p.rendered || p.rendering || p.queued) return;
    p.queued = true;
    this.renderQueue.push(i);
    this.pump();
  }

  pump() {
    while (this.inFlight < MAX_CONCURRENT && this.renderQueue.length) {
      const i = this.renderQueue.shift();
      const p = this.pages[i - 1];
      p.queued = false;
      this.renderPage(i);
    }
  }

  async renderPage(i) {
    const p = this.pages[i - 1];
    if (p.rendered || p.rendering || this.destroyed) return;
    p.rendering = true;
    this.inFlight++;
    try {
      const page = await this.pdf.getPage(i);
      if (this.destroyed) return;
      const base = page.getViewport({ scale: 1, rotation: 0 });
      p.baseW = base.width; p.baseH = base.height;
      this.setPlaceholderHeight(p);
      const viewport = page.getViewport({ scale: this.scale, rotation: this.rotation });
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

      p.canvas = canvas; p.textDiv = textDiv;
      // ラベルを残しつつ canvas/textLayer を差し込む
      p.div.querySelector(".page-label")?.remove();
      p.div.append(canvas, textDiv);

      p.renderTask = page.render({
        canvasContext: cctx,
        viewport,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      });
      await p.renderTask.promise;
      if (this.destroyed) return;

      const tl = new pdfjsLib.TextLayer({
        textContentSource: page.streamTextContent(),
        container: textDiv,
        viewport,
      });
      await tl.render();
      p.textLayerObj = tl;
      p.rendered = true;
      this.applyHighlight(p, i);
    } catch (e) {
      if (e && e.name === "RenderingCancelledException") {
        // 範囲外に出てキャンセルされた。正常。
      } else {
        console.warn("page render failed", i, e);
      }
    } finally {
      p.rendering = false;
      p.renderTask = null;
      this.inFlight--;
      this.pump();
    }
  }

  unload(i) {
    const p = this.pages[i - 1];
    if (p.renderTask) { try { p.renderTask.cancel(); } catch {} }
    p.canvas?.remove(); p.textDiv?.remove();
    p.canvas = null; p.textDiv = null; p.rendered = false; p.rendering = false;
    if (!p.div.querySelector(".page-label")) {
      const label = document.createElement("div");
      label.className = "page-label";
      label.textContent = i;
      p.div.appendChild(label);
    }
  }

  reflow() {
    // ズーム/回転変更時: 全プレースホルダ高さを更新し、可視ページを再描画。
    for (let i = 1; i <= this.numPages; i++) {
      const p = this.pages[i - 1];
      this.unload(i);
      this.setPlaceholderHeight(p);
    }
    this.renderQueue = [];
    this.updateVisible();
  }

  // ---- ツールバー操作 ----
  setScale(s) {
    this.scale = Math.max(0.2, Math.min(5, s));
    this.zoomLabel.textContent = Math.round(this.scale * 100) + "%";
    this.reflow();
  }
  zoom(delta) { this.setScale(this.scale + delta); }
  fitWidth(silent) {
    const avail = this.scrollEl.clientWidth - 32;
    const rot = this.rotation % 180 !== 0;
    const baseW = rot ? this.defaultBaseH : this.defaultBaseW;
    const s = avail > 0 && baseW ? avail / baseW : 1.0;
    if (silent) { this.scale = Math.max(0.2, Math.min(5, s)); this.zoomLabel.textContent = Math.round(this.scale * 100) + "%"; }
    else this.setScale(s);
  }
  rotate() { this.rotation = (this.rotation + 90) % 360; this.reflow(); }

  jumpTo(n) {
    if (!n || isNaN(n)) return;
    n = Math.max(1, Math.min(this.numPages, n));
    const div = this.pages[n - 1].div;
    this.scrollEl.scrollTop = div.offsetTop - this.scrollEl.offsetTop;
    this.currentPage = n;
    this.pageInput.value = String(n);
    this.updateVisible();
    if (this.mode === "txt") this.renderTxt();
  }

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
    if (!this.doc.indexed) {
      this.txtEl.innerHTML = '<p class="placeholder">未索引のためテキストはありません。</p>';
      return;
    }
    if (!this.textPages) { this.txtEl.innerHTML = '<p class="placeholder">テキスト取得中…</p>'; return; }
    const pg = this.textPages.find((t) => t.page_no === this.currentPage);
    const head = `<div class="txt-head">${this.currentPage} / ${this.numPages} ページ</div>`;
    this.txtEl.innerHTML = head + `<pre class="txt-body">${escapeHtml(pg ? pg.text : "")}</pre>`;
  }

  async fetchText() {
    try {
      const res = await fetch(`/api/cases/${encodeURIComponent(this.ctx.caseId)}/text/${this.doc.sha256}`);
      const data = await res.json();
      this.textPages = data.pages || [];
      if (this.mode === "txt") this.renderTxt();
    } catch (e) { /* テキスト無しでも閲覧は継続 */ }
  }

  // ---- 文書内検索（D10：正規化部分一致でヒットページ→ジャンプ） ----
  async runDocSearch() {
    const q = normalize(this.searchInput.value.trim());
    this.clearHighlights();
    if (!q) { this.searchHits = []; this.searchIdx = -1; this.searchCount.textContent = ""; this.reHighlightVisible(); return; }
    if (!this.textPages) await this.fetchText();
    const hits = [];
    for (const t of (this.textPages || [])) {
      if (normalize(t.text).includes(q)) hits.push(t.page_no);
    }
    this.searchHits = hits;
    this.searchQuery = q;
    this.searchIdx = hits.length ? 0 : -1;
    this.searchCount.textContent = hits.length ? `1/${hits.length}` : "0件";
    this.reHighlightVisible();
    if (hits.length) this.jumpTo(hits[0]);
  }
  cycleHit(dir) {
    if (!this.searchHits.length) return;
    this.searchIdx = (this.searchIdx + dir + this.searchHits.length) % this.searchHits.length;
    this.searchCount.textContent = `${this.searchIdx + 1}/${this.searchHits.length}`;
    this.jumpTo(this.searchHits[this.searchIdx]);
  }
  clearHighlights() { this.searchQuery = null; this.reHighlightVisible(); }
  reHighlightVisible() {
    for (let i = 1; i <= this.numPages; i++) {
      const p = this.pages[i - 1];
      if (p.rendered) this.applyHighlight(p, i);
    }
  }
  applyHighlight(p, i) {
    if (!p.textDiv) return;
    p.div.classList.toggle("search-hit-page", !!this.searchQuery && this.searchHits.includes(i));
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
    for (const p of this.pages) { if (p.renderTask) { try { p.renderTask.cancel(); } catch {} } }
    try { this.loadingTask?.destroy(); } catch {}
  }
}

// ---- ユーティリティ ----
function sideBtn(label, onClick) {
  const b = document.createElement("button");
  b.className = "side-btn";
  b.textContent = label;
  b.onclick = onClick;
  return b;
}
function tbBtn(label, onClick) {
  const b = document.createElement("button");
  b.className = "tb-btn";
  b.textContent = label;
  b.onclick = onClick;
  return b;
}
async function openExternal(caseId, sha) {
  try {
    await fetch(`/api/cases/${encodeURIComponent(caseId)}/open-file/${sha}`, {
      method: "POST",
      headers: { "X-Kiroku-Viewer": "1" },
    });
  } catch (e) { alert("ファイルを開けませんでした: " + e); }
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}
