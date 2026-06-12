// インラインPDFビューア。
// フェーズ1-B: 行展開の器（とじる・前後移動）。
// フェーズ1-C: PDF.js 仮想化レンダリング・textLayer・FILE/TXT切替・ズーム・回転・
//             文書内検索・DL を実装する。

export function mountInlineViewer(doc, container, ctx) {
  container.innerHTML = "";
  const box = document.createElement("div");
  box.className = "inline-viewer";

  const bar = document.createElement("div");
  bar.className = "viewer-bar";
  bar.innerHTML = `<strong>${doc.evidence_no || "—"}</strong> ${doc.title}`;

  const close = document.createElement("button");
  close.textContent = "とじる";
  close.className = "viewer-close";
  close.onclick = () => ctx.onClose();
  bar.appendChild(close);

  box.appendChild(bar);
  const body = document.createElement("div");
  body.className = "viewer-body";
  body.textContent = "ビューア本体はフェーズ1-Cで実装します。";
  box.appendChild(body);

  container.appendChild(box);
}
