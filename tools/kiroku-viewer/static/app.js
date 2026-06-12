// 記録ビューア フロントエンド エントリ。
// フェーズ0: 疎通確認のみ（cases/documents を取得して件数を表示）。
// 文書DB表・インラインビューア・検索UIはフェーズ1で実装する。

async function boot() {
  const app = document.getElementById("app");
  try {
    const res = await fetch("/api/cases");
    const data = await res.json();
    const cases = data.cases || [];
    if (!cases.length) {
      app.innerHTML = '<p class="placeholder">cases.json に事件が登録されていません。</p>';
      return;
    }
    const first = cases[0];
    const docRes = await fetch(`/api/cases/${encodeURIComponent(first.id)}/documents`);
    const docData = await docRes.json();
    const n = (docData.documents || []).length;
    app.innerHTML =
      `<p class="placeholder">疎通OK: 事件「${first.name}」に ${n} 文書（フェーズ1でUI実装）。</p>`;
  } catch (e) {
    app.innerHTML = `<p class="placeholder">読み込みエラー: ${e}</p>`;
  }
}

boot();
