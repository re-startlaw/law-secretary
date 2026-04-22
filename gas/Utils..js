/**
 * 設定シート（画像の内容）から値を読み込む
 */
function getSettingsFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_SETTINGS);
  
  if (!sheet) {
    throw new Error(`「${CONFIG.SHEET_SETTINGS}」シートが見つかりません。`);
  }

  // 画像の配置に合わせてセルを直接指定
  return {
    model: sheet.getRange('B2').getValue(),    // AIモデル
    example: sheet.getRange('B3').getValue(),  // メール例文
    context: sheet.getRange('B4').getValue(),  // コンテキスト情報
    criteria: sheet.getRange('B5').getValue()  // 返信判断基準
  };
}

/**
 * ログシートへ書き込み（画像のA~E列に対応）
 */
function logToSheet(sender, subject, action, url) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_LOGS);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_LOGS);
    sheet.appendRow(['日時', '送信者', '件名', 'アクション', 'メールURL']);
  }
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy年M月d日HH時mm分');
  sheet.appendRow([timestamp, sender, subject, action, url]);
}

/**
 * APIキー設定
 */
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Gemini APIキー設定', 'APIキーを入力してください:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty(CONFIG.API_PROP_KEY, response.getResponseText().trim());
    ui.alert('保存しました。');
  }
}

/**
 * ラベル取得・作成
 */
function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  return label;
}

/**
 * "Name <a@b.com>" 形式または "a@b.com" から純粋なメールアドレスを抽出
 */
function extractEmail_(fromField) {
  if (!fromField) return '';
  const m = String(fromField).match(/<([^>]+)>/);
  return (m ? m[1] : String(fromField)).trim();
}

/**
 * 引用部（> から始まる行・"On ... wrote:"・日本語の日時引用ヘッダ等）を除去
 */
function stripQuotedReplies_(body) {
  if (!body) return '';
  const markers = [
    /\n[-]{2,}\s*Original Message\s*[-]{2,}/i,
    /\n\d{4}年\d{1,2}月\d{1,2}日.{0,40}(?:のメール|\s(?:に|午前|午後)).{0,40}:\s*\n/,
    /\nOn .{1,80} wrote:\s*\n/,
    /\n>.*\n/,
    /\n\-{2,}\s*\n[\s\S]*$/ // 署名区切り以降を落とす
  ];
  let cut = body.length;
  for (const r of markers) {
    const m = body.match(r);
    if (m && typeof m.index === 'number' && m.index < cut) cut = m.index;
  }
  return body.substring(0, cut).replace(/\s+$/, '').trim();
}

/**
 * 自分が本人として送信した本文を抽出するヘルパー
 */
function collectMyBodies_(query, max) {
  const myEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const bodies = [];
  const threads = GmailApp.search(query, 0, max * 3);
  for (const t of threads) {
    if (bodies.length >= max) break;
    const msgs = t.getMessages();
    for (const m of msgs) {
      if (bodies.length >= max) break;
      const from = extractEmail_(m.getFrom()).toLowerCase();
      if (!from || from !== myEmail) continue;
      const body = stripQuotedReplies_(m.getPlainBody() || '');
      if (body.length < 30) continue; // 短すぎるリプライは文体学習に不向き
      bodies.push(body.substring(0, 1500));
    }
  }
  return bodies;
}

/**
 * 特定の相手への過去送信メールを最大 maxExamples 通取得
 */
function fetchSentExamplesForRecipient(recipientEmail, maxExamples) {
  if (!recipientEmail) return [];
  const safe = recipientEmail.replace(/[^\w@.\-+]/g, '');
  if (!safe) return [];
  return collectMyBodies_(`in:sent to:${safe} newer_than:2y`, maxExamples);
}

/**
 * 直近の送信メールを最大 maxExamples 通取得（相手履歴がない場合のフォールバック）
 */
function fetchRecentSentExamples(maxExamples) {
  return collectMyBodies_('in:sent newer_than:60d', maxExamples);
}