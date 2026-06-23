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
function logToSheet(sender, subject, action, url, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_LOGS);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_LOGS);
    sheet.appendRow(['日時', '送信者', '件名', 'アクション', 'メールURL', '詳細']);
  } else if (sheet.getLastColumn() < 6) {
    sheet.getRange(1, 6).setValue('詳細');
  }
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy年M月d日HH時mm分');
  sheet.appendRow([timestamp, sender, subject, action, url, detail || '']);
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

function hasYonetaniSalutation_(body) {
  const firstLines = String(body || '')
    .split(/\r?\n/)
    .slice(0, 3)
    .join('\n');
  return /米谷|先生/.test(firstLines);
}

function extractEmailsFromText_(text) {
  const matches = String(text || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi);
  if (!matches) return [];
  const seen = {};
  return matches
    .map((email) => email.trim())
    .filter((email) => {
      const key = email.toLowerCase();
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });
}

/**
 * 全員返信相当の Cc を組み立てる。
 *
 * createDraftReplyAll は元To＋元Ccを自動で全員追加し、cc オプションは「追加」しかできない
 * （アカウント主アドレスしか自動除外されない）ため、エイリアス受信や別アカウント実行だと
 * 自分が宛先に残ってしまう。そこで Main 側では createDraftReply（宛先=送信者のみ／自動追加なし）
 * を使い、送信者以外の宛先をこの関数が返す Cc として明示的に追加する。
 *
 * 元To＋元Cc から実体アドレスを抽出し、自分側アドレス（主アドレス＋ CONFIG.OWNER_EMAILS）と
 * 元送信者（createDraftReply の to に入るので重複排除）を除いた一覧を返す。
 * 空文字を渡すと Cc なし＝送信者のみへの返信になる。
 */
function buildReplyAllCc_(sender, originalTo, originalCc, myEmail) {
  const exclude = {};
  CONFIG.OWNER_EMAILS
    .concat([myEmail, extractEmail_(sender)])
    .forEach((e) => {
      const lower = String(e || '').toLowerCase();
      if (lower) exclude[lower] = true;
    });

  return extractEmailsFromText_(`${originalTo || ''}, ${originalCc || ''}`)
    .filter((email) => !exclude[email.toLowerCase()])
    .join(', ');
}

/**
 * 返信モードを決定する。
 * - 通常メール: { mode: 'replyAll' } … 送信者をToに、元To＋元Ccの残り全員をCcにして返信（自分は buildReplyAllCc_ で除外）
 * - ココナラ等の no-reply 相談メール: { mode: 'direct', to } … 本文から拾った相談者アドレスへ直接送る
 *
 * 表示名内カンマ・重複排除・自分の除外は buildReplyAllCc_ 側で扱うため、
 * ここでは特例の宛先抽出だけを担う。
 */
function resolveReplyTarget_(sender, subject, body, myEmail) {
  const senderEmail = extractEmail_(sender).toLowerCase();

  if (
    senderEmail === 'no-reply@mail.coconala.com' &&
    String(subject || '').indexOf('[ココナラ法律相談] お問い合わせメールが届いています') >= 0
  ) {
    const candidates = extractEmailsFromText_(body).filter((email) => {
      const lower = email.toLowerCase();
      return lower !== senderEmail && lower !== String(myEmail || '').toLowerCase();
    });

    if (candidates.length > 0) {
      return { mode: 'direct', to: candidates[0] };
    }
  }

  return { mode: 'replyAll' };
}

/**
 * B5の明確な条件はAIへ渡す前に機械判定する
 */
function evaluateStaticReplyNecessity_(sender, subject, body, originalTo, originalCc, myEmail) {
  const senderEmail = extractEmail_(sender).toLowerCase();
  const subjectText = String(subject || '');
  const bodyText = String(body || '');

  const isCoconalaInquiry =
    senderEmail === 'no-reply@mail.coconala.com' &&
    subjectText.indexOf('[ココナラ法律相談] お問い合わせメールが届いています') >= 0;

  if (isCoconalaInquiry) {
    return {
      needsReply: true,
      reason: 'ココナラ法律相談の例外メールのため、返信下書き作成対象にしました。'
    };
  }

  const senderBlacklist = new Set([
    'no-reply@printing.ne.jp',
    't-hoso-ml@googlegroups.com',
    'shinwazenki@googlegroups.com',
    'noreply@tm.openai.com',
    'noreply-apps-scripts-notifications@google.com',
    'noreply@appsheet.com',
    'n.kometani@re-startlaw.com',
    'nobukosato2020@gmail.com',
    'support@myteam108.jp'
  ]);

  if (senderBlacklist.has(senderEmail)) {
    return {
      needsReply: false,
      reason: `送信元 ${senderEmail} は返信不要ブラックリストに含まれるため、返信不要としました。`
    };
  }

  const subjectBlacklist = ['[t-hoso-ml:', '[shinwazenki'];
  if (subjectBlacklist.some((pattern) => subjectText.indexOf(pattern) >= 0)) {
    return {
      needsReply: false,
      reason: '件名が返信不要ブラックリストに一致するため、返信不要としました。'
    };
  }

  if (!hasYonetaniSalutation_(bodyText)) {
    return {
      needsReply: false,
      reason: '本文冒頭に「米谷」または「先生」の宛名が見当たらないため、返信不要としました。'
    };
  }

  return null;
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
      bodies.push(body.substring(0, 800));
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

/**
 * GeminiがMarkdown等を混ぜても、最初のJSONオブジェクトだけを取り出す
 */
function parseJsonFromGeminiResponse_(text) {
  if (!text) return null;
  const raw = String(text).trim()
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/```\s*$/i, '')
    .trim();

  try {
    return JSON.parse(raw);
  } catch (e) {
    const start = raw.indexOf('{');
    const end = raw.lastIndexOf('}');
    if (start < 0 || end <= start) return null;
    try {
      return JSON.parse(raw.substring(start, end + 1));
    } catch (e2) {
      return null;
    }
  }
}

/**
 * 裁判所・検察庁・警察・相手方代理人等は保留寄りの文体に寄せる
 */
function classifySenderRisk_(sender, subject, body) {
  const text = `${sender || ''}\n${subject || ''}\n${(body || '').substring(0, 500)}`.toLowerCase();
  const highRiskPatterns = [
    '裁判所', '検察', '検察庁', '警察', '弁護士', '法律事務所', '法務局',
    '家庭裁判所', '地方裁判所', '簡易裁判所', '高等裁判所',
    'courts.go.jp', 'moj.go.jp', 'npa.go.jp', 'police', 'law'
  ];
  const isHighRisk = highRiskPatterns.some((p) => text.indexOf(String(p).toLowerCase()) >= 0);

  if (isHighRisk) {
    return {
      level: 'high',
      label: '高リスク相手',
      instruction: [
        'この相手は高リスク相手として扱う。',
        '文体は短く、硬く、保留寄りにする。',
        '原則として「受領しました」「確認します」「追って回答します」「必要資料をご送付ください」程度に留める。',
        '謝罪、譲歩、事実認定、法律判断、事件方針、見通しに見える表現は避ける。'
      ].join('\n')
    };
  }

  return {
    level: 'normal',
    label: '通常相手',
    instruction: [
      'この相手は通常相手として扱う。',
      '丁寧だが簡潔にし、必要な返答・依頼・保留事項だけを書く。',
      '法律判断、事件方針、見通しは通常相手でも必ず【要回答】にする。'
    ].join('\n')
  };
}

function normalizeStringArray_(value, maxItems, maxLength) {
  if (!Array.isArray(value)) return [];
  return value
    .map((v) => String(v || '').replace(/\s+/g, ' ').trim())
    .filter((v) => v)
    .slice(0, maxItems)
    .map((v) => v.substring(0, maxLength));
}

/**
 * & < > " をHTMLエスケープする
 */
function escapeHtml_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * 元メールHTMLを blockquote 埋め込み用に軽量サニタイズする
 * - <style>タグ除去（CSS漏洩防止）
 * - <body>があれば内側だけ抽出（Outlook等のフルHTMLドキュメント対策）
 * - cid:インライン画像を [画像] に置換（参照切れ防止）
 */
function sanitizeQuotedHtml_(html) {
  let s = String(html || '');
  s = s.replace(/<style[\s\S]*?<\/style>/gi, '');
  const bodyMatch = s.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  if (bodyMatch) s = bodyMatch[1];
  s = s.replace(/<img[^>]*src=["']?cid:[^"'>]*["']?[^>]*>/gi, '[画像]');
  return s;
}

/**
 * Gmail の「全員に返信」ボタンと同等の引用付きHTML本文を組み立てる
 * 引用HTMLが100KBを超える場合は省略メッセージに差し替える
 */
function buildReplyHtmlBody_(plainDraftBody, lastMessage) {
  const draftHtml = escapeHtml_(plainDraftBody).replace(/\n/g, '<br>');
  const date = lastMessage.getDate();
  const tz = Session.getScriptTimeZone();
  // ISO曜日(u): 1=月…7=日 → 日本語曜日配列（インデックス0=月）
  const youbi = ['月', '火', '水', '木', '金', '土', '日'][Number(Utilities.formatDate(date, tz, 'u')) - 1];
  const dateStr = Utilities.formatDate(date, tz, 'yyyy年M月d日') +
    '(' + youbi + ') ' + Utilities.formatDate(date, tz, 'H:mm');
  const quoteHeader = dateStr + ' ' + escapeHtml_(lastMessage.getFrom()) + ':';
  let quotedHtml = sanitizeQuotedHtml_(lastMessage.getBody());
  if (quotedHtml.length > 100000) {
    quotedHtml = '（元メッセージが長いため引用を省略しました）';
  }
  return (
    '<div dir="ltr">' + draftHtml + '</div><br>' +
    '<div class="gmail_quote">' +
    '<div dir="ltr" class="gmail_attr">' + quoteHeader + '<br></div>' +
    '<blockquote class="gmail_quote" style="margin:0px 0px 0px 0.8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex">' +
    quotedHtml +
    '</blockquote></div>'
  );
}

/**
 * AI出力を固定フォーマットでGmail下書き本文にする
 */
function formatDraftBody_(parsed, rawResponse, riskProfile) {
  if (!parsed || typeof parsed !== 'object') {
    // 解析失敗時は送信せず確認用。誤送信防止のため明示する
    return [
      '※AI出力をJSONとして解析できませんでした。送信せず内容を確認してください。',
      '',
      '【AI出力】',
      String(rawResponse || '').substring(0, 2000)
    ].join('\n');
  }

  const placeholders = normalizeStringArray_(parsed.placeholders, 10, 120);
  const riskFlags = normalizeStringArray_(parsed.riskFlags, 10, 160);
  const draft = String(parsed.draft || '').trim()
    || 'ご連絡ありがとうございます。\n内容を確認のうえ、改めてご連絡いたします。';

  // そのまま送れる本文＋署名をクリーンに出力する。
  const parts = [draft, '', '米谷尚起'];

  // AIが未確定箇所・リスクを検出したときだけ、本文・署名の下に注記として残す。
  if (placeholders.length > 0) {
    parts.push('', '【要入力】');
    placeholders.forEach((item) => parts.push(`- ${item}`));
  }

  if (riskFlags.length > 0) {
    parts.push('', '【要確認】');
    riskFlags.forEach((item) => parts.push(`- ${item}`));
  }

  return parts.join('\n');
}
