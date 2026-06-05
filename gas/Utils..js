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

function hasAddressInList_(list, email) {
  const lowerEmail = String(email || '').toLowerCase();
  if (!lowerEmail) return false;
  return String(list || '')
    .split(',')
    .map((e) => extractEmail_(e).toLowerCase())
    .some((e) => e === lowerEmail);
}

function hasYonetaniSalutation_(body) {
  const firstLines = String(body || '')
    .split(/\r?\n/)
    .slice(0, 3)
    .join('\n');
  return /米谷|先生/.test(firstLines);
}

function excludeMyAddress_(list, myEmail) {
  const lowerMyEmail = String(myEmail || '').toLowerCase();
  return String(list || '')
    .split(',')
    .map((e) => e.trim())
    .filter((e) => e && (!lowerMyEmail || !e.toLowerCase().includes(lowerMyEmail)))
    .join(', ');
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
 * no-reply系の相談メールは、実際の相談者メールアドレスを本文から拾う
 */
function resolveReplyRecipients_(sender, subject, body, originalTo, originalCc, myEmail) {
  const senderEmail = extractEmail_(sender).toLowerCase();
  const ccList = excludeMyAddress_(originalCc, myEmail);

  if (
    senderEmail === 'no-reply@mail.coconala.com' &&
    String(subject || '').indexOf('[ココナラ法律相談] お問い合わせメールが届いています') >= 0
  ) {
    const candidates = extractEmailsFromText_(body).filter((email) => {
      const lower = email.toLowerCase();
      return lower !== senderEmail && lower !== String(myEmail || '').toLowerCase();
    });

    if (candidates.length > 0) {
      return { to: candidates[0], cc: '' };
    }
  }

  const toList = [sender, excludeMyAddress_(originalTo, myEmail)].filter((e) => e).join(', ');
  return { to: toList, cc: ccList };
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

  if (!hasAddressInList_(originalTo, myEmail) && hasAddressInList_(originalCc, myEmail)) {
    return {
      needsReply: false,
      reason: '自分のアドレスがToに含まれずCCのみ受信のため、返信不要としました。'
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
 * AI出力を固定フォーマットでGmail下書き本文にする
 */
function formatDraftBody_(parsed, rawResponse, riskProfile) {
  const warning = [
    '※以下はAI作成の下書きです。送信前に【要入力】【要回答】を削除・修正してください。',
    '※法律判断・事件方針・見通し・金額・期限は必ず確認してください。'
  ].join('\n');

  if (!parsed || typeof parsed !== 'object') {
    return [
      warning,
      '',
      `【相手種別】${riskProfile.label}`,
      '',
      '【要確認】',
      'AI出力をJSONとして解析できませんでした。以下は送信せず、内容確認用として扱ってください。',
      '',
      '【AI出力】',
      String(rawResponse || '').substring(0, 2000)
    ].join('\n');
  }

  const memo = normalizeStringArray_(parsed.memo, 3, 120);
  const placeholders = normalizeStringArray_(parsed.placeholders, 10, 120);
  const riskFlags = normalizeStringArray_(parsed.riskFlags, 10, 160);
  const draft = String(parsed.draft || '').trim();

  const parts = [
    warning,
    '',
    `【相手種別】${riskProfile.label}`
  ];

  if (memo.length > 0) {
    parts.push('', '【相手の要望メモ】');
    memo.forEach((item) => parts.push(`- ${item}`));
  }

  parts.push('', '【返信本文案】');
  parts.push(draft || 'ご連絡ありがとうございます。\n内容を確認のうえ、改めてご連絡いたします。');

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
