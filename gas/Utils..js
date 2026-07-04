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
  const model = sheet.getRange('B2').getValue();
  const generationModel = sheet.getRange('B6').getValue();
  return {
    model: model,    // AIモデル（返信要否判定用）
    example: sheet.getRange('B3').getValue(),  // メール例文
    context: sheet.getRange('B4').getValue(),  // コンテキスト情報
    criteria: sheet.getRange('B5').getValue(),  // 返信判断基準
    generationModel: generationModel || model  // AIモデル（本文生成用。未設定ならB2を流用）
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
  const substantiveLines = String(body || '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => {
      if (!line) return false; // 空行除外
      if (/^https?:\/\//i.test(line)) return false; // URLのみの行除外
      if (/^\[?(image|画像)[:：]/i.test(line)) return false; // 画像代替テキストらしい行除外
      return true;
    })
    .slice(0, 5);
  return /米谷|先生/.test(substantiveLines.join('\n'));
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

  const toEmails = extractEmailsFromText_(originalTo).map((e) => e.toLowerCase());
  const ownerEmailsLower = CONFIG.OWNER_EMAILS.map((e) => String(e || '').toLowerCase());
  const myEmailLower = String(myEmail || '').toLowerCase();
  const isAddressedToMe = toEmails.some((e) => e === myEmailLower || ownerEmailsLower.indexOf(e) >= 0);
  if (!isAddressedToMe) {
    return {
      needsReply: false,
      reason: '自分がToに含まれず、CCまたはBCCのみの受信のため返信不要としました。'
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
      '法律判断、事件方針、見通しは通常相手でも必ず本文中に【要確認】を埋め込む。'
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
 * 宛名候補を決定する。
 * 優先順位: ①過去にその相手へ送った自分のメールの冒頭宛名 → ②受信メール本文末尾の署名 → ③From表示名
 * どれも確信が持てなければ null（呼び出し側で「【要確認】様」を使う）
 */
function resolveRecipientName_(senderDisplay, receivedBody, pastSentBodies) {
  // ① 過去送信メールの1行目から宛名を抽出（最新のものを優先）
  for (const sentBody of (pastSentBodies || [])) {
    const firstLine = String(sentBody || '').split(/\r?\n/).map((l) => l.trim()).find((l) => l);
    if (!firstLine) continue;
    const m = firstLine.match(/^(.{1,20}?)(様|先生|さん|御中)/);
    if (m && m[1].trim()) {
      return { name: m[1].trim(), honorific: m[2], source: 'past' };
    }
  }

  // ② 受信本文の末尾署名（最終5行程度）から日本語氏名らしい行を探す
  const bodyLines = String(receivedBody || '')
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l);
  const tail = bodyLines.slice(-5);
  for (let i = tail.length - 1; i >= 0; i--) {
    const line = tail[i];
    const m = line.match(/^([一-龠ぁ-んァ-ヶー]{1,10})(様|先生|さん)?$/);
    if (m && /[一-龠ぁ-んァ-ヶー]/.test(m[1])) {
      return { name: m[1], honorific: m[2] || '様', source: 'signature' };
    }
  }

  // ③ From表示名（"名前 <mail>" の名前部）。日本語を含み、会社名等でなければ採用
  const fromMatch = String(senderDisplay || '').match(/^"?([^"<]+?)"?\s*<[^>]+>\s*$/);
  if (fromMatch) {
    const name = fromMatch[1].trim();
    const looksLikeOrg = /(株式会社|有限会社|合同会社|法律事務所|事務所|センター|協会|組合|株式会社|Inc\.|Corp\.|LLC)/i.test(name);
    if (name && /[一-龠ぁ-んァ-ヶー]/.test(name) && !looksLikeOrg) {
      return { name, honorific: '様', source: 'from' };
    }
  }

  return null;
}

/**
 * 下書き先頭行に宛名（様/先生/御中）がなければ強制挿入するガード
 */
function ensureSalutation_(draftText, resolvedName) {
  const text = String(draftText || '');
  const lines = text.split(/\r?\n/);
  const firstNonEmptyIdx = lines.findIndex((l) => l.trim());
  const firstLine = firstNonEmptyIdx >= 0 ? lines[firstNonEmptyIdx] : '';

  if (/(様|先生|御中)/.test(firstLine)) return text;

  const salutation = resolvedName ? `${resolvedName.name}${resolvedName.honorific}` : '【要確認】様';
  return [salutation, '', text].join('\n');
}

/**
 * AI出力を固定フォーマットでGmail下書き本文にする。
 * 【要入力】【要確認】は本文中にインライン化済みの前提で、署名後の別ブロックは作らない。
 * riskFlags はログ記録用としてのみ返す（下書き本文には含めない）。
 */
function formatDraftBody_(parsed, rawResponse, riskProfile, resolvedName) {
  if (!parsed || typeof parsed !== 'object') {
    // 解析失敗時は送信せず確認用。誤送信防止のため明示する
    return {
      body: [
        '※AI出力をJSONとして解析できませんでした。送信せず内容を確認してください。',
        '',
        '【AI出力】',
        String(rawResponse || '').substring(0, 2000)
      ].join('\n'),
      riskFlags: []
    };
  }

  const riskFlags = normalizeStringArray_(parsed.riskFlags, 10, 160);
  const draft = String(parsed.draft || '').trim()
    || 'ご連絡ありがとうございます。\n内容を確認のうえ、改めてご連絡いたします。';

  const draftWithSalutation = ensureSalutation_(draft, resolvedName);
  const body = [draftWithSalutation, '', '米谷尚起'].join('\n');

  return { body, riskFlags };
}

/**
 * 「推奨設定を適用」で設定シートB2〜B6に書き込む文言。
 * docs/mail_auto_reply_settings.md のマスターと内容を一致させること。
 */
const RECOMMENDED_SETTINGS_ = {
  model: 'gemini-2.5-flash',
  generationModel: 'gemini-2.5-pro',
  example: [
    '渡邊様',
    '',
    'お世話になっております。',
    'ニューフォーカス株式会社との間の契約書レビューが完了致しましたので送付致します。',
    'ご不明な点がございましたら何でもお尋ね下さい。',
    '',
    'それと、差し支えなければ先方にお戻しした契約書のご共有を頂ければ、御社の重視しているポイントの把握等、次回以降のリーガルチェックに役立てさせていただきます。（これを拝見する時間についてはタイムチャージは頂きません。）',
    '',
    '引き続きよろしくお願いいたします。',
    '',
    '',
    '米谷尚起',
    '',
    '【文体利用ルール】',
    '- この例文は文体・長さ・敬称・結びの参考に限る。',
    '- 文章を忠実に模倣せず、相手や案件に応じて簡潔に書く。',
    '- 本文の長さ・温度感は、この例文および過去送信メール例に合わせる。事務的すぎる素っ気ない文面にしない。',
    '- 相手の発言を長くオウム返ししない。事情への反応は原則1文、多くても2文までにする。',
    '  - NG例:「〇月〇日に貴社より契約書のドラフトをお送りいただき、その内容について3点のご質問をいただいた件、確かに拝受いたしました。ご質問の1点目は…」',
    '  - OK例:「契約書ドラフトの件、ご質問3点を確認いたしました。」',
    '- 温度感（気遣い・丁寧さ）は挨拶・結び・気遣いの一文で出す。相手の発言の要約部分では出さない。',
    '- 署名はAI本文に入れず、スクリプト側が下書き本文末尾に「米谷尚起」を自動付与する。'
  ].join('\n'),
  context: [
    '弁護士　米谷尚起が作成するメールです。',
    'no-reply@mail.coconala.comから届いたメールについては、このアドレスではなく、本文中記載の依頼者のメールアドレスに返信する。',
    '【返信文作成の特別ルール】',
    '以下のケースでは、通常の返信と異なる対応を行ってください。',
    '',
    '1. **送信元が "no-reply@mail.coconala.com" の場合**',
    '   - 件名が「[ココナラ法律相談] お問い合わせメールが届いています」の場合、送信元ではなく、本文中に記載されている「依頼者のメールアドレス」宛ての返信文として作成してください。',
    '   - **内容の指示**: 相談内容を読み取り、「刑事事件の加害者側」からの相談であれば受任を検討する前向きな内容にしてください。それ以外（民事、離婚、相続など）の依頼については、現在業務過多のためお引き受けできないという丁寧なお断りのメールを作成してください。',
    '',
    '【安全ルール】',
    '- AIは最終判断者ではなく、送信前に弁護士が必ず確認・修正する下書きだけを作る。',
    '- 法律判断、事件方針、見通しは断定せず、本文中のその箇所に直接「【要確認】」を埋め込む（本文の外に注記を作らない）。',
    '- 金額、期限、日付、提出期限、支払期限は勝手に埋めず、本文中のその箇所に「【要入力：金額】」「【要入力：期限】」「【要入力：日付】」の形で埋め込む。',
    '- 謝罪、譲歩、責任認定、事実認定に見える表現は避ける。',
    '- 裁判所、検察庁、警察、弁護士、相手方代理人、公的機関向けは、短く硬く保留寄りの文面にする。'
  ].join('\n'),
  criteria: [
    '# 判定ロジック（上から順に適用し、該当したら即座に判定を確定すること）',
    '',
    '## 【最優先：返信不要（FALSE）の条件】',
    '以下のいずれかに該当する場合は、内容に関わらず「FALSE」とする。',
    '',
    '1. **ブラックリスト（送信元アドレス）**',
    '   以下のドメインまたはアドレスからのメール：',
    '   - no-reply@printing.ne.jp',
    '   - t-hoso-ml@googlegroups.com',
    '   - shinwazenki@googlegroups.com',
    '   - noreply@tm.openai.com',
    '   - noreply-apps-scripts-notifications@google.com',
    '   - noreply@appsheet.com',
    '   - n.kometani@re-startlaw.com',
    '   - nobukosato2020@gmail.com',
    '   - support@myteam108.jp',
    '',
    '2. **ブラックリスト（件名）**',
    '   件名に以下の文字列を含むもの：',
    '   - [t-hoso-ml:',
    '   - [shinwazenki',
    '',
    '3. **自分自身への送信**',
    '   送信元（From）が自分のメールアドレスである場合。',
    '',
    '4. **CC受信**',
    '   自分のメールアドレスが「To」に含まれておらず、「CC」または「BCC」にのみ含まれている場合。',
    '',
    '5. **宛名不在**',
    '   メール本文冒頭の実質的な行（空行・URLのみの行・画像代替テキストのような行を除いた先頭5行）に「米谷」または「先生」という氏名の記載が見当たらない場合。',
    '',
    '6. **自動送信・一斉配信**',
    '   - 自動生成されたシステムメッセージ',
    '   - 一方的なお知らせ、ニュースレター、メルマガ',
    '   - メーリングリスト経由の周知メール',
    '',
    '## 【返信必要（TRUE）の条件】',
    '上記の「返信不要」条件に該当しない場合で、以下に該当するものは「TRUE」とする。',
    '',
    '1. **ココナラ法律相談**',
    '   送信元が "no-reply@mail.coconala.com" で、件名が「[ココナラ法律相談] お問い合わせメールが届いています」の場合。',
    '   ※システムメールだが、これは例外的に返信が必要。',
    '',
    '2. **質問・依頼・重要連絡**',
    '   - 相手からの質問や依頼が含まれている。',
    '   - 個人的なメッセージや重要な業務連絡である。',
    '',
    '## 【デフォルト】',
    '上記いずれにも明確に当てはまらない場合は、迷ったら下書き作成対象にする。ただし、明確なFALSE条件がある場合はそちらを優先する。'
  ].join('\n')
};

/**
 * 設定シートB2〜B6に推奨文言を書き込む（実行前に確認ダイアログ）
 */
function applyRecommendedSettings() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '推奨設定を適用',
    'B2〜B6に推奨のプロンプト文言・モデル設定を書き込みます。現在の内容は上書きされます。よろしいですか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (response !== ui.Button.OK) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_SETTINGS);
  if (!sheet) {
    ui.alert(`「${CONFIG.SHEET_SETTINGS}」シートが見つかりません。`);
    return;
  }

  sheet.getRange('B2').setValue(RECOMMENDED_SETTINGS_.model);
  sheet.getRange('B3').setValue(RECOMMENDED_SETTINGS_.example);
  sheet.getRange('B4').setValue(RECOMMENDED_SETTINGS_.context);
  sheet.getRange('B5').setValue(RECOMMENDED_SETTINGS_.criteria);
  sheet.getRange('A6').setValue('生成用AIモデル');
  sheet.getRange('B6').setValue(RECOMMENDED_SETTINGS_.generationModel);
  sheet.getRange('C6').setValue('返信文生成に使うAIモデル（返信要否判定はB2のモデルを使用）');

  ui.alert('推奨設定を適用しました。');
}
