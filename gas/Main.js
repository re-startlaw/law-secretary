/**
 * 設定定数
 */
const CONFIG = {
  SHEET_SETTINGS: '設定',
  SHEET_LOGS: 'ログ',
  LABEL_DRAFT_CREATED: 'AI_ドラフト作成済', 
  LABEL_NO_REPLY: 'AI_返信不要',
  API_PROP_KEY: 'GEMINI_API_KEY'
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('AIメール自動化')
    .addItem('未読メールを処理 (手動実行)', 'processUnreadEmails')
    .addSeparator()
    .addItem('Gemini APIキーの設定', 'setApiKey')
    .addToUi();
}

/**
 * メイン処理
 */
function processUnreadEmails() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.API_PROP_KEY);
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Gemini APIキーが設定されていません。メニューから設定してください。');
    return;
  }

  // 設定シートから値を読み込む（画像の配置に対応）
  const settings = getSettingsFromSheet(); 
  
  // 検索クエリ
 const query = `is:unread in:inbox newer_than:1d -label:${CONFIG.LABEL_DRAFT_CREATED} -label:${CONFIG.LABEL_NO_REPLY}`;
  
  // スレッド取得（最大20件程度処理）
  const threads = GmailApp.search(query, 0, 20);
  
  if (threads.length === 0) {
    console.log('処理対象の未読メールはありません。');
    return;
  }

  const labelDraft = getOrCreateLabel(CONFIG.LABEL_DRAFT_CREATED);
  const labelNoReply = getOrCreateLabel(CONFIG.LABEL_NO_REPLY);

  for (const thread of threads) {
    const messages = thread.getMessages();
    const lastMessage = messages[messages.length - 1];
    
    // 自分のメールはスキップ
    const myEmail = Session.getActiveUser().getEmail();
    if (myEmail === lastMessage.getFrom()) continue;

    const subject = lastMessage.getSubject();
    const body = lastMessage.getPlainBody();
    const sender = lastMessage.getFrom();
    const threadId = thread.getId();
    const mailUrl = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;

    try {
      // 1. 返信要否の判断
      const needsReply = judgeReplyNecessity(apiKey, settings, sender, subject, body);
      
      if (needsReply) {
        // 2. 返信文面の生成
        const draftBody = generateReplyBody(apiKey, settings, sender, subject, body);
        
        // 3. 全員に返信（To/CC維持）
        const originalTo = lastMessage.getTo() || '';
        const originalCc = lastMessage.getCc() || '';
        
        // 自分のアドレスを除外するヘルパー
        const excludeMe = (list) => list.split(',').map(e => e.trim()).filter(e => !e.toLowerCase().includes(myEmail.toLowerCase())).join(', ');
        
        // To: 送信者 + 元のTo（自分以外）
        const toList = [sender, excludeMe(originalTo)].filter(e => e).join(', ');
        // CC: 元のCC（自分以外）
        const ccList = excludeMe(originalCc);
        
        lastMessage.createDraftReply(draftBody, {
          to: toList,
          cc: ccList
        });
        
        thread.addLabel(labelDraft);
        
        logToSheet(sender, subject, 'ドラフト作成成功', mailUrl);
      } else {
        thread.addLabel(labelNoReply);
        logToSheet(sender, subject, '返信不要と判断', mailUrl);
      }
      
    } catch (e) {
      console.error(e);
      logToSheet(sender, subject, `エラー: ${e.message}`, mailUrl);
    }
  }
}

/**
 * 返信要否判定
 */
function judgeReplyNecessity(apiKey, settings, sender, subject, body) {
  const prompt = `
あなたは秘書です。以下のメールに返信が必要か判断してください。
返信が必要なら "TRUE"、不要なら "FALSE" とだけ出力してください。

【判断基準】
${settings.criteria}

【受信メール】
送信者: ${sender}
件名: ${subject}
本文:
${body.substring(0, 3000)}
`;
  const response = callGeminiApi(apiKey, settings.model, prompt);
  return response.trim().toUpperCase().includes('TRUE');
}

/**
 * 返信文生成
 * 文体を実送信メールに寄せるため、相手への過去送信メール（最大5通）を
 * few-shot例としてプロンプトに注入する。履歴がなければ直近送信メールで代替。
 */
function generateReplyBody(apiKey, settings, sender, subject, body) {
  const senderEmail = extractEmail_(sender);
  let examples = fetchSentExamplesForRecipient(senderEmail, 5);
  let exampleHeader;

  if (examples.length > 0) {
    exampleHeader = `【あなたが過去にこの相手（${senderEmail}）に送ったメール（文体を忠実に再現すること）】`;
  } else {
    examples = fetchRecentSentExamples(3);
    exampleHeader = '【あなたの直近送信メール（文体参考）】';
  }

  const examplesBlock = examples.length === 0
    ? '（過去の送信メールが取得できませんでした）'
    : examples.map((e, i) => `---例${i + 1}---\n${e}`).join('\n\n');

  console.log(`few-shot: recipient=${senderEmail}, samples=${examples.length}`);

  const prompt = `
以下のメールへの返信文面（本文のみ）を作成してください。
件名は不要です。

重要: 下記「過去の送信メール」の文体・語彙・句読点・敬称・段落の切り方を
忠実に再現してください。シート上の例文より、過去の送信メールを優先してください。

【あなたの立場・コンテキスト】
${settings.context}

【シート上の文体・例文（参考）】
${settings.example}

${exampleHeader}
${examplesBlock}

【受信メール】
送信者: ${sender}
件名: ${subject}
本文:
${body.substring(0, 3000)}
`;
  return callGeminiApi(apiKey, settings.model, prompt);
}

/**
 * API呼び出し
 */
function callGeminiApi(apiKey, model, text) {
  const useModel = model ? model : 'gemini-1.5-flash';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${useModel}:generateContent?key=${apiKey}`;
  
  const payload = {
    contents: [{ parts: [{ text: text }] }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error(`API Error: ${json.error.message}`);
  }
  
  return json.candidates[0].content.parts[0].text;
}