// ==========================================
// 【基本設定 & 定数】
// ==========================================
const MASTER_FOLDER_ID = '14kFL7HzFzzdiisi5_QKUEcX8PzrPXu-M';
const TEMPLATE_DOC_ID  = '1XSAgBI3c9Vm6Y_j0CLRtKeMHYhYN91MQefq6P8TvTXM';
const MODEL_FLASH      = 'gemini-3.8-flash';
const MODEL_PRO        = 'gemini-3.1-pro'; 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
  .setTitle('⚖️ 書面作成Cockpit')
  .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function getGeminiApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY がスクリプトプロパティに設定されていません。');
  return apiKey;
}

// ==========================================
// 【基本フォルダ・ファイル操作】
// ==========================================
function getOrCreateCaseFolder(caseName) {
  if (!caseName) caseName = "未分類事件";
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 10秒間、他の処理を待機させる（排他制御）
    const masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    
    let folders = masterFolder.getFoldersByName(caseName);
    if (folders.hasNext()) return folders.next();
    
    // 揺らぎを考慮した検索（「請求」の有無による重複作成防止）
    let altName = caseName;
    if (caseName.match(/(養育費|婚姻費用分担|財産分与|慰謝料)調停事件/)) {
      altName = caseName.replace(/(養育費|婚姻費用分担|財産分与|慰謝料)調停事件/g, '$1請求調停事件');
    } else if (caseName.match(/(養育費|婚姻費用分担|財産分与|慰謝料)審判事件/)) {
      altName = caseName.replace(/(養育費|婚姻費用分担|財産分与|慰謝料)審判事件/g, '$1請求審判事件');
    } else if (caseName.match(/(養育費|婚姻費用分担|財産分与|慰謝料)請求調停事件/)) {
      altName = caseName.replace(/(養育費|婚姻費用分担|財産分与|慰謝料)請求調停事件/g, '$1調停事件');
    } else if (caseName.match(/(養育費|婚姻費用分担|財産分与|慰謝料)請求審判事件/)) {
      altName = caseName.replace(/(養育費|婚姻費用分担|財産分与|慰謝料)請求審判事件/g, '$1審判事件');
    }
    
    if (altName !== caseName) {
      let altFolders = masterFolder.getFoldersByName(altName);
      if (altFolders.hasNext()) {
        let folder = altFolders.next();
        // 見つけた古いフォルダを要求された caseName にリネームして統合する
        folder.setName(caseName);
        let oldSettings = folder.getFilesByName('【設定データ】' + altName);
        if (oldSettings.hasNext()) {
          oldSettings.next().setName('【設定データ】' + caseName);
        }
        return folder;
      }
    }
    
    return masterFolder.createFolder(caseName);
  } catch (e) {
    const masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const folders = masterFolder.getFoldersByName(caseName);
    if (folders.hasNext()) return folders.next();
    return masterFolder.createFolder(caseName);
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateLogSheet(caseName) {
  const caseFolder = getOrCreateCaseFolder(caseName);
  const fileName = '【対話履歴】' + caseName;
  const files = caseFolder.getFilesByName(fileName);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(ss.getId()).moveTo(caseFolder); 
    const sheet = ss.getActiveSheet();
    sheet.appendRow(['最終更新日時', '事件名', '送信者', 'メッセージ']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
  }
  return ss;
}

function getCaseFolderTree(caseName) {
  if (!caseName) caseName = "未分類事件";
  const caseFolder = getOrCreateCaseFolder(caseName);
  let draftFolder, finalFolder, evidenceFolder;
  const draftFolders = caseFolder.getFoldersByName('01_下書き');
  draftFolder = draftFolders.hasNext() ? draftFolders.next() : caseFolder.createFolder('01_下書き');
  const finalFolders = caseFolder.getFoldersByName('02_清書');
  finalFolder = finalFolders.hasNext() ? finalFolders.next() : caseFolder.createFolder('02_清書');
  const evidenceFolders = caseFolder.getFoldersByName('03_画像');
  evidenceFolder = evidenceFolders.hasNext() ? evidenceFolders.next() : caseFolder.createFolder('03_画像');
  return { root: caseFolder, draft: draftFolder, final: finalFolder, evidence: evidenceFolder };
}

function getNextEvidenceNumber(caseName, prefixType) {
  if (!caseName) return 1;
  const masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  const folders = masterFolder.getFoldersByName(caseName);
  if (!folders.hasNext()) return 1; 
  const caseFolder = folders.next();
  let maxNum = 0;
  const searchFolders = [];
  const drafts = caseFolder.getFoldersByName('01_下書き');
  if (drafts.hasNext()) searchFolders.push(drafts.next());
  const evidences = caseFolder.getFoldersByName('03_画像');
  if (evidences.hasNext()) searchFolders.push(evidences.next());
  
  searchFolders.forEach(folder => {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const fileName = files.next().getName();
      const regex = new RegExp(`【.*_${prefixType}第([0-9０-９]+)号証】`);
      const match = fileName.match(regex);
      if (match) {
        const numStr = match[1].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
        const num = parseInt(numStr, 10);
        if (num > maxNum) maxNum = num;
      }
    }
  });
  return maxNum + 1;
}

function getSavedCaseNames() {
  const masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  const folders = masterFolder.getFolders();
  let caseList = [];
  while (folders.hasNext()) {
    const folder = folders.next();
    const name = folder.getName();
    // 未分類事件やシステム用フォルダを除外
    if (name.includes('【Data】') || name.includes('【設定データ】') || name === '未分類事件') continue;
    caseList.push({ name: name, date: folder.getDateCreated() });
  }
  // 名前順（昇順）にソート
  caseList.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
  return caseList.map(c => c.name);
  }

function saveEvidenceFile(caseName, fileName, base64Data, mimeType) {
  const tree = getCaseFolderTree(caseName);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file = tree.evidence.createFile(blob);
  // リンクを知っている全員が閲覧可能に設定（アカウント選択回避）
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return true;
}

// tempフォルダ取得・作成
function getOrCreateTempFolder(caseName) {
  const tree = getCaseFolderTree(caseName);
  const tempFolders = tree.evidence.getFoldersByName('temp');
  if (tempFolders.hasNext()) return tempFolders.next();
  return tree.evidence.createFolder('temp');
}

// temp画像アップロード＆OCR
function uploadTempImage(caseName, base64Data, mimeType, needOcr) {
  try {
    const tempFolder = getOrCreateTempFolder(caseName);
    
    // --- 古いtempファイルをクリーンアップ（3時間以上前のもの） ---
    try {
      const now = new Date().getTime();
      const files = tempFolder.getFiles();
      while (files.hasNext()) {
        const f = files.next();
        if (now - f.getDateCreated().getTime() > 3 * 60 * 60 * 1000) {
          f.setTrashed(true);
        }
      }
    } catch(ce) { console.warn('temp cleanup failed:', ce); }
    // ----------------------------------------------------
    
    const tempName = 'temp_img_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000) + '.jpg';
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, tempName);
    const file = tempFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    let ocrResult = null;
    // needOcrが明示的にfalseでない限りOCRを実行する
    if (needOcr !== false) {
      ocrResult = performOcrAndExtract(base64Data, mimeType, null);
    }
    return { success: true, fileId: file.getId(), tempName: tempName, ocrData: ocrResult };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// temp画像削除
function deleteTempImage(fileId) {
  if (!fileId) return { success: false };
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 【通信・Gemini連携・書証生成】
// ==========================================
function fetchWithRetry(url, options, maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    try { return UrlFetchApp.fetch(url, options); } 
    catch (e) { if (i === maxRetries - 1) throw e; Utilities.sleep(2000 * (i + 1)); }
  }
}

function callGeminiApiWithFallback(payload, initialModel) {
  const API_KEY = getGeminiApiKey();
  // 指定モデルがない場合はデフォルトで gemini-3.8-flash (MODEL_FLASH) を使用
  let currentModel = initialModel || MODEL_FLASH; 
  const maxRetries = 3;

  for (let i = 0; i < maxRetries; i++) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${currentModel}:generateContent?key=` + API_KEY;
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      
      if (statusCode === 200) {
        return response; // 成功時はそのまま返却
      } else if (statusCode === 503 || statusCode === 429) {
        console.warn(`[${statusCode} Error] gemini-3.8-flashでリトライします(${i+1}/${maxRetries})`);
        currentModel = MODEL_FLASH; // エラー時は gemini-3.8-flash に切り替えてリトライ
        Utilities.sleep(3000 * (i + 1)); // 待機時間を徐々に延ばす
        continue; 
      } else {
        throw new Error(`APIエラー (${statusCode}): ${response.getContentText()}`);
      }
    } catch (e) {
      if (i === maxRetries - 1) throw e; 
      currentModel = MODEL_FLASH;
      Utilities.sleep(3000 * (i + 1));
    }
  }
  throw new Error('APIの制限または一時的なエラーが連続しました。しばらく時間を置いてから再度お試しください。');
}

function performOcrAndExtract(base64Data, mimeType, driveFileId) {
  try {
    let docId = null;
    let tempFileId = null;
    
    // 1. GoogleドキュメントのOCR機能を使ってテキスト抽出
    if (driveFileId) {
      const url = "https://www.googleapis.com/drive/v3/files/" + driveFileId + "/copy?ocrLanguage=ja";
      const options = {
        method: "post",
        headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ mimeType: "application/vnd.google-apps.document" }),
        contentType: "application/json",
        muteHttpExceptions: true
      };
      const res = UrlFetchApp.fetch(url, options);
      const docData = JSON.parse(res.getContentText());
      if (!docData.id) throw new Error(docData.error ? docData.error.message : "OCRエラー");
      docId = docData.id;
    } else if (base64Data) {
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, "temp_ocr_image");
      const tempFile = DriveApp.createFile(blob);
      tempFileId = tempFile.getId();
      
      const url = "https://www.googleapis.com/drive/v3/files/" + tempFileId + "/copy?ocrLanguage=ja";
      const options = {
        method: "post",
        headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ mimeType: "application/vnd.google-apps.document" }),
        contentType: "application/json",
        muteHttpExceptions: true
      };
      const res = UrlFetchApp.fetch(url, options);
      const docData = JSON.parse(res.getContentText());
      docId = docData.id;
    }
    
    if (!docId) throw new Error("OCR用ドキュメントの生成に失敗しました");
    
    const text = DocumentApp.openById(docId).getBody().getText();
    
    // 一時ファイルのクリーンアップ
    DriveApp.getFileById(docId).setTrashed(true);
    if (tempFileId) DriveApp.getFileById(tempFileId).setTrashed(true);
    
    if (!text || text.trim() === "") return { success: false, message: "テキストが検出されませんでした" };
    
    // 2. GeminiによるJSON構造化抽出
    const prompt = `以下のOCRテキストから、文書の「標題」「作成日」「作成者」を抽出してください。
    【ルール】
    - 該当データのみを持つJSON形式で出力してください。（Markdownコードブロック不要）
    - キーは "標題", "作成日", "作成者" としてください。
    - 「標題」について、明確なタイトルが記載されていない場合は、内容を端的に表す要約（例：「LINEトーク履歴」「〇〇の通知書」「〇〇の写真」「封筒」など）を自動生成して設定してください。
    - 該当データが存在せず、推測も不可能な項目は空文字("")にしてください。
    - 旧字体（例：診斷書）が含まれている場合は、現代の新字体（例：診断書）に変換して抽出してください。
    - 和暦がある場合は極力そのままの形式で抽出してください。
    - 「作成日」の抽出ルール：
      ・明確な日付の記載がない場合でも、LINEのトーク履歴など複数の日付・時刻や曜日が混在している場合は、トークが行われた日付（複数ある場合は最も古い日付や代表的な日付）を推測して「YYYY年M月D日」等の形式で抽出してください。
      ・和暦（令和、平成、昭和など）に空白が含まれる場合（例：「令　和」）は空白を削除し、「和〇年」や「〇月〇日」のように一部が欠損しているように見える場合は文脈から正しい元号や年を推測して補完してください。
      ・未記入の欄を示す「〇」などの記号はそのまま「〇」として抽出してください。
    - 「作成者」の抽出ルール：
      1. 医療機関等の場合は「(病院・施設名) (医師・個人名)」のようにスペース区切りで結合してください。
      2. 行政機関や公的機関で「（狭山市福祉事務所）（埼玉県狭山市福祉事務所長）」のように(施設名)と(役職名)がほとんど重複する場合は施設名を省略してください。また役職名が長い場合は改行されている可能性を加味してください。
    【OCRテキスト】
    ${text}`;
    
    const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
    const geminiRes = callGeminiApiWithFallback(payload, MODEL_FLASH); // Flashモデルを使用
    const jsonStr = JSON.parse(geminiRes.getContentText()).candidates[0].content.parts[0].text;
    
    const cleanJsonStr = jsonStr.replace(/```(?:json)?\n?([\s\S]*?)\n?```/g, '$1').trim();
    const extractedData = JSON.parse(cleanJsonStr);
    
    return { success: true, data: extractedData };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function generateDocumentPreview(docType, variablesText, selectedModel) {
  if (!selectedModel) selectedModel = MODEL_PRO;
  let templateText = "";
  if (docType === "準備書面") {
    templateText = `{{事件番号}} {{事件名}}
    {{当事者1}} {{当事者名1}}
    {{当事者2}} {{当事者名2}}
    # **{{準備書面名}}**
    {{提出日}}
    {{裁判所名}} 御中
    作成者 土倉功司 印
    {{本文}}`;
  } else if (docType === "証拠説明書") {
    templateText = `{{事件番号}} {{事件名}}
    {{当事者1}} {{当事者名1}}
    {{当事者2}} {{当事者名2}}
    # **{{証拠説明書名}}**
    {{提出日}}
    {{裁判所名}} 御中
    作成者 土倉功司 印
    {{本文}}`;
  }
  const prompt = `あなたは優秀な日本の弁護士・法務スタッフです。
    以下の【変数データ】を【テンプレート】に当てはめて、【作成対象】のテキストを作成してください。
    【出力ルール】
    - 挨拶、前置き、補足説明などは一切不要です。
    - テンプレートに変数データを代入した結果の「本文のみ」を出力してください。
    - テンプレートおよび本文（{{本文}}部分に代入されるテキスト）内に含まれる変数のプレースホルダー（例：「{{提出日}}」「{{準備書面名}}」「{{証拠説明書名}}」等）は、変数データの該当項目に必ず置き換えてください（プレースホルダーの括弧の有無に関わらず一致させてください）。
    - 変数データに存在しない項目は、そのまま空欄（空白）にしておくか、「 」などで補ってください。
    - 【重要：スペースと記号の使い分け】
    1. 通常の本文（表以外の文章、字下げ、氏名の間など）のスペースは、必ず「全角スペース」を使用してください。
    2. 表（マークダウン形式の表）の中のデータにおけるスペースは、必ず「半角スペース」を使用してください。
    3. マークダウンの記号（ | や - など）は必ず半角を使用してください。
    - テンプレートのレイアウトや構成を崩さず、そのまま出力してください。
    - Markdownコードブロック (\`\`\`markdown) で囲まないでください。
    【作成対象】: ${docType}
    【テンプレート】:
    ${templateText}
    【変数データ】:
    ${variablesText}`;
    const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
    try {
    const response = callGeminiApiWithFallback(payload, selectedModel);
    const json = JSON.parse(response.getContentText());
    let text = json.candidates[0].content.parts[0].text;

    text = text.replace(/｜/g, '|'); // Geminiが誤って出力した全角パイプを半角に統一
    text = text.replace(/^[\s\S]*?(?:以下の通りです|作成したテキスト|出力します)[^\n]*\n+/, '');
    text = text.replace(/```(?:markdown)?\n?/g, '').replace(/```\n?/g, '');

    const toZenkakuSafe = (str) => {
      return str.split('\n').map(line => {
        if (line.trim().startsWith('|') && line.trim().endsWith('|')) return line;
        let prefix = "";
        if (line.match(/^[#＃][ \u3000]/)) { prefix = "# "; line = line.substring(2); }
        
        // 引用記号と太字マークダウンの削除
        line = line.replace(/\[?\s*cite\s*[:：]\s*[0-9０-９,\s，]*\]?/gi, '');
        line = line.replace(/[*＊]/g, '');
        
        line = line.replace(/[!-~]/g, function(s) {
          if (s === '(') return '（'; if (s === ')') return '）'; if (s === '（' || s === '）') return s;
          return String.fromCharCode(s.charCodeAt(0) + 0xFEE0);
        }).replace(/[ \x20\u0020\xA0\u200B\t]/g, '\u3000'); // 半角スペース等を強制的に全角スペース(\u3000)に変更
        return prefix + line;
      }).join('\n');
    };

    text = toZenkakuSafe(text);
    return text.trim();
  } catch (e) { return "プレビュー生成エラー: " + e.toString(); }
}

function applyLegalStylesToBody(doc, body, content, title, images, evidenceFolder, caseName) {
  title = title || "";
  // 【下書き】や【要検証_下書き】などの接頭辞を除外して純粋なファイル名で判定する
  let pureTitleForCheck = title.replace(/^【[^】]+】/, '').trim();
  let isEvidence = pureTitleForCheck.match(/^[甲乙丙][第]?[０-９0-9]+号証/) !== null;
  let isEvidenceList = title.includes('証拠説明書');
  // 初期化（元の内容をクリア）
  while (body.getNumChildren() > 1) {
    body.removeChild(body.getChild(0));
  }
  body.getChild(0).asParagraph().clear();
  // 文書の種類ごとに専用の関数へ処理を完全に分岐
  if (isEvidence) {
    applyStylesToEvidence(doc, body, content, title, images);
  } else if (isEvidenceList) {
    applyStylesToEvidenceList(doc, body, content, title);
  } else {
    applyStylesToBrief(doc, body, content, title);
  }
  // 共通のクリーンアップ処理
  if (body.getNumChildren() > 1 && body.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH && body.getChild(0).asParagraph().getText() === "") {
    body.removeChild(body.getChild(0));
  }
  let numChildren = body.getNumChildren();
  if (numChildren > 0) {
    let lastChild = body.getChild(numChildren - 1);
    if (lastChild.getType() === DocumentApp.ElementType.PARAGRAPH && lastChild.asParagraph().getText().trim() === '') {
      let p = lastChild.asParagraph();
      p.setFontSize(1);
      p.setLineSpacing(1.0);
    }
  }
}

// ==========================================
// 【準備書面 専用処理】
// ==========================================
function applyStylesToBrief(doc, body, content, title) {
  const lines = content.split('\n');
  const baseFont = 'Noto Sans JP';
  const baseSize = 12;
  
  let currentLevel = 0;
  let levelStates = [ { firstLineChars: 0, startChars: 0, hasParen: false } ];
  let isFirstLine = true;
  let hasIjo = false;
  let lastWasEmpty = false;
  let isAfterHeader = false;

  const toZenkaku = (str) => {
    return str.replace(/[!-\~]/g, function(s) {
      if (s === '(') return '（';
      if (s === ')') return '）';
      if (s === '（' || s === '）') return s;
      return String.fromCharCode(s.charCodeAt(0) + 0xFEE0);
    }).replace(/[ \x20\u0020\xA0\u200B\t]/g, '\u3000');
  };

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    line = line.replace(/\\?\[?[CcＣｃ][IiＩｉ][TtＴｔ][EeＥｅ][:：][\s\u3000]*[０-９0-9,，\s\u3000]*[\]］]?/g, '');
    
    if (line === '【改ページ】') {
      let pb = body.appendPageBreak();
      pb.getParent().asParagraph().setFontSize(1);
      let pEmpty = body.appendParagraph(' ');
      pEmpty.setFontSize(12).setFontFamily(baseFont).setLineSpacing(1.0);
      lastWasEmpty = true;
      isFirstLine = false;
      continue;
    }
    
    if (line === '') {
      isAfterHeader = false;
      continue;
    }
    
    let isTitle = false;
    if (line.match(/^[\#＃][ \u3000]+/)) {
      isTitle = true;
      line = line.replace(/^[\#＃][ \u3000]+/, '');
    }
    
    line = line.replace(/[*＊]/g, '');
    line = toZenkaku(line);
    if (line === '以上') hasIjo = true;
    
    let mDai = line.match(/^(第[０-９一-九十百]+)(?:[\u3000]|$)/);
    if (mDai && !isFirstLine && !lastWasEmpty && !isAfterHeader) {
      if (mDai[1] !== '第１' && mDai[1] !== '第一') {
        body.appendParagraph('').setLineSpacing(1.5);
      }
    }
    
    // 表直後の空行を再利用するロジックを削除し、純粋に段落を追加
    let p = body.appendParagraph(line);
    
    p.setFontFamily(baseFont).setFontSize(baseSize);
    if (line === ' ') {
      p.setFontSize(1);
      p.setLineSpacing(1.0);
    } else {
      p.setLineSpacing(1.5);
    }
    
    p.setIndentEnd(-6.8);
    p.setBold(false);
    lastWasEmpty = false;
    isFirstLine = false;
    
    if (isTitle) {
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      p.setBold(true);
      p.setFontSize(14);
      isAfterHeader = true;
      continue;
    } else if (line.endsWith('御中') || line.endsWith('殿')) {
      p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      isAfterHeader = true;
      continue;
    } else if (line.endsWith('印') || line === '以上' || line.replace(/[\s\u3000]+/g, '').match(/^[令和平成昭和]+[０-９元]+年[０-９]+月[０-９]+日$/)) {
      p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      p.setIndentEnd(0); // 右インデントを本文枠に合わせるためリセット
      if (line.endsWith('印')) {
        let textObj = p.editAsText();
        let targetIndex = textObj.getText().lastIndexOf('印');
        if (targetIndex !== -1) {
          textObj.setForegroundColor(targetIndex, targetIndex, '#b7b7b7');
        }
      }
      isAfterHeader = true;
      continue;
    }
    
    isAfterHeader = false;
    
    let mParen = line.match(/^([（『【《][０-９Ａ-Ｚａ-ｚア-ンあ-ん]+[）』】》])/);
    let mNoParen = line.match(/^([０-９Ａ-Ｚａ-ｚア-ンあ-ん]+)[．\u3000]/);
    
    let isList = false;
    let firstLineChars = 0;
    let startChars = 0;
    let hasParen = false;
    
    if (mDai) {
      currentLevel = 1; isList = true; hasParen = false; firstLineChars = 0; startChars = 2; 
      p.setBold(true);
      levelStates[1] = { firstLineChars: firstLineChars, startChars: startChars, hasParen: hasParen };
    } else if (mNoParen && mNoParen[1].match(/^[０-９]+$/)) {
      currentLevel = 2; isList = true; hasParen = false; firstLineChars = 1; startChars = 2; 
      levelStates[2] = { firstLineChars: firstLineChars, startChars: startChars, hasParen: hasParen };
    } else if (mParen && mParen[1].match(/^[（][０-９]+[）]$/)) {
      currentLevel = 3; isList = true; hasParen = true; firstLineChars = 0; startChars = 2; 
      levelStates[3] = { firstLineChars: firstLineChars, startChars: startChars, hasParen: hasParen };
    } else if (mNoParen && mNoParen[1].match(/^[ア-ン]+$/)) {
      currentLevel = 4; isList = true; hasParen = false; firstLineChars = 2; startChars = 3; 
      levelStates[4] = { firstLineChars: firstLineChars, startChars: startChars, hasParen: hasParen };
    } else if (mParen && mParen[1].match(/^[（][ア-ン]+[）]$/)) {
      currentLevel = 5; isList = true; hasParen = true; firstLineChars = 1; startChars = 3; 
      levelStates[5] = { firstLineChars: firstLineChars, startChars: startChars, hasParen: hasParen };
    } else if (mParen || mNoParen) {
      isList = true;
      hasParen = !!mParen;
      let parentState = levelStates[Math.max(0, currentLevel - 1)] || { firstLineChars: 0, startChars: 0, hasParen: false };
      if (parentState.hasParen && !hasParen) {
        firstLineChars = Math.max(0, parentState.firstLineChars + 1);
      } else if (!parentState.hasParen && hasParen) {
        firstLineChars = Math.max(0, parentState.firstLineChars - 1);
      } else {
        firstLineChars = parentState.firstLineChars;
      }
      let len = mParen ? mParen[1].length - 1 : mNoParen[1].length;
      startChars = firstLineChars + len;
      levelStates[currentLevel] = { firstLineChars: firstLineChars, startChars: startChars, hasParen: hasParen };
    } else {
      if (currentLevel === 0) {
        firstLineChars = 0; startChars = 0;
      } else {
        let state = levelStates[currentLevel] || { firstLineChars: 0, startChars: 0, hasParen: false };
        firstLineChars = (currentLevel <= 3) ? 3 : 4;
        startChars = state.startChars;
      }
    }
    p.setIndentFirstLine(firstLineChars * 12);
    p.setIndentStart(startChars * 12);
  }
  
  if (!hasIjo) {
    let p = body.appendParagraph('以上');
    p.setFontFamily(baseFont).setFontSize(baseSize);
    p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  }
}

// ==========================================
// 【証拠説明書 専用処理】
// ==========================================
function applyStylesToEvidenceList(doc, body, content, title) {
  body.setMarginTop(56.7);
  body.setMarginBottom(56.7);
  body.setMarginLeft(70.875);
  body.setMarginRight(56.7);

  const lines = content.split('\n');
  const baseFont = 'Noto Sans JP';
  const baseSize = 12;
  
  let inTable = false;
  let tableData = [];
  let isAfterHeader = false;

  const toZenkaku = (str) => {
    return str.replace(/[!-\~]/g, function(s) {
      if (s === '(') return '（';
      if (s === ')') return '）';
      if (s === '（' || s === '）') return s;
      return String.fromCharCode(s.charCodeAt(0) + 0xFEE0);
    }).replace(/[ \x20\u0020\xA0\u200B\t]/g, '\u3000');
  };

  const flushTable = () => {
    if (inTable && tableData.length > 0) {
      let pSpace = body.appendParagraph(' ');
      pSpace.setFontSize(6).setLineSpacing(1.0);
      
      let table = body.appendTable(tableData);
      
      let paddingPt = 2.835; 
      let cols = table.getRow(0).getNumCells();
      table.setBorderWidth(1);
      
      let colRoles = [];
      for (let c = 0; c < cols; c++) {
        let text = table.getRow(0).getCell(c).getText().trim();
        if (text.includes('号証') || text === '') colRoles[c] = 'num';
        else if (text.includes('標題')) colRoles[c] = 'title';
        else if (text.includes('作成年月日') || text.includes('作成日')) colRoles[c] = 'date';
        else if (text.includes('作成者')) colRoles[c] = 'creator';
        else if (text.includes('立証趣旨')) colRoles[c] = 'purpose';
        else if (text.includes('備考')) colRoles[c] = 'remark';
        else colRoles[c] = 'other';
      }
      
      let remarkIsEmpty = true;
      let remarkCol = colRoles.indexOf('remark');
      if (remarkCol !== -1) {
        for (let r = 1; r < table.getNumRows(); r++) {
          if (table.getRow(r).getCell(remarkCol).getText().replace(/[\s\u3000\u200B\n\r]/g, '') !== '') {
            remarkIsEmpty = false;
            break;
          }
        }
      }
      
      let colWidths = Array(cols).fill(0);
      let fixedWidthSum = 0;
      let flexibleCols = 0;
      
      for (let c = 0; c < cols; c++) {
        let role = colRoles[c];
        if (role === 'num') { colWidths[c] = 21; fixedWidthSum += 21; } 
        else if (role === 'date') { colWidths[c] = 60; fixedWidthSum += 60; }
        else if (role === 'creator') { colWidths[c] = 90; fixedWidthSum += 90; }
        else if (role === 'remark') {
          let w = remarkIsEmpty ? 21 : 85; 
          colWidths[c] = w; fixedWidthSum += w;
        } else {
          flexibleCols++; 
        }
      }
      
      let flexibleWidth = (467.7 - fixedWidthSum) / (flexibleCols > 0 ? flexibleCols : 1);
      for (let c = 0; c < cols; c++) {
        if (colWidths[c] === 0) colWidths[c] = flexibleWidth;
      }
      
      const toHankaku = (str) => {
        return str.replace(/[Ａ-Ｚａ-ｚ０-９．／：]/g, function(s) {
          return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
        });
      };
      
      for (let r = 0; r < table.getNumRows(); r++) {
        let row = table.getRow(r);
        for (let c = 0; c < row.getNumCells(); c++) {
          let cell = row.getCell(c);
          let role = colRoles[c];
          cell.setWidth(colWidths[c]);
          cell.setPaddingTop(paddingPt).setPaddingBottom(paddingPt).setPaddingLeft(paddingPt).setPaddingRight(paddingPt);
          
          let pCount = cell.getNumChildren();
          for (let k = 0; k < pCount; k++) {
            let cp = cell.getChild(k).asParagraph();
            cp.setLineSpacing(1.0);
            cp.setSpacingBefore(0).setSpacingAfter(0);
            
            if (r === 0) {
              if (k === 0) {
                let text = cp.getText().trim();
                if (text !== '') {
                  text = text.replace(/ /g, ' '); 
                  if (role === 'title') {
                    if (!text.includes('\n') && text.includes('（')) text = text.replace('（', '\n（'); 
                    cp.setText(text);
                  } else if (role === 'date') {
                    cp.setText('作成日'); 
                  } else {
                    cp.setText(text);
                  }
                }
              }
              cell.setBackgroundColor('#f3f3f3');
              cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
              cp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            } else {
              if (k === 0) {
                let text = cp.getText();
                if (text !== '') {
                  text = text.replace(/ /g, ' '); 
                  if (role === 'date') {
                    cp.setText(toHankaku(text)); 
                  } else {
                    cp.setText(text);
                  }
                }
              }
              if (role === 'num') {
                cp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
              }
            }
            let textObj = cp.editAsText();
            textObj.setFontFamily('Noto Sans JP').setFontSize(10.5);
            if (r === 0) textObj.setBold(true);
          }
        }
      }
      tableData = [];
    }
    inTable = false;
  };

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    line = line.replace(/\\?\[?[CcＣｃ][IiＩｉ][TtＴｔ][EeＥｅ][:：][\s\u3000]*[０-９0-9,，\s\u3000]*[\]］]?/g, '');

    if (line === '【改ページ】') {
      flushTable();
      let pb = body.appendPageBreak();
      pb.getParent().asParagraph().setFontSize(1);
      let pEmpty = body.appendParagraph(' ');
      pEmpty.setFontSize(12).setFontFamily(baseFont).setLineSpacing(1.0);
      continue;
    }
    
    if (line.startsWith('|') && line.endsWith('|')) {
      inTable = true;
      if (!line.match(/^\|[-: | \u3000]+\|$/)) {
        let rowData = line.split('|').slice(1, -1).map(cell => cell.trim());
        if (rowData.some(cell => cell !== '')) {
          tableData.push(rowData);
        }
      }
      continue;
    } else if (inTable) {
      flushTable();
    }
    
    if (line === '') {
      isAfterHeader = false;
      continue;
    }
    
    let isTitle = false;
    if (line.match(/^[\#＃][ \u3000]+/)) {
      isTitle = true;
      line = line.replace(/^[\#＃][ \u3000]+/, '');
    }
    
    line = line.replace(/[*＊]/g, '');
    line = toZenkaku(line);
    
    let p = body.appendParagraph(line);
    p.setFontFamily(baseFont).setFontSize(baseSize);
    p.setLineSpacing(1.5);
    p.setIndentEnd(-6.8);
    p.setBold(false);
    
    if (isTitle) {
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      p.setBold(true);
      p.setFontSize(14);
      isAfterHeader = true;
      continue;
    } else if (line.endsWith('御中') || line.endsWith('殿')) {
      p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      isAfterHeader = true;
      continue;
    } else if (line.endsWith('印') || line === '以上' || line.replace(/[\s\u3000]+/g, '').match(/^[令和平成昭和]+[０-９元]+年[０-９]+月[０-９]+日$/)) {
      p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      p.setIndentEnd(0); // 右インデントを本文枠に合わせるためリセット
      if (line.endsWith('印')) {
        let textObj = p.editAsText();
        let targetIndex = textObj.getText().lastIndexOf('印');
        if (targetIndex !== -1) {
          textObj.setForegroundColor(targetIndex, targetIndex, '#b7b7b7');
        }
      }
      isAfterHeader = true;
      continue;
    }
    isAfterHeader = false;
  }
  flushTable();
  
  // 表の直下の空白の改行をフォントサイズ6ptにする
  let numChildren = body.getNumChildren();
  if (numChildren > 0) {
    let lastChild = body.getChild(numChildren - 1);
    if (lastChild.getType() === DocumentApp.ElementType.PARAGRAPH && lastChild.asParagraph().getText() === '') {
      lastChild.asParagraph().setFontSize(6).setLineSpacing(1.0);
    } else {
      body.appendParagraph('').setFontSize(6).setLineSpacing(1.0);
    }
  } else {
    body.appendParagraph('').setFontSize(6).setLineSpacing(1.0);
  }
  
  // その直下の行にNoto Sans JP、12pt、右寄せで『以上』を記述
  let pIjo = body.appendParagraph('以上');
  pIjo.setFontFamily(baseFont).setFontSize(baseSize).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
}

// ==========================================
// 【書証（画像） 専用処理】
// ==========================================
function applyStylesToEvidence(doc, body, content, title, images) {
  body.setMarginTop(56.7);
  body.setMarginBottom(56.7);
  body.setMarginLeft(70.875);
  body.setMarginRight(56.7);

  let header = doc.getHeader();
  if (!header) header = doc.addHeader();
  header.clear();
  // 【下書き】や【要検証_下書き】などを確実に除外する
  let pureTitle = title.replace(/【[^】]+】/g, '').replace('.pdf','').trim();
  let hp = header.appendParagraph(pureTitle);
  hp.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  hp.setBold(true);
  hp.setFontSize(14);
  hp.setLineSpacing(1.15);
  hp.setFontFamily('Noto Sans JP');

  const lines = content.split('\n');
  const baseFont = 'Noto Sans JP';
  const baseSize = 12;
  
  let inTable = false;
  let tableData = [];

  const flushTable = () => {
    if (inTable && tableData.length > 0) {
      let table = body.appendTable(tableData);
      table.setBorderWidth(0);
      
      let hasTopNote = false;
      let hasBottomNote = false;
      let hasLeftNote = false;
      let hasRightNote = false;
      let imgRowsCount = 0;
      let maxImgColsCount = 0;
  
      for (let r = 0; r < table.getNumRows(); r++) {
        let row = table.getRow(r);
        let rowCols = row.getNumCells();
        let rText = row.getCell(0).getText().trim();
        if (rText === '補足上側') { hasTopNote = true; continue; }
        if (rText === '補足下側') { hasBottomNote = true; continue; }
    
        imgRowsCount++;
        let rRightText = rowCols > 1 ? row.getCell(rowCols - 1).getText().trim() : '';
        if (rText === '補足左側') hasLeftNote = true;
        if (rRightText === '補足右側') hasRightNote = true;
    
        let cImg = 0;
        for (let c = 0; c < rowCols; c++) {
          if (row.getCell(c).getText().includes('【画像挿入位置:')) cImg++;
        }
        if (cImg > maxImgColsCount) maxImgColsCount = cImg;
      }
      let topCount = (hasTopNote ? 1 : 0) + (hasBottomNote ? 1 : 0);
      let sideCount = (hasLeftNote ? 1 : 0) + (hasRightNote ? 1 : 0);
  
      let wCm = 16.499; let hCm = 24.01;
      if (imgRowsCount === 1 && maxImgColsCount === 1) {
        if (sideCount === 0 && topCount === 0) { wCm = 16.499; hCm = 24.01; }
        else if (sideCount === 1 && topCount === 0) { wCm = 15.499; hCm = 24.11; }
        else if (sideCount === 2 && topCount === 0) { wCm = 14.499; hCm = 24.11; }
        else if (sideCount === 0 && topCount === 1) { wCm = 16.499; hCm = 23.32; }
        else if (sideCount === 0 && topCount === 2) { wCm = 16.499; hCm = 22.76; }
        else if (sideCount === 1 && topCount === 1) { wCm = 15.499; hCm = 23.19; }
        else if (sideCount === 2 && topCount === 1) { wCm = 14.499; hCm = 23.32; }
        else if (sideCount === 1 && topCount === 2) { wCm = 15.499; hCm = 22.52; }
        else if (sideCount === 2 && topCount === 2) { wCm = 14.499; hCm = 22.76; }
      } else if (imgRowsCount === 2 && maxImgColsCount === 1) {
        if (sideCount === 0) { wCm = 16.499; hCm = 11.91; }
        else if (sideCount === 1) { wCm = 15.499; hCm = 11.86; }
        else if (sideCount === 2) { wCm = 14.499; hCm = 11.91; }
      } else if (imgRowsCount === 1 && maxImgColsCount === 2) {
        wCm = 8.25;
        if (topCount === 0) hCm = 24.01;
        else if (topCount === 1) hCm = 23.32;
        else if (topCount === 2) hCm = 22.52;
      } else if (imgRowsCount === 2 && maxImgColsCount === 2) {
        wCm = 8.25; hCm = 11.86;
      } else if (imgRowsCount === 1 && maxImgColsCount === 3) {
        wCm = 5.5;
        if (topCount === 0) hCm = 24.11;
        else if (topCount === 1) hCm = 23.32;
        else if (topCount === 2) hCm = 22.76;
      }
  
      let maxImgWidthPt = wCm * 28.3465; let maxImgHeightPt = hCm * 28.3465; let totalWidthPt = 467.7;
  
      for (let r = 0; r < table.getNumRows(); r++) {
        let row = table.getRow(r);
        let rowCols = row.getNumCells();
        let rText = row.getCell(0).getText().trim();
        let isMergeRow = (rText === '補足上側' || rText === '補足下側');
        for (let c = 0; c < rowCols; c++) {
          let cell = row.getCell(c);
          let cellTextRaw = cell.getText().trim();
          let isSideNote = (cellTextRaw === '補足左側' || cellTextRaw === '補足右側');
          let cWidth = maxImgWidthPt;
      
          if (isMergeRow) {
            if (c === 0) cWidth = totalWidthPt;
            else cWidth = 0.1;
          } else {
            if (isSideNote) cWidth = 28.35;
            else cWidth = maxImgWidthPt;
          }
          cell.setWidth(cWidth);
          cell.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
      
          let cellText = cell.editAsText();
          cellText.setFontFamily(baseFont).setFontSize(10.5);
          let pCount = cell.getNumChildren();
          for(let k=0; k<pCount; k++){
            let cp = cell.getChild(k).asParagraph();
            cp.setLineSpacing(1.0);
            cp.setSpacingBefore(0).setSpacingAfter(0);
      
            let txt = cp.getText().trim();
            if (txt === '補足左側') {
              cp.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
              cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
              cp.editAsText().setForegroundColor('#333333');
            } else if (txt === '補足右側') {
              cp.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
              cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
              cp.editAsText().setForegroundColor('#333333');
            }
      
            const imgMatch = cp.getText().match(/【画像挿入位置:(\d+)】/);
            if (imgMatch) {
              cp.clear();
              const imgIndex = parseInt(imgMatch[1], 10);
              if (images && images.length > imgIndex) {
                const imgObj = images[imgIndex];
                try {
                  let blob;
                  if (imgObj.fileId) {
                    try { blob = DriveApp.getFileById(imgObj.fileId).getBlob(); }
                    catch (be) { blob = Utilities.newBlob(Utilities.base64Decode(imgObj.base64), 'image/jpeg', imgObj.fileName); }
                  } else { blob = Utilities.newBlob(Utilities.base64Decode(imgObj.base64), 'image/jpeg', imgObj.fileName); }
                  cp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                  cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
                
                  const inlineImg = cp.appendInlineImage(blob);
                  const DPI_SCALE = 96 / 72;
                  let targetImgW = maxImgWidthPt * DPI_SCALE;
                  let ratio = targetImgW / inlineImg.getWidth();
                  let targetImgH = inlineImg.getHeight() * ratio;
                  let limitImgH = maxImgHeightPt * DPI_SCALE;
                
                  if (targetImgH > limitImgH) {
                    targetImgH = limitImgH;
                    targetImgW = inlineImg.getWidth() * (targetImgH / inlineImg.getHeight());
                    cell.setWidth(Math.floor(targetImgW / DPI_SCALE));
                  }
                
                  inlineImg.setWidth(Math.floor(targetImgW));
                  inlineImg.setHeight(Math.floor(targetImgH));
                } catch(e) {
                  cp.appendText('[画像挿入エラー]');
                }
              }
            }
          }
        }
      }
      tableData = [];
    }
    inTable = false;
  };

  const toZenkaku = (str) => {
    return str.replace(/[!-\~]/g, function(s) {
      if (s === '(') return '（';
      if (s === ')') return '）';
      if (s === '（' || s === '）') return s;
      return String.fromCharCode(s.charCodeAt(0) + 0xFEE0);
    }).replace(/[ \x20\u0020\xA0\u200B\t]/g, '\u3000');
  };

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    line = line.replace(/\\?\[?[CcＣｃ][IiＩｉ][TtＴｔ][EeＥｅ][:：][\s\u3000]*[０-９0-9,，\s\u3000]*[\]］]?/g, '');
    
    if (i === 0 && line.includes('号証')) {
      continue;
    }
    
    if (line === '【改ページ】') {
      flushTable();
      let pb = body.appendPageBreak();
      pb.getParent().asParagraph().setFontSize(1);
      continue;
    }
    
    if (line.startsWith('|') && line.endsWith('|')) {
      inTable = true;
      if (!line.match(/^\|[-: | \u3000]+\|$/)) {
        let rowData = line.split('|').slice(1, -1).map(cell => cell.trim());
        if (rowData.some(cell => cell !== '')) {
          tableData.push(rowData);
        }
      }
      continue;
    } else if (inTable) {
      flushTable();
    }
    
    if (line === '') continue;
    
    line = toZenkaku(line);
    
    let p;
    let numChildren = body.getNumChildren();
    let lastChild = numChildren > 0 ? body.getChild(numChildren - 1) : null;
    // 表の直後などに自動生成される空の段落を再利用して不要な改行（空行）を防ぐ
    if (lastChild && lastChild.getType() === DocumentApp.ElementType.PARAGRAPH && lastChild.asParagraph().getText() === '') {
      p = lastChild.asParagraph();
      p.setText(line);
    } else {
      p = body.appendParagraph(line);
    }
    
    p.setFontFamily(baseFont).setFontSize(baseSize);
    p.setLineSpacing(1.15);
    
    if (line === '補足上側' || line === '補足下側') {
      p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      p.editAsText().setForegroundColor('#333333').setFontSize(10.5);
      p.setIndentFirstLine(0).setIndentStart(0);
      p.setSpacingBefore(0).setSpacingAfter(0); 
    }
    
    p.setIndentEnd(-6.8);
    p.setBold(false);
  }
  flushTable();
} 

function savePreviewToDraft(caseName, title, content, images, clearPrefix) {
  const tree = getCaseFolderTree(caseName);
  if (clearPrefix) clearEvidenceImagesByPrefix(tree.evidence, clearPrefix);
  // 先にすべての画像を保存
  if (images && images.length > 0) {
    images.forEach(imgObj => {
      try {
        const existings = tree.evidence.getFilesByName(imgObj.fileName);
        while(existings.hasNext()) existings.next().setTrashed(true);
      
        if (imgObj.fileId) {
          try {
            const file = DriveApp.getFileById(imgObj.fileId);
            const copiedFile = file.makeCopy(imgObj.fileName, tree.evidence);
            copiedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch(e) {
            if (imgObj.base64) {
              const blob = Utilities.newBlob(Utilities.base64Decode(imgObj.base64), 'image/jpeg', imgObj.fileName);
              const savedFile = tree.evidence.createFile(blob);
              savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            }
          }
        } else if (imgObj.base64) {
          const blob = Utilities.newBlob(Utilities.base64Decode(imgObj.base64), 'image/jpeg', imgObj.fileName);
          const savedFile = tree.evidence.createFile(blob);
          savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        }
      } catch(e) {}
    });
  }
  // JS側で【要検証_下書き】を付けた場合はそれを尊重し、そうでない場合は【下書き】を付ける
  const docName = title.includes('【要検証_下書き】') ? title : `【下書き】${title}`;
  const existingFiles = tree.draft.getFilesByName(docName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }
  const newFile = DriveApp.getFileById(TEMPLATE_DOC_ID).makeCopy(docName, tree.draft);
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();
  applyLegalStylesToBody(doc, body, content, title, images, tree.evidence, caseName);
  doc.saveAndClose();
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  appendDraftHistory(caseName, title, content);
  return { success: true, url: doc.getUrl(), name: docName };
}

function savePreviewToFinal(caseName, title, content, images, clearPrefix) {
  const tree = getCaseFolderTree(caseName);
  if (clearPrefix) clearEvidenceImagesByPrefix(tree.evidence, clearPrefix);
  // 先にすべての画像を保存（tempからの移動＆リネーム、または新規作成）
  if (images && images.length > 0) {
    images.forEach(imgObj => {
      try {
        const existings = tree.evidence.getFilesByName(imgObj.fileName);
        while(existings.hasNext()) existings.next().setTrashed(true);
    
        if (imgObj.fileId) {
          try {
            const file = DriveApp.getFileById(imgObj.fileId);
            const copiedFile = file.makeCopy(imgObj.fileName, tree.evidence);
            copiedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch(e) {
            if (imgObj.base64) {
              const blob = Utilities.newBlob(Utilities.base64Decode(imgObj.base64), 'image/jpeg', imgObj.fileName);
              const savedFile = tree.evidence.createFile(blob);
              savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            }
          }
        } else if (imgObj.base64) {
          const blob = Utilities.newBlob(Utilities.base64Decode(imgObj.base64), 'image/jpeg', imgObj.fileName);
          const savedFile = tree.evidence.createFile(blob);
          savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        }
      } catch(e) {}
    });
  }
  const tempFile = DriveApp.getFileById(TEMPLATE_DOC_ID).makeCopy(title);
  const tempDoc = DocumentApp.openById(tempFile.getId());
  const body = tempDoc.getBody();
  applyLegalStylesToBody(tempDoc, body, content, title, images, tree.evidence, caseName);
  tempDoc.saveAndClose();
  tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Utilities.sleep(4000);
  const PYTHON_API_URL = "https://pdf-converter-api-215686008543.asia-northeast1.run.app/generate-pdf";
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ url: tempFile.getUrl() }),
    muteHttpExceptions: true
  };
  try {
    const response = fetchWithRetry(PYTHON_API_URL, options);
    if (response.getResponseCode() !== 200) throw new Error(response.getContentText());
    
    const pdfName = `【清書】${title}.pdf`;
    const existingFiles = tree.final.getFilesByName(pdfName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    const pdfBlob = response.getBlob().setName(pdfName);
    const finalFile = tree.final.createFile(pdfBlob);
    finalFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    tempFile.setTrashed(true);
    return { success: true, url: finalFile.getUrl(), name: finalFile.getName() };
  } catch (e) {
    tempFile.setTrashed(true);
    return { success: false, error: e.toString() };
  }
}

function modifyEvidenceFiles(caseName, sign, num, actionType) {
  const tree = getCaseFolderTree(caseName);
  const numHankaku = num.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
  const numZenkaku = num.replace(/[0-9]/g, s => String.fromCharCode(s.charCodeAt(0) + 0xFEE0));
  
  [tree.draft, tree.final].forEach(folder => {
    const files = folder.getFiles();
    while (files.hasNext()) {
      let file = files.next();
      let name = file.getName();
      if (name.includes(`${sign}第${numHankaku}号証`) || name.includes(`${sign}第${numZenkaku}号証`)) {
        file.setTrashed(true);
      }
    }
  });

  const prefix = `${sign}${numHankaku}_`;
  const evidenceFiles = tree.evidence.getFiles();
  let newImageInfos = [];
  
  while (evidenceFiles.hasNext()) {
    let file = evidenceFiles.next();
    let name = file.getName();
    if (name.startsWith(prefix)) {
      if (actionType === 'delete') {
        file.setTrashed(true);
      } else if (actionType === 'edit') {
        let newName = `temp_${name}`;
        file.setName(newName);
        newImageInfos.push({ id: file.getId(), name: newName });
      }
    }
  }
  return { success: true, action: actionType, images: newImageInfos };
}

// ==========================================
// 【UIの3列モニター更新用エンジン】
// ==========================================
function getFolderMonitorStatus(caseName) {
  if (!caseName) return { draft: [], final: [], evidence: [] };
  const tree = getCaseFolderTree(caseName);
  if (!tree) return { draft: [], final: [], evidence: [] };
  const getFiles = (folder) => {
    let files = [];
    const iterator = folder.getFiles();
    while (iterator.hasNext()) {
      const file = iterator.next();
      files.push({ name: file.getName(), url: file.getUrl(), id: file.getId() });
    }
    return files.sort((a, b) => a.name.localeCompare(b.name));
  };
  return { draft: getFiles(tree.draft), final: getFiles(tree.final), evidence: getFiles(tree.evidence) };
}

// ==========================================
// 【セーブデータ管理・リネーム機能】
// ==========================================
function saveCaseSettings(caseName, settingsJson) {
  const oldJsonStr = loadCaseSettings(caseName);
  let oldData = null;
  if (oldJsonStr) {
    try { oldData = JSON.parse(oldJsonStr); } catch(e) {}
  }

  const folder = getOrCreateCaseFolder(caseName);
  const fileName = '【設定データ】' + caseName;
  const lock = LockService.getScriptLock();
  let ss;
  try {
    lock.waitLock(10000); // 同時作成を防ぐ排他制御
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(fileName);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
    }
  } catch (e) {
    // ロック取得エラー時のフォールバック
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(fileName);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
    }
  } finally {
    lock.releaseLock();
  }
  const sheet = ss.getSheets()[0];
  sheet.clear();
  sheet.getRange(1, 1).setValue(settingsJson);

  try {
    syncRelatedCases(JSON.parse(settingsJson), oldData);
  } catch (e) {
    console.warn("相互自動追加エラー:", e);
  }

  return true;
}

function loadCaseSettings(caseName) {
  const folder = getOrCreateCaseFolder(caseName);
  const files = folder.getFilesByName('【設定データ】' + caseName);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next()).getSheets()[0].getRange(1, 1).getValue();
  }
  return null;
}

function renameCaseFolder(oldName, newName) {
  if (oldName === newName) return { success: true };
  const masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  const folders = masterFolder.getFoldersByName(oldName);
  if (folders.hasNext()) {
    const folder = folders.next();
    folder.setName(newName); 
    const files = folder.getFilesByName('【設定データ】' + oldName);
    if (files.hasNext()) {
      files.next().setName('【設定データ】' + newName);
    }
    return { success: true };
  }
  return { success: false, error: 'フォルダが見つかりません' };
}

function getFolderUrls(folderNames) {
  const masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  const urlMap = {};
  folderNames.forEach(name => {
    if (!name) return;
    let folders = masterFolder.getFoldersByName(name);
    if (folders.hasNext()) {
      urlMap[name] = folders.next().getUrl();
    } else {
      // 揺らぎ検索
      let altName = name;
      if (name.match(/(養育費|婚姻費用分担|財産分与|慰謝料)調停事件/)) {
        altName = name.replace(/(養育費|婚姻費用分担|財産分与|慰謝料)調停事件/g, '$1請求調停事件');
      } else if (name.match(/(養育費|婚姻費用分担|財産分与|慰謝料)審判事件/)) {
        altName = name.replace(/(養育費|婚姻費用分担|財産分与|慰謝料)審判事件/g, '$1請求審判事件');
      } else if (name.match(/(養育費|婚姻費用分担|財産分与|慰謝料)請求調停事件/)) {
        altName = name.replace(/(養育費|婚姻費用分担|財産分与|慰謝料)請求調停事件/g, '$1調停事件');
      } else if (name.match(/(養育費|婚姻費用分担|財産分与|慰謝料)請求審判事件/)) {
        altName = name.replace(/(養育費|婚姻費用分担|財産分与|慰謝料)請求審判事件/g, '$1審判事件');
      }
      
      if (altName !== name) {
        let altFolders = masterFolder.getFoldersByName(altName);
        if (altFolders.hasNext()) {
          urlMap[name] = altFolders.next().getUrl();
          return;
        }
      }
      urlMap[name] = null;
    }
  });
  return urlMap;
}

// ==========================================
// 【経緯・進捗のAI解析機能】
// ==========================================
function performAiAnalysisForCaseHistory(historyText, varsJson, listBodyText) {
  const prompt = `あなたは優秀な日本の弁護士・法務スタッフです。
  以下の【経緯・進捗テキスト】を解析し、【現在の変数】および【証拠一覧（マークダウン表）】の空欄部分を推測して補完し、内容に矛盾があれば指摘してください。

  【ルール】
  - 該当データのみを持つJSON形式で出力してください。（Markdownコードブロック不要）
  - キーは "補完変数", "証拠補完", "本文補完", "指摘事項" としてください。
  - "補完変数"には、【現在の変数】のうち空欄になっている項目、あるいはテキストから明らかに修正すべきと判断される項目（当事者名や裁判所名など）のみを抽出してください。
  - "証拠補完"には、【証拠一覧】の表にある各証拠（号証）について、テキストから推測される「立証趣旨」や「作成日」「作成者」が判明した場合に、その号証と補完内容を含めてください。（例：{ "号証": "甲1", "立証趣旨": "〇〇を立証するため" }）
  - "本文補完"には、テキスト内の「【証拠未定】」というプレースホルダーや、書証の引用漏れと思われる箇所について、推測される号証（例：甲1）に置き換えた該当文の修正案を文字列の配列で提示してください。
  - "指摘事項"には、日付の矛盾や当事者の取り違えなど、論理的な間違い探しを行って文字列の配列で提示してください。何もなければ空配列としてください。

  【経緯・進捗テキスト】
  ${historyText}

  【現在の変数】
  ${varsJson}

  【証拠一覧】
  ${listBodyText}`;

  const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
  try {
    const response = callGeminiApiWithFallback(payload, MODEL_PRO);
    let jsonStr = response.getContentText();
    jsonStr = JSON.parse(jsonStr).candidates[0].content.parts[0].text;
    jsonStr = jsonStr.replace(/```(?:json)?\n?([\s\S]*?)\n?```/g, '$1').trim();
    return jsonStr;
  } catch (e) {
    throw new Error(e.toString());
  }
}

// ==========================================
// 【履歴管理機能】
// ==========================================
function appendDraftHistory(caseName, title, content) {
  const folder = getOrCreateCaseFolder(caseName);
  const fileName = '【設定データ】' + caseName;
  const files = folder.getFilesByName(fileName);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
  }
  
  let sheet = ss.getSheetByName('DraftHistory');
  if (!sheet) {
    sheet = ss.insertSheet('DraftHistory');
    sheet.appendRow(['保存日時', 'タイトル', '本文']);
    sheet.setFrozenRows(1);
  }
  
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  sheet.appendRow([nowStr, title, content]);
}

function getDraftHistory(caseName) {
  const folder = getOrCreateCaseFolder(caseName);
  
  // ドライブ上の下書きフォルダに存在するファイル名を取得
  const draftsFolder = folder.getFoldersByName('01_下書き');
  const validTitles = new Set();
  if (draftsFolder.hasNext()) {
    const files = draftsFolder.next().getFiles();
    while (files.hasNext()) {
      let fName = files.next().getName();
      if (fName.startsWith('【下書き】')) {
        fName = fName.replace('【下書き】', '').replace('.pdf', '');
      }
      validTitles.add(fName);
    }
  }

  const files = folder.getFilesByName('【設定データ】' + caseName);
  if (!files.hasNext()) return [];
  const ss = SpreadsheetApp.open(files.next());
  const sheet = ss.getSheetByName('DraftHistory');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const history = [];
  const rowsToDelete = [];
  
  for (let i = 1; i < data.length; i++) {
    const title = data[i][1];
    // ドライブ上に実在するファイルのみ履歴を読み込む
    if (validTitles.has(title)) {
      let dateVal = data[i][0];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
      }
      history.push({
        date: dateVal,
        title: title,
        content: data[i][2]
      });
    } else {
      // 存在しないファイルの履歴行は削除対象
      rowsToDelete.push(i + 1);
    }
  }
  
  // ゴミ箱に入れた等で存在しなくなったファイルの履歴を一括削除（行ずれ防止のため下から）
  if (rowsToDelete.length > 0) {
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
  }
  return history;
}

function deleteSpecificDraftHistory(caseName, dateStr, title) {
  const folder = getOrCreateCaseFolder(caseName);
  const files = folder.getFilesByName('【設定データ】' + caseName);
  if (!files.hasNext()) return { success: false };
  const ss = SpreadsheetApp.open(files.next());
  const sheet = ss.getSheetByName('DraftHistory');
  if (!sheet) return { success: false };
  
  const data = sheet.getDataRange().getValues();
  // 下から検索して該当する日時の版を削除
  for (let i = data.length - 1; i >= 1; i--) {
    let rowDate = data[i][0];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    }
    if (rowDate === dateStr && data[i][1] === title) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

function renameDriveFile(fileId, newName) {
  try {
    const file = DriveApp.getFileById(fileId);
    file.setName(newName);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function convertDraftToPdf(caseName, docId, docName) {
  const tree = getCaseFolderTree(caseName);
  const docFile = DriveApp.getFileById(docId);
  const originalAccess = docFile.getSharingAccess();
  const originalPermission = docFile.getSharingPermission();
  // 一時的に閲覧許可
  docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Utilities.sleep(4000);
  const PYTHON_API_URL = "https://pdf-converter-api-215686008543.asia-northeast1.run.app/generate-pdf";
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ url: docFile.getUrl() }),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(PYTHON_API_URL, options);
    docFile.setSharing(originalAccess, originalPermission); // 権限を戻す
    if (response.getResponseCode() !== 200) throw new Error(response.getContentText());
    const pdfName = docName.replace('【下書き】', '【清書】') + '.pdf';

    // 同名ファイルの削除（上書きの代用）
    const existingFiles = tree.final.getFilesByName(pdfName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    const pdfBlob = response.getBlob().setName(pdfName);
    const finalFile = tree.final.createFile(pdfBlob);
    // リンクを知っている全員が閲覧可能に設定（アカウント選択回避）
    finalFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: finalFile.getUrl() };
  } catch(e) {
    docFile.setSharing(originalAccess, originalPermission);
    return { success: false, error: e.toString() };
  }
}

function getPdfBase64ForShare(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  return {
    name: file.getName(),
    mimeType: blob.getContentType(),
    base64: Utilities.base64Encode(blob.getBytes())
  };
}

// ==========================================
// 【Google Picker API & ドライブファイル取得】
// ==========================================
function getPickerConfig() {
  return {
    token: ScriptApp.getOAuthToken(),
    apiKey: PropertiesService.getScriptProperties().getProperty('PLACES_API_KEY')
  };
}

function getDriveFileBase64(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    let blob;
    // ドキュメントやスプレッドシートは直接Base64化できないためPDFに変換
    if (mimeType === MimeType.GOOGLE_DOCS || mimeType === MimeType.GOOGLE_SHEETS || mimeType === MimeType.GOOGLE_SLIDES) {
      blob = file.getAs(MimeType.PDF);
    } else {
      blob = file.getBlob();
    }
    return {
      name: file.getName(),
      mimeType: blob.getContentType(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (e) {
    throw new Error('Driveファイルの取得に失敗しました: ' + e.message);
  }
}

function getDriveFileText(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    if (mimeType === MimeType.GOOGLE_DOCS) {
      return DocumentApp.openById(fileId).getBody().getText();
    } else {
      return file.getBlob().getDataAsString();
    }
  } catch (e) {
    throw new Error('ファイルの読み込みに失敗しました: ' + e.message);
  }
}

function deleteDriveFiles(fileIds) {
  let count = 0;
  fileIds.forEach(id => {
    try {
      DriveApp.getFileById(id).setTrashed(true);
      count++;
    } catch(e) {}
  });
  return count;
}

function clearEvidenceImagesByPrefix(evidenceFolder, prefix) {
  if (!prefix) return;
  const files = evidenceFolder.getFiles();
  while (files.hasNext()) {
    let file = files.next();
    if (file.getName().startsWith(prefix)) {
      file.setTrashed(true);
    }
  }
}

function checkExistingEvidenceImages(caseName, prefix) {
  if (!caseName || !prefix) return 0;
  const tree = getCaseFolderTree(caseName);
  const files = tree.evidence.getFiles();
  let count = 0;
  while (files.hasNext()) {
    if (files.next().getName().startsWith(prefix)) count++;
  }
  return count;
}

// ==========================================
// 【自動同期用の最新ロック情報取得】
// ==========================================
function getLatestCaseLock(caseName) {
  const jsonStr = loadCaseSettings(caseName);
  if (jsonStr) {
    try {
      const data = JSON.parse(jsonStr);
      return data.deviceLock || null;
    } catch(e) { return null; }
  }
  return null;
}

// ==========================================
// 【関連事件の相互自動追加・削除処理】
// ==========================================
function syncRelatedCases(mainData, oldData) {
  if (!mainData || !mainData.main) return;
    const getFullNameLocal = (obj) => {
    if (!obj) return null;
    let era = obj.era || "令和";
    let numStr = obj.year ? `${era}${obj.year}年(${obj.symbol})第${obj.num}号 ` : "";
    let name = obj.name || "";
    let fullName = `${numStr}${name}`;
    return fullName.replace(/(養育費|婚姻費用分担|財産分与|慰謝料)(調停事件|審判事件)/g, '$1請求$2');
  };
  const mainFullName = getFullNameLocal(mainData.main);
  const masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const currentSubCases = mainData.subCases || [];
  const oldSubCases = (oldData && oldData.subCases) ? oldData.subCases : [];
  const currentSubNames = currentSubCases.map(sub => getFullNameLocal(sub));
  const removedSubCases = oldSubCases.filter(sub => !currentSubNames.includes(getFullNameLocal(sub)));
  // 削除処理
  removedSubCases.forEach(removedSub => {
    let targetFullName = getFullNameLocal(removedSub);
    let folders = masterFolder.getFoldersByName(targetFullName);
    if (!folders.hasNext()) {
        let altSub = JSON.parse(JSON.stringify(removedSub));
        altSub.name = altSub.name.replace(/\(第\d+事件\)/g, '');
        targetFullName = getFullNameLocal(altSub);
        folders = masterFolder.getFoldersByName(targetFullName);
    }
    if (!folders.hasNext()) return;
    
    const folder = folders.next();
    let actualFolderName = folder.getName();
    let files = folder.getFilesByName('【設定データ】' + actualFolderName);
    if (!files.hasNext()) {
        files = folder.getFilesByName('【設定データ】' + targetFullName);
    }
    if (!files.hasNext()) return;
    
    const ss = SpreadsheetApp.open(files.next());
    const sheet = ss.getSheets()[0];
    const jsonStr = sheet.getRange(1, 1).getValue();
    if (!jsonStr) return;
    let targetData;
    try { targetData = JSON.parse(jsonStr); } catch(e) { return; }
    if (targetData.subCases) {
      const initialLength = targetData.subCases.length;
      
      const oldMainEra = oldData.main ? oldData.main.era : mainData.main.era;
      const oldMainYear = oldData.main ? oldData.main.year : mainData.main.year;
      const oldMainSymbol = oldData.main ? oldData.main.symbol : mainData.main.symbol;
      const oldMainNum = oldData.main ? oldData.main.num : mainData.main.num;
      targetData.subCases = targetData.subCases.filter(s => {
          // 型の違い（文字列と数値）を吸収するため「==」で比較
          if (s.era == oldMainEra && s.year == oldMainYear && s.symbol == oldMainSymbol && s.num == oldMainNum) return false;
          if (s.era == mainData.main.era && s.year == mainData.main.year && s.symbol == mainData.main.symbol && s.num == mainData.main.num) return false;
          return true;
      });
      if (targetData.subCases.length !== initialLength) {
        let hasMerge = targetData.subCases.some(s => s.type === '併合');
        if (!hasMerge && targetData.main.name.match(/\(第\d+事件\)/)) {
            targetData.main.name = targetData.main.name.replace(/\(第\d+事件\)/g, '');
            const newBFullName = getFullNameLocal(targetData.main);
            if (actualFolderName !== newBFullName) {
              folder.setName(newBFullName);
              const oldFiles = folder.getFilesByName('【設定データ】' + actualFolderName);
              if (oldFiles.hasNext()) oldFiles.next().setName('【設定データ】' + newBFullName);
            }
        }
        sheet.getRange(1, 1).setValue(JSON.stringify(targetData));
      }
    }
  });
  // 追加・更新処理
  if (currentSubCases.length === 0) return;
  const group = [
    { type: currentSubCases[0].type, obj: mainData.main, fullName: mainFullName }
  ];
  currentSubCases.forEach(sub => {
    group.push({ type: sub.type, obj: sub, fullName: getFullNameLocal(sub) });
  });
  currentSubCases.forEach(targetSub => {
    let targetFullName = getFullNameLocal(targetSub);
    let folders = masterFolder.getFoldersByName(targetFullName);
    if (!folders.hasNext()) {
        let altSub = JSON.parse(JSON.stringify(targetSub));
        altSub.name = altSub.name.replace(/\(第\d+事件\)/g, '');
        folders = masterFolder.getFoldersByName(getFullNameLocal(altSub));
    }
    if (!folders.hasNext()) return;
    const folder = folders.next();
    let actualFolderName = folder.getName();
    let files = folder.getFilesByName('【設定データ】' + actualFolderName);
    if (!files.hasNext()) return;
    
    const ss = SpreadsheetApp.open(files.next());
    const sheet = ss.getSheets()[0];
    const jsonStr = sheet.getRange(1, 1).getValue();
    if (!jsonStr) return;
    
    let targetData;
    try { targetData = JSON.parse(jsonStr); } catch(e) { return; }
    if (!targetData.subCases) targetData.subCases = [];
    let isChanged = false;
    group.forEach(member => {
      let isSelf = false;
      // 型の違いを吸収するため「==」で比較
      if (member.obj.era == targetData.main.era && member.obj.year == targetData.main.year &&
          member.obj.symbol == targetData.main.symbol && member.obj.num == targetData.main.num) {
          isSelf = true;
      }
      if (isSelf) return;
      
      const existsIdx = targetData.subCases.findIndex(s =>
          (s.era == member.obj.era && s.year == member.obj.year && s.symbol == member.obj.symbol && s.num == member.obj.num)
      );
      if (existsIdx === -1) {
        let newSub = {
          type: member.type,
          era: member.obj.era,
          year: member.obj.year,
          symbol: member.obj.symbol,
          num: member.obj.num,
          name: member.obj.name,
          isPast: false
        };
        targetData.subCases.push(newSub);
        isChanged = true;
      } else {
        const existingSub = targetData.subCases[existsIdx];
        if (existingSub.name !== member.obj.name || existingSub.type !== member.type) {
            existingSub.name = member.obj.name;
            existingSub.type = member.type;
            isChanged = true;
        }
      }
    });
    if (targetSub.type === '併合' && targetData.main.name !== targetSub.name) {
        targetData.main.name = targetSub.name;
        isChanged = true;
        const newFullName = getFullNameLocal(targetData.main);
        if (actualFolderName !== newFullName) {
          folder.setName(newFullName);
          const oldFiles = folder.getFilesByName('【設定データ】' + actualFolderName);
          if (oldFiles.hasNext()) oldFiles.next().setName('【設定データ】' + newFullName);
        }
    }
    if (isChanged) {
      sheet.getRange(1, 1).setValue(JSON.stringify(targetData));
    }
  });
}