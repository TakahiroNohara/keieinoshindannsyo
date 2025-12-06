/**
 * 会計診断システム GASコード (v11.0 + OCR統合版)
 * 
 * 構成:
 * 1. システム設定・メニュー (v11.0ベース)
 * 2. メイン処理: データ読込・集計 (v11.0ベース)
 * 3. OCR機能: PDF読取・データ整形 (旧コードより移植・改修)
 */

// =================================================================================
// 1. システム設定・定数
// =================================================================================

const SHEETS = {
  CONFIG: '設定',
  INPUT: 'OCR貼付',
  WORK: '仕訳作業',
  DB: '学習データ',
  OUTPUT: '計算用データ'
};

const FIXED_ATTRIBUTES = [
  "営業外収益", "営業外費用", "特別利益", "特別損失",
  "BS_有形固定資産", "BS_無形固定資産", "BS_投資その他"
];

// API KEYのキャッシュ
let CACHED_API_KEY = null;

function getApiKey() {
  if (CACHED_API_KEY === null) {
    CACHED_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!CACHED_API_KEY) {
      throw new Error('GEMINI_API_KEYが設定されていません。スクリプトプロパティで設定してください。');
    }
  }
  return CACHED_API_KEY;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('★会計システム')
    .addItem('0. PDFからOCR取込 (Gemini)', 'startStep1_showFileDialog')
    .addSeparator()
    .addItem('1. OCRデータ読込・初期分類', 'processOCRData')
    .addSeparator()
    .addItem('2. 確定・集計実行', 'finalizeAndAggregate')
    .addToUi();
}

// =================================================================================
// 2. メイン処理 (v11.0)
// =================================================================================

function processOCRData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(SHEETS.INPUT);
  const workSheet = ss.getSheetByName(SHEETS.WORK);
  const dbSheet = ss.getSheetByName(SHEETS.DB);

  if (!inputSheet || !workSheet || !dbSheet) {
    SpreadsheetApp.getUi().alert('必要なシート（OCR貼付, 仕訳作業, 学習データ）が見つかりません。');
    return;
  }
  
  const inputData = inputSheet.getDataRange().getValues();
  const dbData = dbSheet.getDataRange().getValues();
  
  if (inputData.length < 2) {
    SpreadsheetApp.getUi().alert('「OCR貼付」シートにデータがありません。先にOCR取込を行ってください。');
    return;
  }

  // DBをメモリに展開
  const dbMap = new Map();
  for (let i = 1; i < dbData.length; i++) {
    const row = dbData[i];
    if (!row) continue;
    if (row[0]) {
      dbMap.set(String(row[0]), { name: row[1], attr: row[2] });
    }
  }
  
  let output = [];
  let skippedCount = 0;
  
  // OCRデータを走査
  for (let i = 1; i < inputData.length; i++) {
    const row = inputData[i];
    if (!row) continue; 
    
    const rawName = (row.length > 0 && row[0]) ? String(row[0]) : "";
    const ocrName = rawName.trim();
    
    let amount = 0;
    let hasAmount = false;
    if (row.length > 1 && row[1] !== "") {
      amount = Number(row[1]);
      if (!isNaN(amount)) {
        hasAmount = true;
      } else {
        amount = 0;
      }
    }
    const period = (row.length > 2 && row[2]) ? String(row[2]).trim() : "";
    
    // 空行スキップ
    if (!ocrName && !hasAmount) {
      skippedCount++;
      continue;
    }
    
    let standardName = ocrName;
    if (!standardName) {
        standardName = "【名称不明】";
    }
    
    let attribute = "";
    let status = "★新規";
    
    // 学習データとの照合
    if (ocrName && dbMap.has(ocrName)) {
      const known = dbMap.get(ocrName);
      standardName = known.name;
      attribute = known.attr;
      status = "学習済";
    }
    
    // 出力行: [元名称, 金額, 標準名称, 属性, (確認用)標準名称, (確認用)属性, ステータス, 対象期]
    // ※v11.0コードに基づき、確認用列も初期値としてセット
    const outputRow = [
      String(rawName), amount, String(standardName), String(attribute),
      String(standardName), String(attribute), String(status), String(period)
    ];
    output.push(outputRow);
  }
  
  if (output.length > 0) {
    const lastRow = workSheet.getLastRow();
    if (lastRow > 1) {
      try { workSheet.getRange(2, 1, lastRow - 1, 8).clearContent(); } catch(e) {}
    }
    try {
      workSheet.getRange(2, 1, output.length, 8).setValues(output);
      let msg = `読み込み完了：${output.length}件`;
      if (skippedCount > 0) msg += `\n（${skippedCount}件の完全な空行をスキップ）`;
      SpreadsheetApp.getUi().alert(msg);
    } catch (e) {
      SpreadsheetApp.getUi().alert('書き込みエラー。\n詳細: ' + e.message);
    }
  } else {
    SpreadsheetApp.getUi().alert('有効なデータなし。');
  }
}

function finalizeAndAggregate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const workSheet = ss.getSheetByName(SHEETS.WORK);
    const dbSheet = ss.getSheetByName(SHEETS.DB);
    const configSheet = ss.getSheetByName(SHEETS.CONFIG);
    const outputSheet = ss.getSheetByName(SHEETS.OUTPUT);
    
    if (!workSheet || !dbSheet || !configSheet || !outputSheet) throw new Error("シート不足: 必要なシートが存在しません。");

    const workData = workSheet.getDataRange().getValues();
    if (workData.length < 2) {
      SpreadsheetApp.getUi().alert('「仕訳作業」シートにデータがありません。');
      return;
    }

    // 1. 学習処理（上書き更新）
    const dbData = dbSheet.getDataRange().getValues();
    const dbMap = new Map(); 
    for (let i = 1; i < dbData.length; i++) {
      if (dbData[i] && dbData[i][0]) {
        dbMap.set(String(dbData[i][0]), { name: dbData[i][1], attr: dbData[i][2] });
      }
    }
    
    let hasUpdate = false;
    for (let i = 1; i < workData.length; i++) {
      const row = workData[i];
      if (!row) continue;
      // workSheet列: 0:元名, 1:金額, 2:標準, 3:属性, 4:確認後標準, 5:確認後属性
      const ocrName = (row.length > 0 && row[0]) ? String(row[0]).trim() : "";
      const confirmedName = (row.length > 4) ? row[4] : "";
      const confirmedAttr = (row.length > 5) ? row[5] : "";
      
      if (ocrName && confirmedName !== "【名称不明】" && confirmedName && confirmedAttr) {
        const current = dbMap.get(ocrName);
        // 変更があればDBマップを更新
        if (!current || current.name !== confirmedName || current.attr !== confirmedAttr) {
          dbMap.set(ocrName, { name: confirmedName, attr: confirmedAttr });
          hasUpdate = true;
        }
      }
    }
    
    if (hasUpdate) {
      let newDbData = [];
      dbMap.forEach((val, key) => { newDbData.push([key, val.name, val.attr]); });
      const lastRow = dbSheet.getLastRow();
      if (lastRow > 1) { dbSheet.getRange(2, 1, lastRow - 1, 3).clearContent(); }
      if (newDbData.length > 0) { dbSheet.getRange(2, 1, newDbData.length, 3).setValues(newDbData); }
    }
    
    // 2. 設定読込
    const configMap = new Map();
    const configRaw = configSheet.getDataRange().getValues();
    for (let i = 0; i < configRaw.length; i++) {
      const row = configRaw[i];
      if (!row) continue;
      if(row.length > 1 && row[0] && row[1] !== "" && row[1] !== undefined) {
        configMap.set(String(row[0]), row[1]);
      }
    }
    
    // 3. 集計処理
    const aggregated = {}; // { Period: { Attr: { Name: Amount } } }
    for (let i = 1; i < workData.length; i++) {
      const row = workData[i];
      if (!row) continue;
      let name = (row.length > 4 && row[4]) ? String(row[4]) : ""; 
      const attr = (row.length > 5 && row[5]) ? String(row[5]) : ""; 
      let amount = 0;
      if (row.length > 1 && row[1] !== "") {
          amount = Number(row[1]);
          if (isNaN(amount)) amount = 0;
      }
      const period = (row.length > 7 && row[7]) ? String(row[7]) : ""; 
      
      if (!attr || !name) continue; 
      if (name === "【名称不明】") continue;

      // 固定属性の場合は、名称を属性名に統一して合算
      if (FIXED_ATTRIBUTES.includes(attr)) name = attr;
      
      if (!aggregated[period]) aggregated[period] = {};
      if (!aggregated[period][attr]) aggregated[period][attr] = {};
      if (!aggregated[period][attr][name]) aggregated[period][attr][name] = 0;
      aggregated[period][attr][name] += amount;
    }
    
    // 4. 出力データ生成 (Top N処理)
    let finalOutput = [];
    for (const period in aggregated) {
      for (const attr in aggregated[period]) {
        let items = [];
        for (const name in aggregated[period][attr]) {
          items.push({ name: name, amount: aggregated[period][attr][name] });
        }
        // 金額順（降順? 昇順? v11.0は `b - a` なので降順と思われますが、負債などは逆の可能性も。ここではそのまま踏襲）
        items.sort((a, b) => b.amount - a.amount);
        
        if (FIXED_ATTRIBUTES.includes(attr)) {
            items.forEach(item => finalOutput.push([item.name, attr, item.amount, period]));
            continue; 
        }

        // Top N 制限
        let limit = 9999;
        for (const [key, val] of configMap) {
          // 設定シートのキー例: "変動費_表示数"
          if (key.includes(attr) && String(key).includes("表示数")) {
            limit = Number(val); break;
          }
        }
        
        let exportItems = [];
        if (items.length <= limit) {
          exportItems = items;
        } else {
          const topItems = [];
          const otherItems = [];
          for(let k=0; k < items.length; k++) {
            if (k < limit - 1) topItems.push(items[k]);
            else otherItems.push(items[k]);
          }
          let otherTotal = 0;
          otherItems.forEach(item => otherTotal += item.amount);
          exportItems = topItems;
          if (otherTotal > 0) { exportItems.push({ name: "その他", amount: otherTotal }); }
        }
        
        exportItems.forEach(item => { finalOutput.push([item.name, attr, item.amount, period]); });
      }
    }
    
    // 5. 書き込み
    const outputLastRow = outputSheet.getLastRow();
    if (outputLastRow > 1) { outputSheet.getRange(2, 1, outputLastRow - 1, 5).clearContent(); }
    
    if (finalOutput.length > 0) {
      outputSheet.getRange(2, 1, finalOutput.length, 4).setValues(finalOutput);
      let msg = '集計完了！\n「計算用データ」シートを更新しました。';
      if (hasUpdate) msg += '\n\n★「学習データ」も更新されました。';
      SpreadsheetApp.getUi().alert(msg);
    } else {
      SpreadsheetApp.getUi().alert('集計結果がありません。');
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}

// =================================================================================
// 3. OCR機能 (PDF -> OCR貼付シート)
// =================================================================================

/**
 * ダイアログ表示
 */
function startStep1_showFileDialog() {
  const html = HtmlService.createHtmlOutputFromFile('dialog.html').setWidth(450).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '処理する決算書PDFを選択');
}

/**
 * 1つのファイル（3期比較表）からOCR抽出して「OCR貼付」シートへ
 */
function setFileIdAndExtract(fileId) {
  const file = DriveApp.getFileById(fileId);
  if (!file) {
    throw new Error('ファイルが見つかりませんでした。');
  }
  
  try {
    const resultText = callGeminiApi(file); // 3期分(4カラム)抽出
    const results = parseItemValueCsvResult(resultText); // [[項目, 前々期, 前期, 当期], ...]
    
    if (results.length > 0) {
      // データを縦積みに変換 [項目, 金額, 期]
      const stackedData = [];
      results.forEach(row => {
        const item = row[0];
        if (row[1]) stackedData.push([item, row[1], '前々期']);
        if (row[2]) stackedData.push([item, row[2], '前期']);
        if (row[3]) stackedData.push([item, row[3], '当期']);
      });
      
      writeToInputSheet(stackedData);
      return {
        success: true,
        message: `OCR完了。\n「${SHEETS.INPUT}」シートにデータを貼り付けました。\n\nメニューの「1. OCRデータ読込・初期分類」を実行してください。`
      };
    } else {
      throw new Error('有効なデータを抽出できませんでした。');
    }
  } catch (e) {
    Logger.log(e);
    throw new Error('処理エラー: ' + e.message);
  }
}

/**
 * 3つのファイル（各期）からOCR抽出して「OCR貼付」シートへ
 */
function setMultipleFilesAndExtract(fileCurrentId, filePreviousId, file2PeriodsAgoId) {
  try {
    const fileCurrent = DriveApp.getFileById(fileCurrentId);
    const filePrevious = DriveApp.getFileById(filePreviousId);
    const file2PeriodsAgo = DriveApp.getFileById(file2PeriodsAgoId);

    if (!fileCurrent || !filePrevious || !file2PeriodsAgo) {
      throw new Error('いずれかのファイルが見つかりませんでした。');
    }

    const stackedData = [];

    // 1. 前々期
    const res2Ago = callGeminiApiSinglePeriod(file2PeriodsAgo, '前々期');
    const data2Ago = parseSinglePeriodCsvResult(res2Ago);
    data2Ago.forEach(row => stackedData.push([row[0], row[1], '前々期']));

    // 2. 前期
    const resPrev = callGeminiApiSinglePeriod(filePrevious, '前期');
    const dataPrev = parseSinglePeriodCsvResult(resPrev);
    dataPrev.forEach(row => stackedData.push([row[0], row[1], '前期']));

    // 3. 当期
    const resCurr = callGeminiApiSinglePeriod(fileCurrent, '当期');
    const dataCurr = parseSinglePeriodCsvResult(resCurr);
    dataCurr.forEach(row => stackedData.push([row[0], row[1], '当期']));

    if (stackedData.length > 0) {
      writeToInputSheet(stackedData);
      return {
        success: true,
        message: `3期分のOCR完了。\n「${SHEETS.INPUT}」シートに${stackedData.length}件のデータを貼り付けました。\n\nメニューの「1. OCRデータ読込・初期分類」を実行してください。`
      };
    } else {
      throw new Error('有効なデータを抽出できませんでした。');
    }

  } catch (e) {
    Logger.log(e);
    throw new Error('処理エラー: ' + e.message);
  }
}

/**
 * OCRデータを「OCR貼付」シートに書き込む共通関数
 * @param {Array} data - [[科目, 金額, 期], ...]
 */
function writeToInputSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.INPUT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.INPUT);
    sheet.getRange(1, 1, 1, 3).setValues([['勘定科目', '金額', '対象期']]);
  }
  
  // 既存データをクリア
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }
  
  // 書き込み
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 3).setValues(data);
  }
}

// =================================================================================
// 4. OCR用ヘルパー関数・API呼び出し (旧コードより)
// =================================================================================

function callGeminiApi(file) {
  return callWithRetry(() => {
    const API_KEY = getApiKey();
    const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + API_KEY;
    const prompt = `このファイルには3期分の決算書（貸借対照表・損益計算書）が含まれています。すべての勘定科目と3期分の金額（前々期、前期、当期）を抽出してください。
出力形式: CSV
列順序: 勘定科目, 前々期金額, 前期金額, 当期金額
金額は数値のみ（カンマなし、マイナスは▲または-）。金額がない場合は0。`;
    
    const requestBody = {
      "contents": [{"parts": [{ "text": prompt }, { "inline_data": { "mime_type": file.getMimeType(), "data": Utilities.base64Encode(file.getBlob().getBytes()) } }] }],
      "generationConfig": { "temperature": 0.0, "maxOutputTokens": 8192 }
    };
    const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(requestBody), 'muteHttpExceptions': true };
    const response = UrlFetchApp.fetch(API_ENDPOINT, options);
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      return json.candidates?.[0]?.content?.parts?.[0]?.text || "";
    }
    throw new Error(`API Error: ${response.getResponseCode()}`);
  }, 3, 1000);
}

function callGeminiApiSinglePeriod(file, label) {
  return callWithRetry(() => {
    const API_KEY = getApiKey();
    const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + API_KEY;
    const prompt = `あなたは決算書OCRエンジンです。
提供されたPDF（貸借対照表、損益計算書、製造原価報告書、販管費内訳書）から、勘定科目と金額をすべて抽出してください。

【出力ルール】
1. 出力はCSV形式のみ。「` + "```" + `csv」などのMarkdownタグや、挨拶文、説明は一切不要。
2. 1行に「勘定科目,金額」の形式で出力。
3. 金額は半角数字のみ（カンマ「,」や円マーク「¥」は削除）。マイナスは「-」または「▲」。
4. 貸借対照表、損益計算書などの主要な表の項目はすべて網羅すること。

例:
現金及び預金,12345000
売上高,98765000
...`;

    const requestBody = {
      "contents": [{"parts": [{ "text": prompt }, { "inline_data": { "mime_type": file.getMimeType(), "data": Utilities.base64Encode(file.getBlob().getBytes()) } }] }],
      "generationConfig": { "temperature": 0.0, "maxOutputTokens": 8192 }
    };
    const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(requestBody), 'muteHttpExceptions': true };
    const response = UrlFetchApp.fetch(API_ENDPOINT, options);
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      const text = json.candidates?.[0]?.content?.parts?.[0]?.text || "";
      Logger.log(`[OCR Debug] ${label} Response (First 100 chars): ${text.substring(0, 100)}...`);
      return text;
    }
    throw new Error(`API Error (${label}): ${response.getResponseCode()}`);
  }, 3, 1000);
}

function parseSinglePeriodCsvResult(text) {
  if (!text) return [];
  
  // Markdownコードブロックの削除と行分割
  // ` `` ` の文字列を直接指定
  const cleanText = text.replace("```csv", '').replace("```", '').trim();
  const lines = cleanText.split('\n');
  const results = [];
  
  for (const line of lines) {
    // カンマで分割（金額にカンマが含まれている可能性も考慮し、最後のカンマで区切る等の工夫も考えられるが、
    // プロンプトでカンマ削除を指示しているので標準的なsplitで対応）
    // ただし、行末のカンマや空行には注意
    if (!line.trim()) continue;

    // 最後のカンマで分割して「科目」と「金額」に分ける（科目にカンマが入るケースへの安全策）
    const lastCommaIndex = line.lastIndexOf(',');
    if (lastCommaIndex === -1) continue; // カンマがない行はスキップ

    const itemPart = line.substring(0, lastCommaIndex).trim();
    const amountPart = line.substring(lastCommaIndex + 1).trim();

    if (itemPart && amountPart) {
      const item = normalizeJapaneseText(itemPart);
      // ヘッダー行っぽいものはスキップ（"勘定科目" や "金額" という文字そのもの）
      if (item === '勘定科目' || item === '科目' || item === '金額') continue;
      if (!isSafeText(item)) continue;

      const amount = parseJapaneseNumber(amountPart);
      // 金額が0の場合は、パース失敗の可能性もあるが、本当に0円の可能性もあるため採用する
      // ただし、金額部分が明らかに数値でない文字列だった場合は parseJapaneseNumber が 0 を返す仕様なので
      // ここではそのまま採用。
      
      results.push([item, amount]);
    }
  }
  Logger.log(`[OCR Debug] Parsed ${results.length} items from text.`);
  return results;
}

function normalizeJapaneseText(text) {
  if (!text) return "";
  let normalized = String(text).normalize('NFKC').trim();
  // 基本的なクリーニング（スペース除去など）
  normalized = normalized.replace(/\s+/g, '');
  return normalized;
}

function parseJapaneseNumber(val) {
  if (!val) return 0;
  let s = String(val).trim().replace(/,/g, '').replace(/[¥円]/g, '');
  if (s.startsWith('▲') || s.startsWith('△')) {
    s = '-' + s.substring(1);
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function isSafeText(text) {
  return !/^[^=+\-@]/.test(text); // 数式インジェクション対策
}

function callWithRetry(func, maxRetries, delay) {
  for (let i = 0; i < maxRetries; i++) {
    try { return func(); } catch (e) {
      if (i === maxRetries - 1) throw e;
      Utilities.sleep(delay);
    }
  }
}

// ダイアログ用: フォルダ一覧取得
function getRecentFolders() {
  const folders = DriveApp.searchFolders('trashed=false');
  const folderList = [];
  let count = 0;
  while (folders.hasNext() && count < 50) {
    const f = folders.next();
    folderList.push({ id: f.getId(), name: f.getName(), lastUpdated: f.getLastUpdated().getTime() });
    count++;
  }
  folderList.sort((a, b) => b.lastUpdated - a.lastUpdated);
  
  const result = [{ id: DriveApp.getRootFolder().getId(), name: 'マイドライブ (ルート)', date: '-' }];
  folderList.slice(0, 20).forEach(f => {
    result.push({ id: f.id, name: f.name, date: new Date(f.lastUpdated).toLocaleDateString() });
  });
  return result;
}

// ダイアログ用: ファイル一覧取得
function getPDFFilesInFolder(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.searchFiles('mimeType="application/pdf" and trashed=false');
    const fileList = [];
    while (files.hasNext()) {
      const f = files.next();
      fileList.push({ id: f.getId(), name: f.getName(), date: f.getLastUpdated().toLocaleDateString(), lastUpdated: f.getLastUpdated().getTime() });
    }
    fileList.sort((a, b) => b.lastUpdated - a.lastUpdated);
    return fileList;
  } catch (e) {
    throw new Error(`フォルダ読み込みエラー: ${e.message}`);
  }
}

// =================================================================================
// ヘルパー関数 (旧コードより移植 - parseItemValueCsvResultを追加)
// =================================================================================
function parseItemValueCsvResult(text) {
  // 4列 (項目, 2ago, prev, curr)
  // ` ``` ` の文字列を直接指定
  const lines = text.replace("```csv", '').replace("```", '').trim().split('\n');
  const results = [];
  for (const line of lines) {
    const parts = line.split(',');
    if (parts.length >= 4) {
      const item = normalizeJapaneseText(parts[0]);
      if (!isSafeText(item)) continue;
      results.push([
        item,
        parseJapaneseNumber(parts[1]),
        parseJapaneseNumber(parts[2]),
        parseJapaneseNumber(parts[3])
      ]);
    }
  }
  return results;
}
