/**
 * =================================================================================
 * グローバル設定
 * =================================================================================
 */

/** @typedef {{item: string, value: string}} MappingEntry */
/** @typedef {{name: string, id: string}} FileInfo */
const VERSION = '3.7.0'; // スクリプトバージョン管理（v3.7.0: マッピング提案に勘定科目マスターを使用）

// API KEYのグローバルキャッシュ（パフォーマンス最適化）
let CACHED_API_KEY = null;

/**
 * Gemini API KEYを取得（キャッシュ利用）
 * @return {string} API KEY
 */
function getApiKey() {
  if (CACHED_API_KEY === null) {
    CACHED_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!CACHED_API_KEY) {
      throw new Error('GEMINI_API_KEYが設定されていません。スクリプトプロパティで設定してください。\n設定方法: プロジェクトの設定 > スクリプトプロパティ > プロパティを追加');
    }
    Logger.log('API KEYをキャッシュしました');
  }
  return CACHED_API_KEY;
}

const CONFIG = Object.freeze({
  VERSION: VERSION,
  OCR_SHEET_NAME: 'OCR作業シート',
  MAPPING_SHEET_NAME: '勘定科目マッピング',
  RECONCILIATION_LOG_SHEET: '調整ログ',

  // ========================================================
  // 新規追加：エクセル転記先シート設定
  // ========================================================
  EXCEL_TRANSFER_CONFIG: Object.freeze({
    SHEET_NAME: '４．３期比較表',
    PERIOD_COLUMNS: Object.freeze({
      前々期: { 金額: 'B', 売上比: 'C', 前年比: null, 対年変更: null },
      前期:   { 金額: 'D', 売上比: 'E', 前年比: 'F', 対年変更: 'G' },
      当期:   { 金額: 'H', 売上比: 'I', 前年比: 'J', 対年変更: 'K' }
    }),
    PERIOD_COLUMNS_EXTENDED: Object.freeze({
      前々期: { 金額: 'N', 売上比: 'O', 前年比: null, 対年変更: null },
      前期:   { 金額: 'P', 売上比: 'Q', 前年比: 'R', 対年変更: 'S' },
      当期:   { 金額: 'T', 売上比: 'U', 前年比: 'V', 対年変更: 'W' }
    }),
    // 各表の詳細設定
    TABLES: Object.freeze({
      '①売上高内訳表': Object.freeze({
        tableNumber: '①',
        tableName: '売上高内訳表',
        columnLayout: 'standard',  // A-K列を使用
        headerRow: 5,
        dataStartRow: 6,
        dataEndRow: 8,
        totalRow: 9,
        itemColumn: 'A',
        dataRowCount: 3,
        type: 'simple',  // シンプル転記型
        description: '売上高を項目別に転記'
      }),
      '②販売費及び一般管理費比較表': Object.freeze({
        tableNumber: '②',
        tableName: '販売費及び一般管理費比較表',
        columnLayout: 'extended',  // M-W列を使用
        headerRow: 5,
        dataStartRow: 6,
        dataEndRow: 28,
        subtotalRow: 12,
        itemColumn: 'M',
        dataRowCount: 6,
        type: 'classified',  // 分類集計型
        groups: Object.freeze({
          人件費: {
            dataStartRow: 6,
            dataEndRow: 10,  // 人件費項目は6-10行（集計行の直前）
            maxDataRows: 5,
            otherRow: 11,    // その他人件費
            subtotalRow: 12,
            label: '（人件費小計）'
          },
          その他経費: {
            dataStartRow: 13,
            dataEndRow: 28,  // その他経費の最終行
            maxDataRows: 15,
            otherRow: 29,    // その他は29行に固定
            label: 'その他'
          }
        }),
        description: '人件費とその他経費に分類して転記'
      }),
      '③変動費内訳比較表': Object.freeze({
        tableNumber: '③',
        tableName: '変動費内訳比較表',
        columnLayout: 'standard',  // A-K列を使用
        headerRow: 13,
        dataStartRow: 14,
        dataEndRow: 21,
        subtotalRow: 17,
        totalRow: 20,
        itemColumn: 'A',
        dataRowCount: 8,
        type: 'classified',  // グループ分類型に変更
        groups: Object.freeze({
          '変動費': {
            dataStartRow: 14,
            dataEndRow: 15,
            maxDataRows: 2,
            label: '変動費'
          },
          '期首棚卸高': {
            dataStartRow: 18,
            dataEndRow: 18,
            maxDataRows: 1,
            label: '期首棚卸高'
          },
          '期末棚卸高': {
            dataStartRow: 19,
            dataEndRow: 21,
            maxDataRows: 3,
            otherRow: 21,
            label: '期末棚卸高'
          }
        }),
        description: '変動費、期首棚卸高、期末棚卸高に分類して転記'
      }),
      '④製造経費比較表': Object.freeze({
        tableNumber: '④',
        tableName: '製造経費比較表',
        columnLayout: 'standard',  // A-K列を使用
        headerRow: 24,
        dataStartRow: 25,
        dataEndRow: 39,
        subtotalRow1: 30,
        subtotalRow2: 40,
        totalRow: 41,
        itemColumn: 'A',
        dataRowCount: 15,
        type: 'classified_by_category',  // 費目別分類型
        groups: Object.freeze({
          '労務費': {
            dataStartRow: 25,
            dataEndRow: 29,
            maxDataRows: 4,
            subtotalRow: 30,
            label: '（労務費計）'
          },
          'その他経費': {
            dataStartRow: 31,
            dataEndRow: 39,
            maxDataRows: 10,
            otherRow: 39,
            label: 'その他経費'
          }
        }),
        categories: Object.freeze({
          '賃金': {
            row: 25,
            name: '賃金給料',
            aggregateItems: ['賃金', '給料', '賃金給料']
          },
          '賞与': {
            row: 26,
            name: '賞与',
            aggregateItems: ['賞与', 'ボーナス']
          },
          '法定福利費': {
            row: 27,
            name: '法定福利費',
            aggregateItems: ['法定福利', '法定福利費']
          },
          '福利厚生費': {
            row: 28,
            name: '福利厚生費',
            aggregateItems: ['福利厚生', '福利厚生費']
          },
          '労務費計': {
            row: 30,
            name: '（労務費計）',
            type: 'subtotal'
          },
          'その他経費': {
            row: 39,
            name: 'その他経費',
            aggregateItems: ['経費', 'その他']
          },
          '経費計': {
            row: 40,
            name: '（経費計）',
            type: 'subtotal'
          }
        }),
        description: '製造経費を賃金/賞与/福利等で分類し集計'
      }),
      '⑤その他損益比較表': Object.freeze({
        tableNumber: '⑤',
        tableName: 'その他損益比較表',
        columnLayout: 'extended',  // M-W列を使用
        headerRow: 34,
        dataStartRow: 35,
        dataEndRow: 44,
        totalRow: 44,
        itemColumn: 'M',
        dataRowCount: 10,
        type: 'simple_aggregate',  // シンプル集計型
        categories: Object.freeze({
          '営業外収益': {
            row: 37,
            name: '営業外収益'
          },
          '営業外費用': {
            row: 38,
            name: '営業外費用'
          },
          '特別利益': {
            row: 40,
            name: '特別利益'
          },
          '特別損失': {
            row: 41,
            name: '特別損失'
          }
        }),
        description: '営業外収益/費用、特別利益/損失を集計'
      })
    }),

    // ========================================================
    // BS（３期比較表ＢＳ）用設定
    // ========================================================
    BS_SHEET: Object.freeze({
      SHEET_NAME: '３期比較表ＢＳ',
      PERIOD_COLUMNS: Object.freeze({
        前々期: { 金額: 'D', 構成比: 'E' },
        前期:   { 金額: 'F', 構成比: 'G', 前年比: 'H' },
        当期:   { 金額: 'I', 構成比: 'J', 前年比: 'K' }
      }),
      GROUPS: Object.freeze({
        '流動資産': {
          dataStartRow: 4,
          dataEndRow: 10,
          subtotalRow: 11,
          label: '棚卸資産計'
        },
        '固定資産': {
          dataStartRow: 12,
          dataEndRow: 20,
          subtotalRow: 21,
          label: '固定資産計'
        },
        '流動負債': {
          dataStartRow: 23,
          dataEndRow: 27,
          subtotalRow: 32,
          label: '流動負債計'
        },
        '固定負債': {
          dataStartRow: 28,
          dataEndRow: 34,
          subtotalRow: 35,
          label: '固定負債計'
        },
        '純資産': {
          dataStartRow: 37,
          dataEndRow: 39,
          subtotalRow: 39,
          label: '純資産合計'
        }
      }),
      ITEM_COLUMN: 'A'
    })
  }),

  // API設定
  API: Object.freeze({
    MAX_RETRIES: 3,
    RETRY_DELAY_MS: 1000,
    TIMEOUT_MS: 60000
  })
});

/**
 * =================================================================================
 * メニュー設定
 * =================================================================================
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('決算書OCR (v' + CONFIG.VERSION + ')')
    .addItem('【ステップ１】ファイル選択＆OCR実行', 'startStep1_showFileDialog')
    .addSeparator()
    .addItem('【ステップ２】統合マッピングシート準備', 'startStep2_createUnifiedMappingSheet')
    .addSeparator()
    .addItem('【ステップ３】PL/BSへデータを転記', 'startStep3_transferDataToExcel')
    .addSeparator()
    .addItem('【管理】転記ログを表示', 'showTransferLog')
    .addItem('【デバッグ】テスト実行', 'debugTest')
    .addItem('【デバッグ】基本ロギングテスト', 'testBasicLogging')
    .addToUi();
}

/**
 * デバッグ用テスト関数
 */
function debugTest() {
  Logger.log('=== DEBUG TEST START ===');

  try {
    Logger.log('STEP 1: CONFIG check');
    Logger.log(`CONFIG.VERSION: ${CONFIG.VERSION}`);

    Logger.log('STEP 2: Sheet access');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`Active sheet: ${ss.getName()}`);

    Logger.log('STEP 3: OCR sheet check');
    const ocrSheet = ss.getSheetByName(CONFIG.OCR_SHEET_NAME);
    Logger.log(`OCR sheet found: ${ocrSheet ? 'YES' : 'NO'}`);

    Logger.log('STEP 4: Mapping sheet check');
    const mappingSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET_NAME);
    Logger.log(`Mapping sheet found: ${mappingSheet ? 'YES' : 'NO'}`);

    Logger.log('STEP 5: Target sheet check');
    const targetSheet = ss.getSheetByName(CONFIG.EXCEL_TRANSFER_CONFIG.SHEET_NAME);
    Logger.log(`Target sheet found: ${targetSheet ? 'YES' : 'NO'} (looking for: ${CONFIG.EXCEL_TRANSFER_CONFIG.SHEET_NAME})`);

    Logger.log('=== DEBUG TEST COMPLETE ===');
  } catch (e) {
    Logger.log(`ERROR at step: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * =================================================================================
 * 【ステップ１】ファイル選択とOCR実行
 * =================================================================================
 */
function startStep1_showFileDialog() {
  const html = HtmlService.createHtmlOutputFromFile('dialog.html').setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '処理する決算書PDFを選択');
}

function setFileIdAndExtract(fileId) {
  const file = DriveApp.getFileById(fileId);
  if (!file) {
    SpreadsheetApp.getUi().alert('ファイルが見つかりませんでした。');
    return;
  }
  SpreadsheetApp.getUi().alert('OCR抽出を開始します。処理には1分ほどかかる場合があります。完了したら再度通知します。');
  try {
    const resultText = callGeminiApi(file);
    const results = parseItemValueCsvResult(resultText);
    if (results.length > 0) {
      const ocrSheet = getOrCreateSheet(CONFIG.OCR_SHEET_NAME);
      ocrSheet.clear();
      
      // ヘッダー行を追加（3期分対応）
      const headers = [['勘定科目', '前々期', '前期', '当期']];
      ocrSheet.getRange(1, 1, 1, 4).setValues(headers)
        .setFontWeight('bold')
        .setBackground('#E8EAED')
        .setHorizontalAlignment('center');
      
      // OCRデータを書き込み（4列：項目名、前々期、前期、当期）
      ocrSheet.getRange(2, 1, results.length, 4).setValues(results)
        .setHorizontalAlignment('left')
        .setNumberFormat('@STRING@');  // 項目名は文字列、金額は数値として表示
      
      // 列幅を調整
      ocrSheet.setColumnWidth(1, 300);  // 勘定科目列を広く
      ocrSheet.setColumnWidths(2, 3, 120);  // 金額列は標準幅
      ocrSheet.setFrozenRows(1);  // ヘッダー行を固定
      
      SpreadsheetApp.getUi().alert(`OCR抽出が完了し、「${CONFIG.OCR_SHEET_NAME}」に出力しました。\n\n抽出項目数: ${results.length}件\n3期分（前々期、前期、当期）のデータを取得しました。\n\n次に「ステップ２」を実行してください。`);
    } else {
      SpreadsheetApp.getUi().alert('有効なデータを抽出できませんでした。');
    }
  } catch (e) {
    Logger.log(`エラー詳細: ${e.message}\nスタック: ${e.stack}`);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + e.message);
  }
}

/**
 * 3つの決算書PDFから3期分のデータを抽出する
 * @param {string} fileCurrentId - 当期の決算書PDFファイルID
 * @param {string} filePreviousId - 前期の決算書PDFファイルID
 * @param {string} file2PeriodsAgoId - 前々期の決算書PDFファイルID
 */
function setMultipleFilesAndExtract(fileCurrentId, filePreviousId, file2PeriodsAgoId) {
  try {
    SpreadsheetApp.getUi().alert('3期分のOCR抽出を開始します。処理には数分かかる場合があります。完了したら再度通知します。');

    // 3つのファイルを取得
    const fileCurrent = DriveApp.getFileById(fileCurrentId);
    const filePrevious = DriveApp.getFileById(filePreviousId);
    const file2PeriodsAgo = DriveApp.getFileById(file2PeriodsAgoId);

    if (!fileCurrent || !filePrevious || !file2PeriodsAgo) {
      SpreadsheetApp.getUi().alert('いずれかのファイルが見つかりませんでした。');
      return;
    }

    // 3つのPDFから個別にOCR抽出
    Logger.log('当期の決算書をOCR抽出中...');
    const resultCurrent = callGeminiApiSinglePeriod(fileCurrent, '当期');
    const dataCurrent = parseSinglePeriodCsvResult(resultCurrent);
    Logger.log(`当期: ${dataCurrent.length}件抽出完了`);

    Logger.log('前期の決算書をOCR抽出中...');
    const resultPrevious = callGeminiApiSinglePeriod(filePrevious, '前期');
    const dataPrevious = parseSinglePeriodCsvResult(resultPrevious);
    Logger.log(`前期: ${dataPrevious.length}件抽出完了`);

    Logger.log('前々期の決算書をOCR抽出中...');
    const result2PeriodsAgo = callGeminiApiSinglePeriod(file2PeriodsAgo, '前々期');
    const data2PeriodsAgo = parseSinglePeriodCsvResult(result2PeriodsAgo);
    Logger.log(`前々期: ${data2PeriodsAgo.length}件抽出完了`);

    // 3期分のデータを統合（勘定科目をキーにマージ）
    const mergedData = merge3PeriodData(data2PeriodsAgo, dataPrevious, dataCurrent);

    if (mergedData.length > 0) {
      const ocrSheet = getOrCreateSheet(CONFIG.OCR_SHEET_NAME);
      ocrSheet.clear();

      // ヘッダー行を追加（3期分対応）
      const headers = [['勘定科目', '前々期', '前期', '当期']];
      ocrSheet.getRange(1, 1, 1, 4).setValues(headers)
        .setFontWeight('bold')
        .setBackground('#E8EAED')
        .setHorizontalAlignment('center');

      // OCRデータを書き込み（4列：項目名、前々期、前期、当期）
      ocrSheet.getRange(2, 1, mergedData.length, 4).setValues(mergedData)
        .setHorizontalAlignment('left')
        .setNumberFormat('@STRING@');  // 項目名は文字列、金額は数値として表示

      // 列幅を調整
      ocrSheet.setColumnWidth(1, 300);  // 勘定科目列を広く
      ocrSheet.setColumnWidths(2, 3, 120);  // 金額列は標準幅
      ocrSheet.setFrozenRows(1);  // ヘッダー行を固定

      SpreadsheetApp.getUi().alert(`3期分のOCR抽出が完了し、「${CONFIG.OCR_SHEET_NAME}」に出力しました。\n\n統合項目数: ${mergedData.length}件\n前々期: ${data2PeriodsAgo.length}件\n前期: ${dataPrevious.length}件\n当期: ${dataCurrent.length}件\n\n次に「ステップ２」を実行してください。`);
    } else {
      SpreadsheetApp.getUi().alert('有効なデータを抽出できませんでした。');
    }
  } catch (e) {
    Logger.log(`エラー詳細: ${e.message}\nスタック: ${e.stack}`);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + e.message);
  }
}

/**
 * Gemini APIを呼び出して単一期のPDFから勘定科目と金額を抽出
 * @param {GoogleAppsScript.Drive.File} file - 抽出対象のPDFファイル
 * @param {string} periodLabel - 期の名称（デバッグ用）
 * @return {string} 抽出されたテキスト
 */
function callGeminiApiSinglePeriod(file, periodLabel) {
  return callWithRetry(() => {
    const API_KEY = getApiKey();
    const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + API_KEY;
    const prompt = `このファイルには、「貸借対照表」「損益計算書」「製造原価報告書」「販売費及び一般管理費内訳書」が含まれています。これらの書類から、記載されているすべての勘定科目と金額を抽出してください。

【重要】
- 各勘定科目について、記載されている金額を抽出してください
- 金額が記載されていない項目は「0」と出力してください
- 出力形式：「勘定科目,金額」のCSV形式
- 金額に通貨記号(¥)や桁区切りのカンマ(,)は含めず、半角数字のみで出力してください

【抽出順序】
1. 「貸借対照表」のすべての項目を上から順に
2. 「損益計算書」のすべての項目を上から順に
3. 「製造原価報告書」のすべての項目を上から順に（存在する場合）
4. 「販売費及び一般管理費内訳書」のすべての項目を上から順に（存在する場合）

【対象項目】
- 小計、合計、内訳項目も含め、勘定科目と金額が記載されている行は例外なくすべてを抽出対象とします
- タイトル、日付、会社名など、勘定科目ではないテキスト行は無視してください

【出力例】
流動資産,50000000
現金及び預金,20000000
売上高,123456789
材料費,5000000`;

    const requestBody = { "contents": [{"parts": [ {"text": prompt}, {"inline_data": {"mime_type": file.getMimeType(), "data": Utilities.base64Encode(file.getBlob().getBytes())}} ] }], "generationConfig": { "temperature": 0.0, "topP": 1, "maxOutputTokens": 32768 } };
    const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(requestBody), 'muteHttpExceptions': true };
    const response = UrlFetchApp.fetch(API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const json = JSON.parse(responseBody);
      if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0].text) {
        return json.candidates[0].content.parts[0].text;
      } else {
        return "";
      }
    } else {
      throw new Error(`APIリクエストに失敗しました (${periodLabel})。ステータスコード: ${responseCode}\nレスポンス: ${responseBody}`);
    }
  }, CONFIG.API.MAX_RETRIES, CONFIG.API.RETRY_DELAY_MS);
}

/**
 * 単一期のCSV形式のOCR結果をパースする
 * @param {string} text - CSV形式のテキスト（勘定科目,金額）
 * @return {Array<Array<string>>} [[勘定科目, 金額], ...] の配列
 */
function parseSinglePeriodCsvResult(text) {
  const cleanedText = text.replace(/```csv\n?/g, '').replace(/```/g, '').trim();
  if (!cleanedText) return [];

  const lines = cleanedText.split('\n');
  const results = [];
  const seenItems = new Set();

  for (const line of lines) {
    const parts = line.split(',');

    if (parts.length >= 2) {
      let item = parts[0].trim();
      item = normalizeJapaneseText(item);

      if (!isSafeText(item)) {
        Logger.log(`警告: 安全でない項目名を検出してスキップしました: ${item}`);
        continue;
      }

      const amount = parseJapaneseNumber(parts[1].trim());

      if (seenItems.has(item)) {
        Logger.log(`警告: 重複項目を検出しました: ${item} (既存の値を保持します)`);
        continue;
      }

      seenItems.add(item);
      results.push([item, amount]);
    }
  }

  return results;
}

/**
 * 3期分のデータを統合する
 * @param {Array<Array>} data2PeriodsAgo - 前々期データ [[勘定科目, 金額], ...]
 * @param {Array<Array>} dataPrevious - 前期データ [[勘定科目, 金額], ...]
 * @param {Array<Array>} dataCurrent - 当期データ [[勘定科目, 金額], ...]
 * @return {Array<Array>} 統合データ [[勘定科目, 前々期, 前期, 当期], ...]
 */
function merge3PeriodData(data2PeriodsAgo, dataPrevious, dataCurrent) {
  Logger.log('========== 3期分データ統合開始（当期基準・類似度マッチング対応） ==========');

  // 勘定科目をキーにマップを作成（正規化後の項目名をキーとする）
  const itemMap = new Map();
  const normalizedToOriginalMap = new Map(); // 正規化後 → 元の項目名（当期優先）

  // ステップ1: 当期のデータを基準として登録
  Logger.log(`\n[ステップ1] 当期データを基準として登録（${dataCurrent.length}件）`);
  dataCurrent.forEach(([item, amount]) => {
    const normalizedItem = normalizeJapaneseText(item);
    itemMap.set(normalizedItem, {
      current: amount,
      previous: 0,
      twoPeriodsAgo: 0,
      originalName: item  // 当期の項目名を保存
    });
    normalizedToOriginalMap.set(normalizedItem, item);
    Logger.log(`  当期: "${item}" (正規化: "${normalizedItem}") = ${amount}`);
  });

  // ステップ2: 前期のデータをマッチング
  Logger.log(`\n[ステップ2] 前期データをマッチング（${dataPrevious.length}件）`);
  const currentItemNames = Array.from(normalizedToOriginalMap.values());

  dataPrevious.forEach(([item, amount]) => {
    const normalizedItem = normalizeJapaneseText(item);

    // 2-1: 正規化後の完全一致を試みる
    if (itemMap.has(normalizedItem)) {
      itemMap.get(normalizedItem).previous = amount;
      Logger.log(`  前期: "${item}" → 完全一致 "${itemMap.get(normalizedItem).originalName}" = ${amount}`);
    } else {
      // 2-2: 類似度マッチングを試みる
      const bestMatch = findBestMatch(item, currentItemNames, 0.7);
      if (bestMatch) {
        const normalizedMatch = normalizeJapaneseText(bestMatch);
        itemMap.get(normalizedMatch).previous = amount;
        Logger.log(`  前期: "${item}" → 類似マッチ "${bestMatch}" (類似度70%以上) = ${amount}`);
      } else {
        // 2-3: マッチングしない場合は新規項目として追加（当期=0）
        itemMap.set(normalizedItem, {
          current: 0,
          previous: amount,
          twoPeriodsAgo: 0,
          originalName: item
        });
        Logger.log(`  前期: "${item}" → 新規項目（当期に該当なし）= ${amount}`);
      }
    }
  });

  // ステップ3: 前々期のデータをマッチング
  Logger.log(`\n[ステップ3] 前々期データをマッチング（${data2PeriodsAgo.length}件）`);

  data2PeriodsAgo.forEach(([item, amount]) => {
    const normalizedItem = normalizeJapaneseText(item);

    // 3-1: 正規化後の完全一致を試みる
    if (itemMap.has(normalizedItem)) {
      itemMap.get(normalizedItem).twoPeriodsAgo = amount;
      Logger.log(`  前々期: "${item}" → 完全一致 "${itemMap.get(normalizedItem).originalName}" = ${amount}`);
    } else {
      // 3-2: 類似度マッチングを試みる
      const bestMatch = findBestMatch(item, currentItemNames, 0.7);
      if (bestMatch) {
        const normalizedMatch = normalizeJapaneseText(bestMatch);
        itemMap.get(normalizedMatch).twoPeriodsAgo = amount;
        Logger.log(`  前々期: "${item}" → 類似マッチ "${bestMatch}" (類似度70%以上) = ${amount}`);
      } else {
        // 3-3: マッチングしない場合は新規項目として追加（当期=0、前期は既存値を保持）
        if (itemMap.has(normalizedItem)) {
          // すでに前期で追加されている場合
          itemMap.get(normalizedItem).twoPeriodsAgo = amount;
          Logger.log(`  前々期: "${item}" → 既存項目に追加 = ${amount}`);
        } else {
          // 前々期のみに存在する項目
          itemMap.set(normalizedItem, {
            current: 0,
            previous: 0,
            twoPeriodsAgo: amount,
            originalName: item
          });
          Logger.log(`  前々期: "${item}" → 新規項目（当期・前期に該当なし）= ${amount}`);
        }
      }
    }
  });

  // ステップ4: マップを配列に変換 [勘定科目, 前々期, 前期, 当期]
  // 項目名は当期の名称を優先、当期に存在しない場合は前期→前々期の順で使用
  const mergedData = [];
  itemMap.forEach((amounts, normalizedItem) => {
    mergedData.push([
      amounts.originalName,  // 元の項目名（当期優先）
      amounts.twoPeriodsAgo,
      amounts.previous,
      amounts.current
    ]);
  });

  Logger.log(`\n========== 統合結果: ${mergedData.length}件の勘定科目を統合しました ==========`);

  // 統合結果のサマリー
  const currentOnlyCount = mergedData.filter(([_, a, b, c]) => c !== 0 && a === 0 && b === 0).length;
  const previousOnlyCount = mergedData.filter(([_, a, b, c]) => b !== 0 && c === 0 && a === 0).length;
  const twoPeriodsAgoOnlyCount = mergedData.filter(([_, a, b, c]) => a !== 0 && b === 0 && c === 0).length;
  const allPeriodsCount = mergedData.filter(([_, a, b, c]) => a !== 0 && b !== 0 && c !== 0).length;

  Logger.log(`  - 3期すべてに存在: ${allPeriodsCount}件`);
  Logger.log(`  - 当期のみ: ${currentOnlyCount}件`);
  Logger.log(`  - 前期のみ: ${previousOnlyCount}件`);
  Logger.log(`  - 前々期のみ: ${twoPeriodsAgoOnlyCount}件`);
  Logger.log(`  - その他（2期に存在）: ${mergedData.length - currentOnlyCount - previousOnlyCount - twoPeriodsAgoOnlyCount - allPeriodsCount}件`);

  return mergedData;
}

/**
 * =================================================================================
 * 【ステップ２】統合マッピングシートの準備（勘定科目マスター参照版）
 * =================================================================================
 */
function startStep2_createUnifiedMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ocrSheet = ss.getSheetByName(CONFIG.OCR_SHEET_NAME);
  if (!ocrSheet || ocrSheet.getLastRow() === 0) {
    SpreadsheetApp.getUi().alert('先に「ステップ１」を実行して、「' + CONFIG.OCR_SHEET_NAME + '」にデータを抽出してください。');
    return;
  }
  const sourceItems = ocrSheet.getRange(2, 1, ocrSheet.getLastRow() - 1, 1).getValues().flat().filter(String);  // ヘッダー行をスキップして項目名のみ取得

  // 転記先シートの確認
  const targetSheetPL = ss.getSheetByName(CONFIG.EXCEL_TRANSFER_CONFIG.SHEET_NAME);
  const targetSheetBS = ss.getSheetByName(CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET.SHEET_NAME);
  if (!targetSheetPL || !targetSheetBS) {
    SpreadsheetApp.getUi().alert('転記先のシート「' + CONFIG.EXCEL_TRANSFER_CONFIG.SHEET_NAME + '」または「' + CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET.SHEET_NAME + '」が見つかりません。');
    return;
  }

  // ★追加：勘定科目マスターの読み込み
  const MASTER_SHEET_NAME = '勘定科目マスター';
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  let masterData = [];
  let useMaster = false;

  if (masterSheet && masterSheet.getLastRow() > 1) {
    // マスターデータの読み込み（A列:勘定科目, B列:転記先分類, C列:詳細グループ）
    const masterRange = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 3);
    const rawMasterData = masterRange.getValues();
    
    // 検索用に正規化したデータを準備
    masterData = rawMasterData.map(row => {
      if (!row[0]) return null;
      return {
        name: normalizeJapaneseText(row[0], true), // 比較用に正規化
        originalName: row[0], // 表示用
        table: row[1],
        group: row[2]
      };
    }).filter(item => item !== null);
    
    if (masterData.length > 0) {
      useMaster = true;
      Logger.log(`勘定科目マスターから${masterData.length}件のデータを読み込みました。`);
    }
  }

  // マッピングシートを作成/クリア
  const mappingSheet = getOrCreateSheet(CONFIG.MAPPING_SHEET_NAME);
  mappingSheet.clear();
  mappingSheet.setFrozenRows(1);

  // ヘッダー行を設定
  const headerValues = [['OCR抽出項目', '転記先分類（推奨）', '詳細グループ（推奨）', '転記先行（任意）', '転記先列（任意）', '説明']];
  mappingSheet.getRange('A1:F1').setValues(headerValues)
    .setFontWeight('bold')
    .setBackground('#4A86E8')
    .setFontColor('#FFFFFF');

  // データ行を準備（自動分類を適用）
  const dataRows = [];
  const descriptions = [];

  sourceItems.forEach(item => {
    const normalizedItem = normalizeJapaneseText(item, true);  // 勘定科目名: 強い正規化
    
    let suggestedTable = '';
    let suggestedGroup = '';
    let description = '自動推奨なし（手動で設定）';
    let isMatched = false;

    // 1. マスターデータとの完全一致チェック
    if (useMaster) {
      const exactMatch = masterData.find(m => m.name === normalizedItem);
      if (exactMatch) {
        suggestedTable = exactMatch.table;
        suggestedGroup = exactMatch.group;
        description = '★マスター完全一致';
        isMatched = true;
      }
    }

    // 2. マスターデータとの類似度チェック（完全一致がなく、マスターがある場合）
    if (!isMatched && useMaster) {
      let bestScore = 0;
      let bestMatch = null;

      // 全マスター項目と比較して最も似ているものを探す
      for (const mData of masterData) {
        const score = calculateSimilarity(normalizedItem, mData.name);
        if (score > bestScore) {
          bestScore = score;
          bestMatch = mData;
        }
      }

      // 類似度70%以上なら採用
      if (bestScore >= 0.7 && bestMatch) {
        suggestedTable = bestMatch.table;
        suggestedGroup = bestMatch.group;
        description = `☆マスター類似一致(${Math.round(bestScore * 100)}%): ${bestMatch.originalName}`;
        isMatched = true;
      }
    }

    // 3. 従来のキーワード判定ロジック（マスターで決まらなかった場合）
    if (!isMatched) {
      const classification = classifyItemToTable(normalizedItem);
      if (classification && classification.tableKey) {
        suggestedTable = classification.tableKey;
        suggestedGroup = classification.categoryKey || '';
        description = 'キーワード自動判定';
      }
    }

    // D列・E列は空白（手動指定する場合のみ入力）
    dataRows.push([item, suggestedTable, suggestedGroup, '', '']);
    descriptions.push([description]);
  });

  // データ行に値を一括設定
  if (dataRows.length > 0) {
    mappingSheet.getRange(2, 1, dataRows.length, 5).setValues(dataRows);
    mappingSheet.getRange(2, 6, descriptions.length, 1).setValues(descriptions).setFontColor('#999999');
  }

  // B列：転記先分類のドロップダウン（PLとBSを統合）
  const tableOptions = [
    // PL
    '①売上高内訳表', '②販売費及び一般管理費比較表', '③変動費内訳比較表', '④製造経費比較表', '⑤その他損益比較表',
    // BS
    'BS:流動資産', 'BS:固定資産', 'BS:流動負債', 'BS:固定負債', 'BS:純資産',
    // 製造原価
    '製造原価報告書'
  ];
  const tableRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(tableOptions, true)
    .setAllowInvalid(true)
    .build();
  mappingSheet.getRange(2, 2, sourceItems.length, 1).setDataValidation(tableRule);

  // C列：詳細グループのドロップダウン（PLとBSを統合）
  const groupRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      // PL
      '人件費', 'その他', 'その他経費', '変動費', '期首棚卸高', '期末棚卸高',
      '営業外収益', '営業外費用', '特別利益', '特別損失',
      // 製造原価
      '材料費', '労務費', '経費', '期首仕掛品', '期末仕掛品',
      // BS
      '現金・預金', '売上債権', '棚卸資産', 'その他流動資産',
      '有形固定資産', '無形固定資産', '投資その他',
      '仕入債務', '短期借入金', '未払金', 'その他流動負債',
      '長期借入金', '社債', '退職給付引当金',
      '資本金', '資本剰余金', '利益剰余金', '自己株式'
    ], true)
    .setAllowInvalid(true)
    .build();
  mappingSheet.getRange(2, 3, sourceItems.length, 1).setDataValidation(groupRule);

  // 列幅を調整
  mappingSheet.setColumnWidths(1, 1, 250);  // A列: OCR抽出項目
  mappingSheet.setColumnWidths(2, 1, 220);  // B列: 転記先分類
  mappingSheet.setColumnWidths(3, 1, 180);  // C列: 詳細グループ
  mappingSheet.setColumnWidths(4, 1, 100);  // D列: 転記先行
  mappingSheet.setColumnWidths(5, 1, 100);  // E列: 転記先列
  mappingSheet.setColumnWidths(6, 1, 250);  // F列: 説明

  SpreadsheetApp.setActiveSheet(mappingSheet);

  // ヘルプメッセージ
  let helpMessage = `「${CONFIG.MAPPING_SHEET_NAME}」を準備しました。\n\n`;
  
  if (useMaster) {
    helpMessage += `✅ 「${MASTER_SHEET_NAME}」を使用して自動分類を行いました。\n`;
    helpMessage += `・完全一致または類似度70%以上の項目はマスターの設定を反映しています。\n`;
    helpMessage += `・「説明」列で判定根拠（★マスター完全一致、☆類似一致など）を確認できます。\n\n`;
  } else {
    helpMessage += `ℹ️ 「${MASTER_SHEET_NAME}」が見つからないか空のため、標準のキーワード判定のみ行いました。\n`;
    helpMessage += `（精度向上のため、マスターシートの作成をお勧めします）\n\n`;
  }

  helpMessage += `【記入方法】
1. B列・C列：自動推奨値を確認し、必要に応じて変更してください。
2. D列・E列（任意）：転記先の行・列を固定したい場合のみ入力してください。

✅ 内容を確認後、「ステップ３」を実行してください。`;

  SpreadsheetApp.getUi().alert(helpMessage);
}



/**
 * =================================================================================
 * 【ステップ２】統合マッピングシートの準備（勘定科目マスター参照版）
 * =================================================================================
 */
function startStep2_createUnifiedMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ocrSheet = ss.getSheetByName(CONFIG.OCR_SHEET_NAME);
  if (!ocrSheet || ocrSheet.getLastRow() === 0) {
    SpreadsheetApp.getUi().alert('先に「ステップ１」を実行して、「' + CONFIG.OCR_SHEET_NAME + '」にデータを抽出してください。');
    return;
  }
  const sourceItems = ocrSheet.getRange(2, 1, ocrSheet.getLastRow() - 1, 1).getValues().flat().filter(String);  // ヘッダー行をスキップして項目名のみ取得

  // 転記先シートの確認
  const targetSheetPL = ss.getSheetByName(CONFIG.EXCEL_TRANSFER_CONFIG.SHEET_NAME);
  const targetSheetBS = ss.getSheetByName(CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET.SHEET_NAME);
  if (!targetSheetPL || !targetSheetBS) {
    SpreadsheetApp.getUi().alert('転記先のシート「' + CONFIG.EXCEL_TRANSFER_CONFIG.SHEET_NAME + '」または「' + CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET.SHEET_NAME + '」が見つかりません。');
    return;
  }

  // ★追加：勘定科目マスターの読み込み
  const MASTER_SHEET_NAME = '勘定科目マスター';
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  let masterData = [];
  let useMaster = false;

  if (masterSheet && masterSheet.getLastRow() > 1) {
    // マスターデータの読み込み（A列:勘定科目, B列:転記先分類, C列:詳細グループ）
    const masterRange = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 3);
    const rawMasterData = masterRange.getValues();
    
    // 検索用に正規化したデータを準備
    masterData = rawMasterData.map(row => {
      if (!row[0]) return null;
      return {
        name: normalizeJapaneseText(row[0], true), // 比較用に正規化
        originalName: row[0], // 表示用
        table: row[1],
        group: row[2]
      };
    }).filter(item => item !== null);
    
    if (masterData.length > 0) {
      useMaster = true;
      Logger.log(`勘定科目マスターから${masterData.length}件のデータを読み込みました。`);
    }
  }

  // マッピングシートを作成/クリア
  const mappingSheet = getOrCreateSheet(CONFIG.MAPPING_SHEET_NAME);
  mappingSheet.clear();
  mappingSheet.setFrozenRows(1);

  // ヘッダー行を設定
  const headerValues = [['OCR抽出項目', '転記先分類（推奨）', '詳細グループ（推奨）', '転記先行（任意）', '転記先列（任意）', '説明']];
  mappingSheet.getRange('A1:F1').setValues(headerValues)
    .setFontWeight('bold')
    .setBackground('#4A86E8')
    .setFontColor('#FFFFFF');

  // データ行を準備（自動分類を適用）
  const dataRows = [];
  const descriptions = [];

  sourceItems.forEach(item => {
    const normalizedItem = normalizeJapaneseText(item, true);  // 勘定科目名: 強い正規化
    
    let suggestedTable = '';
    let suggestedGroup = '';
    let description = '自動推奨なし（手動で設定）';
    let isMatched = false;

    // 1. マスターデータとの完全一致チェック
    if (useMaster) {
      const exactMatch = masterData.find(m => m.name === normalizedItem);
      if (exactMatch) {
        suggestedTable = exactMatch.table;
        suggestedGroup = exactMatch.group;
        description = '★マスター完全一致';
        isMatched = true;
      }
    }

    // 2. マスターデータとの類似度チェック（完全一致がなく、マスターがある場合）
    if (!isMatched && useMaster) {
      let bestScore = 0;
      let bestMatch = null;

      // 全マスター項目と比較して最も似ているものを探す
      for (const mData of masterData) {
        const score = calculateSimilarity(normalizedItem, mData.name);
        if (score > bestScore) {
          bestScore = score;
          bestMatch = mData;
        }
      }

      // 類似度70%以上なら採用
      if (bestScore >= 0.7 && bestMatch) {
        suggestedTable = bestMatch.table;
        suggestedGroup = bestMatch.group;
        description = `☆マスター類似一致(${Math.round(bestScore * 100)}%): ${bestMatch.originalName}`;
        isMatched = true;
      }
    }

    // 3. 従来のキーワード判定ロジック（マスターで決まらなかった場合）
    if (!isMatched) {
      const classification = classifyItemToTable(normalizedItem);
      if (classification && classification.tableKey) {
        suggestedTable = classification.tableKey;
        suggestedGroup = classification.categoryKey || '';
        description = 'キーワード自動判定';
      }
    }

    // D列・E列は空白（手動指定する場合のみ入力）
    dataRows.push([item, suggestedTable, suggestedGroup, '', '']);
    descriptions.push([description]);
  });

  // データ行に値を一括設定
  if (dataRows.length > 0) {
    mappingSheet.getRange(2, 1, dataRows.length, 5).setValues(dataRows);
    mappingSheet.getRange(2, 6, descriptions.length, 1).setValues(descriptions).setFontColor('#999999');
  }

  // B列：転記先分類のドロップダウン（PLとBSを統合）
  const tableOptions = [
    // PL
    '①売上高内訳表', '②販売費及び一般管理費比較表', '③変動費内訳比較表', '④製造経費比較表', '⑤その他損益比較表',
    // BS
    'BS:流動資産', 'BS:固定資産', 'BS:流動負債', 'BS:固定負債', 'BS:純資産',
    // 製造原価
    '製造原価報告書'
  ];
  const tableRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(tableOptions, true)
    .setAllowInvalid(true)
    .build();
  mappingSheet.getRange(2, 2, sourceItems.length, 1).setDataValidation(tableRule);

  // C列：詳細グループのドロップダウン（PLとBSを統合）
  const groupRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      // PL
      '人件費', 'その他', 'その他経費', '変動費', '期首棚卸高', '期末棚卸高',
      '営業外収益', '営業外費用', '特別利益', '特別損失',
      // 製造原価
      '材料費', '労務費', '経費', '期首仕掛品', '期末仕掛品',
      // BS
      '現金・預金', '売上債権', '棚卸資産', 'その他流動資産',
      '有形固定資産', '無形固定資産', '投資その他',
      '仕入債務', '短期借入金', '未払金', 'その他流動負債',
      '長期借入金', '社債', '退職給付引当金',
      '資本金', '資本剰余金', '利益剰余金', '自己株式'
    ], true)
    .setAllowInvalid(true)
    .build();
  mappingSheet.getRange(2, 3, sourceItems.length, 1).setDataValidation(groupRule);

  // 列幅を調整
  mappingSheet.setColumnWidths(1, 1, 250);  // A列: OCR抽出項目
  mappingSheet.setColumnWidths(2, 1, 220);  // B列: 転記先分類
  mappingSheet.setColumnWidths(3, 1, 180);  // C列: 詳細グループ
  mappingSheet.setColumnWidths(4, 1, 100);  // D列: 転記先行
  mappingSheet.setColumnWidths(5, 1, 100);  // E列: 転記先列
  mappingSheet.setColumnWidths(6, 1, 250);  // F列: 説明

  SpreadsheetApp.setActiveSheet(mappingSheet);

  // ヘルプメッセージ
  let helpMessage = `「${CONFIG.MAPPING_SHEET_NAME}」を準備しました。\n\n`;
  
  if (useMaster) {
    helpMessage += `✅ 「${MASTER_SHEET_NAME}」を使用して自動分類を行いました。\n`;
    helpMessage += `・完全一致または類似度70%以上の項目はマスターの設定を反映しています。\n`;
    helpMessage += `・「説明」列で判定根拠（★マスター完全一致、☆類似一致など）を確認できます。\n\n`;
  } else {
    helpMessage += `ℹ️ 「${MASTER_SHEET_NAME}」が見つからないか空のため、標準のキーワード判定のみ行いました。\n`;
    helpMessage += `（精度向上のため、マスターシートの作成をお勧めします）\n\n`;
  }

  helpMessage += `【記入方法】
1. B列・C列：自動推奨値を確認し、必要に応じて変更してください。
2. D列・E列（任意）：転記先の行・列を固定したい場合のみ入力してください。

✅ 内容を確認後、「ステップ３」を実行してください。`;

  SpreadsheetApp.getUi().alert(helpMessage);
}



/**
 * =================================================================================
 * ヘルパー関数群（内部処理用）
 * =================================================================================
 */

/**
 * 日本語文字列を正規化（全角/半角、空白、Unicode正規化、勘定科目の表記揺れ対応）
 * @param {string} text - 正規化する文字列
 * @param {boolean} isAccountItem - true: 勘定科目名（強い正規化）、false: テーブル名等（軽い正規化）
 * @return {string} 正規化された文字列
 */
function normalizeJapaneseText(text, isAccountItem = true) {
  try {
    if (!text) return '';
    // Unicode正規化（NFKC: 互換文字を標準形に統一）
    let normalized = String(text).normalize('NFKC');
    // 前後の空白を削除
    normalized = normalized.trim();

    // 勘定科目の場合のみ強い正規化を適用
    if (isAccountItem) {
      // 1. 括弧内の補足情報を除去（例：「減価償却費（製造）」→「減価償却費」）
      normalized = normalized.replace(/[（(].*?[）)]/g, '');

      // 2. 「及び」「及」「および」を統一
      normalized = normalized.replace(/及び|及|および/g, '及');

      // 3. 区切り文字を統一（「・」「、」「,」「　」などを削除）
      normalized = normalized.replace(/[・、,\u3000]+/g, '');

      // 4. 長音記号を統一（「ー」「－」「—」「―」など）
      normalized = normalized.replace(/[－—―]/g, 'ー');

      // 5. 連続する空白を削除
      normalized = normalized.replace(/\s+/g, '');
    } else {
      // テーブル名/グループ名の場合は軽い正規化のみ（空白の統一のみ）
      normalized = normalized.replace(/\s+/g, ' ');
    }

    // 前後の空白を再度削除
    normalized = normalized.trim();

    return normalized;
  } catch (e) {
    Logger.log(`警告: テキスト正規化エラー: ${e.message}`);
    return String(text || '').trim();
  }
}

/**
 * 2つの文字列のレーベンシュタイン距離（編集距離）を計算
 * @param {string} str1 - 比較する文字列1
 * @param {string} str2 - 比較する文字列2
 * @return {number} 編集距離（小さいほど類似）
 */
function levenshteinDistance(str1, str2) {
  const len1 = str1.length;
  const len2 = str2.length;
  const matrix = [];

  // 初期化
  for (let i = 0; i <= len1; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= len2; j++) {
    matrix[0][j] = j;
  }

  // 動的計画法で編集距離を計算
  for (let i = 1; i <= len1; i++) {
    for (let j = 1; j <= len2; j++) {
      const cost = str1[i - 1] === str2[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,     // 削除
        matrix[i][j - 1] + 1,     // 挿入
        matrix[i - 1][j - 1] + cost  // 置換
      );
    }
  }

  return matrix[len1][len2];
}

/**
 * 2つの文字列の類似度を計算（0.0〜1.0、1.0が完全一致）
 * @param {string} str1 - 比較する文字列1
 * @param {string} str2 - 比較する文字列2
 * @return {number} 類似度（0.0〜1.0）
 */
function calculateSimilarity(str1, str2) {
  if (str1 === str2) return 1.0;
  if (!str1 || !str2) return 0.0;

  const distance = levenshteinDistance(str1, str2);
  const maxLen = Math.max(str1.length, str2.length);

  return 1.0 - (distance / maxLen);
}

/**
 * 当期の項目リストから、前期・前々期の項目に最も類似する項目を見つける
 * @param {string} targetItem - マッチング対象の項目名（前期または前々期）
 * @param {Array<string>} currentItems - 当期の項目名リスト
 * @param {number} threshold - 類似度の閾値（デフォルト: 0.7）
 * @return {string|null} マッチした当期の項目名、またはnull
 */
function findBestMatch(targetItem, currentItems, threshold = 0.7) {
  let bestMatch = null;
  let bestScore = threshold;

  const normalizedTarget = normalizeJapaneseText(targetItem);

  currentItems.forEach(currentItem => {
    const normalizedCurrent = normalizeJapaneseText(currentItem);
    const similarity = calculateSimilarity(normalizedTarget, normalizedCurrent);

    if (similarity > bestScore) {
      bestScore = similarity;
      bestMatch = currentItem;
    }
  });

  if (bestMatch) {
    Logger.log(`類似度マッチング: "${targetItem}" → "${bestMatch}" (類似度: ${bestScore.toFixed(2)})`);
  }

  return bestMatch;
}

/**
 * 日本の会計書類特有の数値表現を堅牢に解析
 * @param {string} valueStr - 解析する数値文字列
 * @return {number|string} 解析された数値、または元の文字列
 */
function parseJapaneseNumber(valueStr) {
  try {
    if (!valueStr) return '';

    let str = String(valueStr).trim();

    // 空文字列やダッシュは空文字列として返す
    if (str === '' || str === '−' || str === '-' || str === '—' || str === '―') {
      return '';
    }

    // ▲や△は負数を表す（日本の会計書類の慣例）
    let isNegative = false;
    if (str.startsWith('▲') || str.startsWith('△')) {
      isNegative = true;
      str = str.substring(1).trim();
    }

    // 括弧付きも負数を表す
    if (str.startsWith('(') && str.endsWith(')')) {
      isNegative = true;
      str = str.substring(1, str.length - 1).trim();
    }

    // 通貨記号と桁区切りカンマを削除
    str = str.replace(/[¥$,]/g, '');

    // 単位の処理（千円、百万円など）
    let multiplier = 1;
    if (str.includes('千円') || str.includes('千')) {
      multiplier = 1000;
      str = str.replace(/千円?/g, '');
    } else if (str.includes('百万円') || str.includes('百万')) {
      multiplier = 1000000;
      str = str.replace(/百万円?/g, '');
    } else if (str.includes('億円') || str.includes('億')) {
      multiplier = 100000000;
      str = str.replace(/億円?/g, '');
    }

    // 円を削除
    str = str.replace(/円/g, '').trim();

    // 数値に変換
    const num = parseFloat(str);
    if (isNaN(num)) {
      return valueStr; // 解析できない場合は元の文字列を返す
    }

    const result = num * multiplier * (isNegative ? -1 : 1);
    return result;
  } catch (e) {
    Logger.log(`警告: 数値解析エラー (入力: "${valueStr}"): ${e.message}`);
    return valueStr; // エラー時は元の値を返す
  }
}

/**
 * テキストに数式やコードインジェクションが含まれていないか検証
 * @param {string} text - 検証するテキスト
 * @return {boolean} 安全な場合true
 */
function isSafeText(text) {
  if (!text) return true;
  const str = String(text);

  // 数式の開始記号をチェック
  if (/^[=+\-@]/.test(str)) {
    Logger.log(`セキュリティ警告: 数式として解釈される可能性のある文字列を検出: ${str.substring(0, 50)}`);
    return false;
  }

  // スクリプトタグや危険な文字列を検出
  if (/<script|javascript:|onerror=|onclick=|onload=/i.test(str)) {
    Logger.log(`セキュリティ警告: 危険なスクリプトコードを検出: ${str.substring(0, 50)}`);
    return false;
  }

  // 制御文字をチェック
  if (/[\u0000-\u001f\u007f]/i.test(str)) {
    Logger.log(`セキュリティ警告: 制御文字を検出: ${str.substring(0, 50)}`);
    return false;
  }

  return true;
}

/**
 * API呼び出しをリトライ/バックオフで実行
 * @param {Function} apiCall - API呼び出し関数
 * @param {number} maxRetries - 最大リトライ回数
 * @param {number} delayMs - リトライ間隔（ミリ秒）
 * @return {*} API呼び出し結果
 */
function callWithRetry(apiCall, maxRetries, delayMs) {
  let lastError;
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return apiCall();
    } catch (e) {
      lastError = e;
      Logger.log(`API呼び出し失敗 (試行 ${attempt + 1}/${maxRetries}): ${e.message}`);
      if (attempt < maxRetries - 1) { // 最後の試行でなければ待機
        Utilities.sleep(delayMs * Math.pow(2, attempt)); // 指数バックオフ
      }
    }
  }
  throw new Error(`API呼び出しが${maxRetries}回失敗しました: ${lastError.message}`);
}

/**
 * Gemini APIを呼び出してPDFから勘定科目と金額を抽出（リトライ機能付き）
 * @param {GoogleAppsScript.Drive.File} file - 抽出対象のPDFファイル
 * @return {string} 抽出されたテキスト
 */
function callGeminiApi(file) {
  return callWithRetry(() => {
    const API_KEY = getApiKey();  // キャッシュされたAPI KEYを取得
    const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + API_KEY;
    const prompt = `このファイルには、「貸借対照表」「損益計算書」「製造原価報告書」「販売費及び一般管理費内訳書」の3期比較表が含まれています。これらの書類から、記載されているすべての勘定科目と3期分の金額（前々期、前期、当期）を抽出してください。

【重要】
- 各勘定科目について、前々期・前期・当期の3つの金額を必ず抽出してください
- 金額が記載されていない期は「0」と出力してください
- 出力形式：「勘定科目,前々期金額,前期金額,当期金額」のCSV形式
- 金額に通貨記号(¥)や桁区切りのカンマ(,)は含めず、半角数字のみで出力してください

【抽出順序】
1. 「貸借対照表」のすべての項目を上から順に
2. 「損益計算書」のすべての項目を上から順に
3. 「製造原価報告書」のすべての項目を上から順に（存在する場合）
4. 「販売費及び一般管理費内訳書」のすべての項目を上から順に（存在する場合）

【対象項目】
- 小計、合計、内訳項目も含め、勘定科目と金額が記載されている行は例外なくすべてを抽出対象とします
- タイトル、日付、会社名など、勘定科目ではないテキスト行は無視してください

【出力例】
流動資産,45000000,48000000,50000000
現金及び預金,18000000,19000000,20000000
売上高,110000000,115000000,123456789
材料費,0,0,5000000`;
    const requestBody = { "contents": [{"parts": [ {"text": prompt}, {"inline_data": {"mime_type": file.getMimeType(), "data": Utilities.base64Encode(file.getBlob().getBytes())}} ] }], "generationConfig": { "temperature": 0.0, "topP": 1, "maxOutputTokens": 32768 } };
    const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(requestBody), 'muteHttpExceptions': true };
    const response = UrlFetchApp.fetch(API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    if (responseCode === 200) {
      const json = JSON.parse(responseBody);
      if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0].text) {
        return json.candidates[0].content.parts[0].text;
      } else {
        return "";
      }
    } else {
      throw new Error(`APIリクエストに失敗しました。ステータスコード: ${responseCode}\nレスポンス: ${responseBody}`);
    }
  }, CONFIG.API.MAX_RETRIES, CONFIG.API.RETRY_DELAY_MS);
}

/**
 * CSV形式のOCR結果をパースし、正規化とバリデーションを適用（3期分対応）
 * @param {string} text - CSV形式のテキスト（項目名,前々期,前期,当期）
 * @return {Array<Array<string>>} [[項目名, 前々期, 前期, 当期], ...] の配列
 */
function parseItemValueCsvResult(text) {
  const cleanedText = text.replace(/```csv\n?/g, '').replace(/```/g, '').trim();
  if (!cleanedText) return [];

  const lines = cleanedText.split('\n');
  const results = [];
  const seenItems = new Set(); // 重複チェック用

  for (const line of lines) {
    const parts = line.split(',');

    // 3期分のデータ（項目名 + 前々期 + 前期 + 当期 = 4列）が必要
    if (parts.length >= 4) {
      let item = parts[0].trim();

      // 日本語テキストを正規化
      item = normalizeJapaneseText(item);

      // 安全性チェック（数式インジェクション防止）
      if (!isSafeText(item)) {
        Logger.log(`警告: 安全でない項目名を検出してスキップしました: ${item}`);
        continue;
      }

      // 3期分の金額を解析
      const amount2PeriodsAgo = parseJapaneseNumber(parts[1].trim()); // 前々期
      const amount1PeriodAgo = parseJapaneseNumber(parts[2].trim());  // 前期
      const currentAmount = parseJapaneseNumber(parts[3].trim());     // 当期

      // 重複項目のログ記録
      if (seenItems.has(item)) {
        Logger.log(`警告: 重複項目を検出しました: ${item} (既存の値を保持します)`);
        continue;
      }

      seenItems.add(item);
      results.push([item, amount2PeriodsAgo, amount1PeriodAgo, currentAmount]);

      Logger.log(`✓ 解析成功: ${item} (前々期:${amount2PeriodsAgo}, 前期:${amount1PeriodAgo}, 当期:${currentAmount})`);
    } else if (parts.length >= 2) {
      // 後方互換性：2列形式（旧形式）の場合は当期のみとして扱う
      let item = parts[0].trim();
      item = normalizeJapaneseText(item);

      if (!isSafeText(item)) continue;
      if (seenItems.has(item)) continue;

      const currentAmount = parseJapaneseNumber(parts[1].trim());
      seenItems.add(item);
      results.push([item, 0, 0, currentAmount]); // 前々期・前期は0

      Logger.log(`⚠ 2列形式（旧形式）で解析: ${item} (当期のみ:${currentAmount})`);
    }
  }

  Logger.log(`OCR結果: ${results.length}件の項目を抽出しました（3期分対応）`);
  return results;
}

function getFolders() {
  const folders = DriveApp.getRootFolder().getFolders();
  const folderList = [];
  while (folders.hasNext()) {
    let folder = folders.next();
    folderList.push({ name: folder.getName(), id: folder.getId() });
  }
  return folderList;
}

function getPDFFilesInFolder(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.searchFiles('mimeType="application/pdf"');
    const fileList = [];
    while (files.hasNext()) {
      let file = files.next();
      fileList.push({ name: file.getName(), id: file.getId() });
    }
    return fileList;
  } catch (e) {
    Logger.log(`エラー: PDFファイル検索に失敗しました: ${e.message}`);
    throw new Error(`フォルダからPDFファイルを取得できませんでした: ${e.message}`);
  }
}

/**
 * 直近のPDFファイルを取得する（ダイアログ用）
 * @return {Array<Object>} ファイル情報の配列
 */
function getRecentPdfFiles() {
  // 直近更新されたPDFを取得
  // 注意: searchFilesは順序保証がないため、多めに取得してソートする
  const files = DriveApp.searchFiles('mimeType="application/pdf" and trashed=false');
  const fileList = [];
  let count = 0;
  
  // 最大50件取得してソート
  while (files.hasNext() && count < 50) {
    const file = files.next();
    fileList.push({
      id: file.getId(),
      name: file.getName(),
      lastUpdated: file.getLastUpdated().getTime()
    });
    count++;
  }
  
  // 更新日時で降順ソート（新しい順）
  fileList.sort((a, b) => b.lastUpdated - a.lastUpdated);
  
  // 上位20件を返す
  return fileList.slice(0, 20).map(f => ({
    id: f.id,
    name: f.name,
    date: new Date(f.lastUpdated).toLocaleDateString() + ' ' + new Date(f.lastUpdated).toLocaleTimeString()
  }));
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * 調整ログシートに転記結果を記録（バッチ操作で最適化）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @param {Array<Array>} logEntries - ログエントリの配列
 */
function writeReconciliationLog(ss, logEntries) {
  const logSheet = getOrCreateSheet(CONFIG.RECONCILIATION_LOG_SHEET);

  if (logEntries.length === 0) {
    Logger.log('記録するログエントリがありません');
    return;
  }

  const headerRow = logEntries[0];
  const dataRows = logEntries.slice(1);

  // ヘッダーが存在しない場合は初期化
  if (logSheet.getLastRow() === 0) {
    logSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow])
      .setFontWeight('bold')
      .setBackground('#4A86E8')
      .setFontColor('#FFFFFF');
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidths(1, headerRow.length, 150);
  }

  // データ行がない場合は終了
  if (dataRows.length === 0) {
    Logger.log('記録するデータがありません');
    return;
  }

  // 背景色を一括で準備（バッチ操作用）
  const backgroundColors = dataRows.map(entry => {
    const status = entry[4]; // ステータス列
    let color = '#FFFFFF'; // デフォルト白

    if (status === '成功') {
      color = '#D9EAD3'; // 緑
    } else if (status === 'エラー') {
      color = '#F4CCCC'; // 赤
    } else if (status === '未マッピング') {
      color = '#FFF2CC'; // 黄
    }

    // 行全体に同じ色を適用
    return new Array(headerRow.length).fill(color);
  });

  // 一括書き込み（パフォーマンス最適化）
  const startRow = logSheet.getLastRow() + 1;
  const range = logSheet.getRange(startRow, 1, dataRows.length, headerRow.length);

  range.setValues(dataRows);            // 1回のAPI呼び出しで全データ書き込み
  range.setBackgrounds(backgroundColors); // 1回のAPI呼び出しで全背景色設定

  Logger.log(`調整ログシートに${dataRows.length}件のエントリを記録しました（バッチ操作）`);
}

/**
 * 数式セルを保護（誤って上書きしないように）- バッチ最適化版
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 保護するシート
 * @param {Object} formulasConfig - 数式設定オブジェクト
 */
function protectFormulaCells(sheet, formulasConfig) {
  if (!formulasConfig || Object.keys(formulasConfig).length === 0) return;

  try {
    // 範囲をまとめて取得（API呼び出し削減）
    const cellAddresses = Object.keys(formulasConfig);

    // 各範囲に対して保護を設定
    // 注: Protection APIは個別呼び出しが必要なため、完全なバッチ化は不可
    // しかし、範囲取得は事前にまとめて行うことで若干最適化
    cellAddresses.forEach(cellAddress => {
      const range = sheet.getRange(cellAddress);
      const protection = range.protect().setDescription('自動計算セル（編集しないでください）');
      // すべてのユーザーが編集できないように設定（スクリプトからは編集可能）
      protection.setWarningOnly(true); // 警告のみ（完全ロックではない）
    });

    Logger.log(`${sheet.getName()}シートの${cellAddresses.length}個の数式セルを保護しました`);
  } catch (e) {
    Logger.log(`数式セルの保護中にエラーが発生しました: ${e.message}`);
  }
}

/**
 * =================================================================================
 * 新規機能：Google Sheetsからエクセル用データへの転記機能
 * =================================================================================
 */

/**
 * Google SheetsのOCRデータをエクセル転記用に正規化・準備する
 * @param {Array<Array>} ocrData - [[項目名, 金額], ...] 形式のOCRデータ
 * @return {Array<Object>} {{itemName: string, amount: number, period: string}, ...}
 */
function prepareExcelTransferData(ocrData) {
  const transferData = [];

  ocrData.forEach(row => {
    if (row[0] && row[1] !== undefined && row[1] !== '') {
      const itemName = normalizeJapaneseText(row[0], true);  // 勘定科目名: 強い正規化
      const amount = parseJapaneseNumber(row[1]);

      // 安全性チェック
      if (!isSafeText(itemName)) {
        Logger.log(`警告: 安全でない項目を検出してスキップ: ${itemName}`);
        return;
      }

      transferData.push({
        itemName: itemName,
        amount: amount,
        amountOriginal: row[1]
      });
    }
  });

  Logger.log(`Excel転記用に${transferData.length}件のデータを準備しました`);
  return transferData;
}

/**
 * 項目名から該当する表と転記先を判定する（精度改善版）
 * マッピングがない場合の自動判定ロジック
 * @param {string} itemName - 勘定科目名
 * @return {Object} {{tableKey: string, categoryKey: string}} または null
 */
function classifyItemToTable(itemName) {
  const normalized = itemName.toLowerCase();

  // 集計行を除外
  if (normalized.match(/計$|合計|小計|総計|^計/)) {
    return null;
  }

  // ===== 優先度A: BS項目 =====
  // 純資産
  if (normalized.match(/資本金/)) return { tableKey: 'BS:純資産', categoryKey: '資本金' };
  if (normalized.match(/準備金/)) return { tableKey: 'BS:純資産', categoryKey: '資本剰余金' };
  if (normalized.match(/剰余金/)) return { tableKey: 'BS:純資産', categoryKey: '利益剰余金' };
  if (normalized.match(/自己株式/)) return { tableKey: 'BS:純資産', categoryKey: '自己株式' };
  // 固定負債
  if (normalized.match(/長期借入金/)) return { tableKey: 'BS:固定負債', categoryKey: '長期借入金' };
  if (normalized.match(/社債/)) return { tableKey: 'BS:固定負債', categoryKey: '社債' };
  if (normalized.match(/退職給付引当金/)) return { tableKey: 'BS:固定負債', categoryKey: '退職給付引当金' };
  // 流動負債
  if (normalized.match(/支払手形|買掛金/)) return { tableKey: 'BS:流動負債', categoryKey: '仕入債務' };
  if (normalized.match(/短期借入金/)) return { tableKey: 'BS:流動負債', categoryKey: '短期借入金' };
  if (normalized.match(/未払金|未払費用/)) return { tableKey: 'BS:流動負債', categoryKey: '未払金' };
  if (normalized.match(/前渡金|前払費用|未収入金|貸倒引当金/)) return { tableKey: 'BS:流動負債', categoryKey: 'その他流動負債' };
  // 固定資産
  if (normalized.match(/建物|構築物|機械|装置|車両|運搬具|工具|器具|備品|土地/)) return { tableKey: 'BS:固定資産', categoryKey: '有形固定資産' };
  if (normalized.match(/のれん|営業権/)) return { tableKey: 'BS:固定資産', categoryKey: '無形固定資産' };
  if (normalized.match(/投資有価証券|長期貸付金|繰延資産/)) return { tableKey: 'BS:固定資産', categoryKey: '投資その他' };
  // 流動資産
  if (normalized.match(/現金|預金/)) return { tableKey: 'BS:流動資産', categoryKey: '現金・預金' };
  if (normalized.match(/受取手形|売掛金/)) return { tableKey: 'BS:流動資産', categoryKey: '売上債権' };
  if (normalized.match(/商品|製品|仕掛品|原材料|貯蔵品/)) return { tableKey: 'BS:流動資産', categoryKey: '棚卸資産' };
  if (normalized.match(/前渡金|前払費用|未収入金|貸倒引当金/)) return { tableKey: 'BS:流動資産', categoryKey: 'その他流動資産' };

  // ===== 優先度B: 製造原価報告書項目 =====
  if (normalized.match(/期首.*仕掛品/)) return { tableKey: '製造原価報告書', categoryKey: '期首仕掛品' };
  if (normalized.match(/期末.*仕掛品/)) return { tableKey: '製造原価報告書', categoryKey: '期末仕掛品' };
  if (normalized.match(/材料費|原料費|主要材料費|買入部品費/)) return { tableKey: '製造原価報告書', categoryKey: '材料費' };
  if (normalized.match(/労務費|製造.*賃金|製造.*給料/)) return { tableKey: '製造原価報告書', categoryKey: '労務費' };
  if (normalized.match(/製造経費|工場経費|外注加工費|減価償却費/)) return { tableKey: '製造原価報告書', categoryKey: '経費' };


  // ===== 優先度C: PL項目（その他損益など、より具体的なものから）=====
  // ⑤ その他損益
  if (normalized.match(/受取利息|預金利息|貸付金利息|受取配当金|配当金|有価証券利息|雑収入|受取手数料/)) return { tableKey: '⑤その他損益比較表', categoryKey: '営業外収益' };
  if (normalized.match(/支払利息|借入金利息|社債利息|手形売却損|雑損失/)) return { tableKey: '⑤その他損益比較表', categoryKey: '営業外費用' };
  if (normalized.match(/固定資産売却益|投資有価証券売却益|特別利益/)) return { tableKey: '⑤その他損益比較表', categoryKey: '特別利益' };
  if (normalized.match(/固定資産売却損|固定資産除却損|減損損失|特別損失/)) return { tableKey: '⑤その他損益比較表', categoryKey: '特別損失' };

  // ③ 変動費・原価関連
  if (normalized.match(/期首.*棚卸|期首.*在庫|期首商品/)) return { tableKey: '③変動費内訳比較表', categoryKey: '期首棚卸高' };
  if (normalized.match(/期末.*棚卸|期末.*在庫|期末商品/)) return { tableKey: '③変動費内訳比較表', categoryKey: '期末棚卸高' };
  if (normalized.match(/商品仕入|期中仕入|当期.*仕入|仕入高|材料仕入|原材料費|材料費|副資材|外注.*費|外注加工費|製品仕入|商品.*原価/)) return { tableKey: '③変動費内訳比較表', categoryKey: '変動費' };

  // ④ 製造経費
  if (normalized.match(/製造.*給料|製造.*賃金|製造.*賞与|工場.*給料|工場.*賃金|作業員.*給料/)) return { tableKey: '④製造経費比較表', categoryKey: '労務費' };
  if (normalized.match(/製造経費|工場経費|製造.*費|動力費|燃料費|工具.*費|機械.*費|設備.*費|作業.*費/) && !normalized.match(/販売費|管理費/)) return { tableKey: '④製造経費比較表', categoryKey: 'その他経費' };

  // ② 販管費（人件費）
  if (normalized.match(/役員報酬|役員.*給料|給料.*手当|給与.*手当|賃金|賞与|ボーナス|雑給|法定福利費|福利厚生費|厚生費|退職金|退職給付/) && !normalized.match(/製造|工場/)) return { tableKey: '②販売費及び一般管理費比較表', categoryKey: '人件費' };

  // ② 販管費（その他経費）
  if (normalized.match(/支払家賃|地代家賃|賃借料|水道.*費|光熱費|電気代|ガス代|通信費|電話代|旅費.*交通費|交通費|旅費|出張.*費|広告.*費|宣伝費|販促費|荷造.*費|運賃|保険料|租税.*公課|公租公課|修繕費|消耗品費|事務.*費|会議費|交際費|接待.*費|寄付金|諸会費|雑費|減価償却費|リース料/) && !normalized.match(/製造|工場/)) return { tableKey: '②販売費及び一般管理費比較表', categoryKey: 'その他' };

  // ① 売上高
  if (normalized.match(/売上高|商品.*売上|製品.*売上|売上|収益|販売.*収入/) && !normalized.match(/売上原価|売上総利益/)) return { tableKey: '①売上高内訳表', categoryKey: null };

  Logger.log(`⚠ 自動分類できませんでした: ${itemName}（手動で設定してください）`);
  return null;
}
/**
 * Google Sheetsの範囲からセル値を取得（複数期対応）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {string} itemColumn - 項目列（例: 'A' または 'M'）
 * @param {number} row - 行番号
 * @param {Object} periodColumns - 期別列設定
 * @return {Object} {{itemName: string, amounts: {前々期: number, 前期: number, 当期: number}}}
 */
function readExcelRowData(sheet, itemColumn, row, periodColumns) {
  const itemCell = sheet.getRange(`${itemColumn}${row}`);
  const itemName = itemCell.getValue();

  if (!itemName) return null;

  const amounts = {};

  Object.keys(periodColumns).forEach(period => {
    const colObj = periodColumns[period];
    const amountCol = colObj.金額;
    const amountCell = sheet.getRange(`${amountCol}${row}`);
    amounts[period] = amountCell.getValue() || 0;
  });

  return {
    itemName: normalizeJapaneseText(itemName, true),  // 勘定科目名: 強い正規化
    amounts: amounts,
    row: row
  };
}

/**
 * Google Sheetsの複数行からデータを一括読み込み
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {string} itemColumn - 項目列
 * @param {number} startRow - 開始行
 * @param {number} endRow - 終了行
 * @param {Object} periodColumns - 期別列設定
 * @return {Array<Object>} {{itemName, amounts, row}, ...}
 */
function readExcelTableData(sheet, itemColumn, startRow, endRow, periodColumns) {
  const tableData = [];

  for (let row = startRow; row <= endRow; row++) {
    const rowData = readExcelRowData(sheet, itemColumn, row, periodColumns);
    if (rowData && rowData.itemName) {
      tableData.push(rowData);
    }
  }

  return tableData;
}

/**
 * Google Sheetsのセルに複数期の金額を書き込む（個別書き込み用・非推奨）
 * @deprecated バッチ操作の方が高速です。writeExcelBatchDataを使用してください。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {string} itemColumn - 項目列（デフォルト）
 * @param {number} row - 行番号（デフォルト）
 * @param {string} itemName - 項目名
 * @param {Object} amounts - {{前々期: number, 前期: number, 当期: number}}
 * @param {Object} periodColumns - 期別列設定
 * @param {number|null} targetRow - マッピングシートで指定された行（優先）
 * @param {string|null} targetColumn - マッピングシートで指定された列（優先）
 */
function writeExcelRowData(sheet, itemColumn, row, itemName, amounts, periodColumns, targetRow = null, targetColumn = null) {
  try {
    // マッピングシートで指定された行・列を優先
    const actualRow = targetRow || row;
    const actualItemColumn = targetColumn || itemColumn;

    // 項目名を書き込み
    if (actualItemColumn && actualRow) {
      sheet.getRange(`${actualItemColumn}${actualRow}`).setValue(itemName);
      Logger.log(`📝 項目名書き込み: "${itemName}" → ${actualItemColumn}${actualRow}`);
    }

    // 各期の金額を書き込み（0も含めてすべて書き込む）
    Object.keys(periodColumns).forEach(period => {
      const colObj = periodColumns[period];
      const amountCol = colObj.金額;
      const amount = amounts[period] || 0;

      if (amountCol && actualRow) {
        sheet.getRange(`${amountCol}${actualRow}`).setValue(amount);
        Logger.log(`💰 金額書き込み: ${period}=${amount} → ${amountCol}${actualRow}`);
      }
    });

    Logger.log(`✓ 書き込み成功: "${itemName}" → ${itemColumn}${row} (前々期:${amounts['前々期']}, 前期:${amounts['前期']}, 当期:${amounts['当期']})`);
  } catch (e) {
    Logger.log(`✗ 書き込みエラー: "${itemName}" (行${row}) - ${e.message}`);
    throw e;
  }
}

/**
 * バッチ書き込み用：複数行のデータを一括でシートに書き込む（パフォーマンス最適化）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} startRow - 開始行番号
 * @param {number} startCol - 開始列番号（A=1, B=2, ...）
 * @param {Array<Array>} dataMatrix - 書き込むデータの2D配列
 */
function writeExcelBatchData(sheet, startRow, startCol, dataMatrix) {
  try {
    if (!dataMatrix || dataMatrix.length === 0) {
      Logger.log('書き込むデータがありません');
      return;
    }

    const numRows = dataMatrix.length;
    const numCols = dataMatrix[0].length;

    sheet.getRange(startRow, startCol, numRows, numCols).setValues(dataMatrix);

    Logger.log(`✓ バッチ書き込み成功: ${numRows}行 x ${numCols}列のデータを書き込みました`);
  } catch (e) {
    Logger.log(`✗ バッチ書き込みエラー: ${e.message}`);
    throw e;
  }
}

/**
 * データを当期金額で降順ソート
 * @param {Array<Object>} data - {{itemName, amounts, ...}} 形式
 * @return {Array<Object>} ソート済みデータ
 */
function sortByCurrentPeriod(data) {
  try {
    if (!Array.isArray(data)) {
      Logger.log(`警告: sortByCurrentPeriod - 配列でないデータを受け取りました`);
      return [];
    }
    return data.sort((a, b) => {
      const amountA = a.amounts ? a.amounts['当期'] || 0 : 0;
      const amountB = b.amounts ? b.amounts['当期'] || 0 : 0;
      return amountB - amountA;  // 降順
    });
  } catch (e) {
    Logger.log(`エラー: ソート処理失敗: ${e.message}`);
    return data; // エラー時は元のデータを返す
  }
}

/**
 * 複数のデータオブジェクトの金額を合計
 * @param {Array<Object>} dataList - {{amounts: {前々期, 前期, 当期}}, ...}
 * @return {Object} {{前々期: number, 前期: number, 当期: number}}
 */
function aggregateAmounts(dataList) {
  try {
    const result = {
      '前々期': 0,
      '前期': 0,
      '当期': 0
    };

    if (!Array.isArray(dataList)) {
      Logger.log(`警告: aggregateAmounts - 配列でないデータを受け取りました`);
      return result;
    }

    dataList.forEach(data => {
      if (data && data.amounts) {
        result['前々期'] += Number(data.amounts['前々期']) || 0;
        result['前期'] += Number(data.amounts['前期']) || 0;
        result['当期'] += Number(data.amounts['当期']) || 0;
      }
    });

    return result;
  } catch (e) {
    Logger.log(`エラー: 金額集計処理失敗: ${e.message}`);
    return { '前々期': 0, '前期': 0, '当期': 0 };
  }
}

/**
 * =================================================================================
 * 表別転記ロジック
 * =================================================================================
 */

/**
 * ① 売上高内訳表への転記
 * シンプル転記型：OCRデータをそのまま行順に転記
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array<Object>} ocrItems - {{itemName: string, amount: number}, ...}
 */
function transferToTable01_SalesBreakdown(sheet, ocrItems) {
  const tableConfig = CONFIG.EXCEL_TRANSFER_CONFIG.TABLES['①売上高内訳表'];
  const periodCols = CONFIG.EXCEL_TRANSFER_CONFIG.PERIOD_COLUMNS;

  Logger.log(`【①売上高内訳表】への転記を開始します（${ocrItems.length}件）`);

  if (ocrItems.length === 0) {
    Logger.log('警告: 転記対象の項目がありません');
    return [];
  }

  const transferLog = [];
  let writeRowIndex = 0;

  // 最大データ行数まで転記
  for (let i = 0; i < ocrItems.length && writeRowIndex < tableConfig.dataRowCount; i++) {
    const item = ocrItems[i];
    const currentRow = tableConfig.dataStartRow + writeRowIndex;

    if (item.itemName && item.amounts) {
      // マッピングシートで指定された行・列を優先
      writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols, item.targetRow, item.targetColumn);
      transferLog.push({
        tableNumber: '①',
        itemName: item.itemName,
        row: item.targetRow || currentRow,  // 実際に転記された行
        status: '成功',
        amounts: item.amounts
      });
      writeRowIndex++;
    }
  }

  // 合計を計算（オプション）
  const totalAmounts = aggregateAmounts(ocrItems);
  Logger.log(`①売上高内訳表: ${writeRowIndex}件を転記し、合計を計算しました`);

  return transferLog;
}

/**
 * ③ 変動費内訳比較表への転記
 * グループ分類型：変動費、期首棚卸高、期末棚卸高に分類し、金額順でソート
 * ユーザーマッピングで指定されたグループを使用
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array<Object>} ocrItems - {{itemName: string, amounts: {...}, mappedGroup: string}, ...}
 */
function transferToTable03_VariableCosts(sheet, ocrItems) {
  const tableConfig = CONFIG.EXCEL_TRANSFER_CONFIG.TABLES['③変動費内訳比較表'];
  const periodCols = CONFIG.EXCEL_TRANSFER_CONFIG.PERIOD_COLUMNS;
  const groups = tableConfig.groups;

  Logger.log(`【③変動費内訳比較表】への転記を開始します（${ocrItems.length}件）`);
  Logger.log(`③使用可能グループ: ${Object.keys(groups).join(', ')}`);

  const transferLog = [];

  if (ocrItems.length === 0) {
    Logger.log('警告: 転記対象の項目がありません');
    return [];
  }

  // 項目をグループに分類（ユーザーマッピングの mappedGroup を使用）
  const groupedItems = {
    '変動費': [],
    '期首棚卸高': [],
    '期末棚卸高': []
  };

  ocrItems.forEach(item => {
    const group = item.mappedGroup || '変動費';  // デフォルトは変動費
    Logger.log(`処理中: ${item.itemName} (mappedGroup="${item.mappedGroup}")`);
    if (groupedItems[group]) {
      groupedItems[group].push(item);
      Logger.log(`✓ グループ分類: ${item.itemName} → ${group}`);
    } else {
      Logger.log(`警告: ${item.itemName}のグループ「${group}」が無効です（有効値: ${Object.keys(groupedItems).join(', ')}）`);
    }
  });

  // 各グループの転記
  Object.keys(groupedItems).forEach(groupName => {
    const items = groupedItems[groupName];
    if (items.length === 0) return;

    const group = groups[groupName];
    const maxRows = group.maxDataRows;
    const sortedItems = sortByCurrentPeriod(items);

    Logger.log(`
③-${groupName}: ${sortedItems.length}件を転記`);

    // 個別転記
    for (let i = 0; i < Math.min(sortedItems.length, maxRows); i++) {
      const currentRow = group.dataStartRow + i;
      const item = sortedItems[i];

      // マッピングシートで指定された行・列を優先
      writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols, item.targetRow, item.targetColumn);

      transferLog.push({
        tableNumber: '③',
        group: groupName,
        itemName: item.itemName,
        row: item.targetRow || currentRow,  // 実際に転記された行
        status: '成功',
        amounts: item.amounts
      });

      Logger.log(`  ${item.itemName} (当期: ${item.amounts['当期']})`);
    }

    // 超過分は「その他」として合計
    if (sortedItems.length > maxRows && group.otherRow) {
      const excessItems = sortedItems.slice(maxRows);
      const excessAmounts = aggregateAmounts(excessItems);
      const otherRow = group.otherRow;

      writeExcelRowData(sheet, tableConfig.itemColumn, otherRow, 'その他', excessAmounts, periodCols);

      transferLog.push({
        tableNumber: '③',
        group: groupName,
        itemName: 'その他',
        row: otherRow,
        status: '成功（集計）',
        amounts: excessAmounts,
        aggregatedCount: excessItems.length
      });

      Logger.log(`  その他: ${excessItems.length}件を集計`);
    }
  });

  Logger.log(`③変動費内訳比較表: 転記完了`);

  return transferLog;
}

/**
 * ④ 製造経費比較表への転記
 * グループ分類型：労務費/その他経費に分類し、金額順でソート
 * ユーザーマッピングで指定されたグループを使用
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array<Object>} ocrItems - {{itemName: string, amounts: {...}, mappedGroup: string}, ...}
 */
function transferToTable04_ManufacturingExpenses(sheet, ocrItems) {
  const tableConfig = CONFIG.EXCEL_TRANSFER_CONFIG.TABLES['④製造経費比較表'];
  const periodCols = CONFIG.EXCEL_TRANSFER_CONFIG.PERIOD_COLUMNS;
  const groups = tableConfig.groups;

  Logger.log(`【④製造経費比較表】への転記を開始します（${ocrItems.length}件）`);

  const transferLog = [];

  if (ocrItems.length === 0) {
    Logger.log('警告: 転記対象の項目がありません');
    return [];
  }

  // 項目を労務費/その他経費に分類（ユーザーマッピングの mappedGroup を使用）
  const laborItems = [];      // 労務費：賃金、賞与、福利厚生など
  const otherItems = [];      // その他経費

  ocrItems.forEach(item => {
    const group = item.mappedGroup || '';

    if (group.includes('労務費')) {
      laborItems.push(item);
      Logger.log(`✓ グループ分類: ${item.itemName} → 労務費`);
    } else if (group.includes('その他')) {
      otherItems.push(item);
      Logger.log(`✓ グループ分類: ${item.itemName} → その他経費`);
    } else {
      // グループが空の場合は警告
      Logger.log(`⚠ 警告: ${item.itemName}のグループが未指定`);
    }
  });

  Logger.log(`分類結果: 労務費${laborItems.length}件、その他経費${otherItems.length}件`);

  // 労務費グループの転記
  if (laborItems.length > 0) {
    const sortedLabor = sortByCurrentPeriod(laborItems);
    const laborGroup = groups['労務費'];
    const maxLaborRows = laborGroup.maxDataRows;

    // 個別転記
    for (let i = 0; i < Math.min(sortedLabor.length, maxLaborRows); i++) {
      const currentRow = laborGroup.dataStartRow + i;
      const item = sortedLabor[i];

      // マッピングシートで指定された行・列を優先
      writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols, item.targetRow, item.targetColumn);

      transferLog.push({
        tableNumber: '④',
        group: '労務費',
        itemName: item.itemName,
        row: item.targetRow || currentRow,  // 実際に転記された行
        status: '成功',
        amounts: item.amounts
      });

      Logger.log(`④-労務費: ${item.itemName} (当期: ${item.amounts['当期']})`);
    }

    // 超過分は合計（最後の行に追記）
    if (sortedLabor.length > maxLaborRows) {
      const excessLabor = sortedLabor.slice(maxLaborRows);
      const excessAmounts = aggregateAmounts(excessLabor);
      const excessRow = laborGroup.dataEndRow;

      writeExcelRowData(sheet, tableConfig.itemColumn, excessRow, 'その他労務費', excessAmounts, periodCols);

      transferLog.push({
        tableNumber: '④',
        group: '労務費',
        itemName: 'その他労務費',
        row: excessRow,
        status: '成功（集計）',
        amounts: excessAmounts,
        aggregatedCount: excessLabor.length
      });

      Logger.log(`④-その他労務費: ${excessLabor.length}件を集計`);
    }
  }

  // その他経費グループの転記
  if (otherItems.length > 0) {
    const sortedOther = sortByCurrentPeriod(otherItems);
    const otherGroup = groups['その他経費'];
    const maxOtherRows = otherGroup.maxDataRows - 1;  // 最後の行は「その他」用に確保

    // 個別転記
    for (let i = 0; i < Math.min(sortedOther.length, maxOtherRows); i++) {
      const currentRow = otherGroup.dataStartRow + i;
      const item = sortedOther[i];

      // マッピングシートで指定された行・列を優先
      writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols, item.targetRow, item.targetColumn);

      transferLog.push({
        tableNumber: '④',
        group: 'その他経費',
        itemName: item.itemName,
        row: item.targetRow || currentRow,  // 実際に転記された行
        status: '成功',
        amounts: item.amounts
      });
    }

    // その他経費が行数を超える場合は「その他」として合計（29行に固定）
    if (sortedOther.length > maxOtherRows) {
      const excessOther = sortedOther.slice(maxOtherRows);
      const excessAmounts = aggregateAmounts(excessOther);
      const otherRow = 29;  // 29行に固定

      writeExcelRowData(sheet, tableConfig.itemColumn, otherRow, 'その他', excessAmounts, periodCols);

      transferLog.push({
        tableNumber: '④',
        group: 'その他経費',
        itemName: 'その他',
        row: otherRow,
        status: '成功（集計）',
        amounts: excessAmounts,
        aggregatedCount: excessOther.length
      });

      Logger.log(`④-その他: ${excessOther.length}件を集計 (行${otherRow})`);
    }
  }

  Logger.log(`④製造経費比較表: 労務費${sortedLabor.length}件、その他経費${sortedOther.length}件を転記完了`);

  return transferLog;
}

/**
 * ② 販売費及び一般管理費比較表への転記
 * グループ分類型：人件費/その他経費に分類し、金額順でソート
 * ユーザーマッピングで指定されたグループを使用
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array<Object>} ocrItems - {{itemName: string, amounts: {...}, mappedGroup: string}, ...}
 */
function transferToTable02_SGA(sheet, ocrItems) {
  const tableConfig = CONFIG.EXCEL_TRANSFER_CONFIG.TABLES['②販売費及び一般管理費比較表'];
  const periodCols = CONFIG.EXCEL_TRANSFER_CONFIG.PERIOD_COLUMNS_EXTENDED;
  const groups = tableConfig.groups;

  Logger.log(`【②販売費及び一般管理費比較表】への転記を開始します（${ocrItems.length}件）`);

  const transferLog = [];

  // 項目を人件費/その他経費に分類（ユーザーマッピングの mappedGroup を使用）
  const personnelItems = [];
  const otherItems = [];

  ocrItems.forEach(item => {
    const group = item.mappedGroup || '';

    if (group.includes('人件費')) {
      personnelItems.push(item);
      Logger.log(`✓ グループ分類: ${item.itemName} → 人件費`);
    } else if (group.includes('その他')) {
      otherItems.push(item);
      Logger.log(`✓ グループ分類: ${item.itemName} → その他経費`);
    } else {
      // グループが空の場合は警告
      Logger.log(`⚠ 警告: ${item.itemName}のグループが未指定`);
    }
  });

  Logger.log(`分類結果: 人件費${personnelItems.length}件、その他経費${otherItems.length}件`);

  // 人件費グループの転記
  const sortedPersonnel = sortByCurrentPeriod(personnelItems);
  const personnelGroup = groups.人件費;
  const maxPersonnelRows = personnelGroup.maxDataRows;

  for (let i = 0; i < Math.min(sortedPersonnel.length, maxPersonnelRows); i++) {
    const currentRow = personnelGroup.dataStartRow + i;
    const item = sortedPersonnel[i];

    // マッピングシートで指定された行・列を優先
    writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols, item.targetRow, item.targetColumn);

    transferLog.push({
      tableNumber: '②',
      group: '人件費',
      itemName: item.itemName,
      row: item.targetRow || currentRow,  // 実際に転記された行
      status: '成功',
      amounts: item.amounts
    });
  }

  // 人件費が行数を超える場合は「その他人件費」として合計（11行に固定）
  if (sortedPersonnel.length > maxPersonnelRows) {
    const excessPersonnel = sortedPersonnel.slice(maxPersonnelRows);
    const excessAmounts = aggregateAmounts(excessPersonnel);
    const excessRow = personnelGroup.otherRow;  // 11行に固定

    writeExcelRowData(sheet, tableConfig.itemColumn, excessRow, 'その他人件費', excessAmounts, periodCols);

    transferLog.push({
      tableNumber: '②',
      group: '人件費',
      itemName: 'その他人件費',
      row: excessRow,
      status: '成功（集計）',
      amounts: excessAmounts,
      aggregatedCount: excessPersonnel.length
    });

    Logger.log(`②-その他人件費: ${excessPersonnel.length}件を集計 (行${excessRow})`);
  }

  // その他経費グループの転記
  const sortedOther = sortByCurrentPeriod(otherItems);
  const otherGroup = groups.その他経費;
  const maxOtherRows = otherGroup.maxDataRows - 1;  // 最後の行は「その他」用に確保

  for (let i = 0; i < Math.min(sortedOther.length, maxOtherRows); i++) {
    const currentRow = otherGroup.dataStartRow + i;
    const item = sortedOther[i];

    // マッピングシートで指定された行・列を優先
    writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols, item.targetRow, item.targetColumn);

    transferLog.push({
      tableNumber: '②',
      group: 'その他経費',
      itemName: item.itemName,
      row: item.targetRow || currentRow,  // 実際に転記された行
      status: '成功',
      amounts: item.amounts
    });
  }

  // その他経費が行数を超える場合は「その他」として合計（29行に固定）
  if (sortedOther.length > maxOtherRows) {
    const excessOther = sortedOther.slice(maxOtherRows);
    const excessAmounts = aggregateAmounts(excessOther);
    const otherRow = 29;  // 29行に固定

    writeExcelRowData(sheet, tableConfig.itemColumn, otherRow, 'その他', excessAmounts, periodCols);

    transferLog.push({
      tableNumber: '②',
      group: 'その他経費',
      itemName: 'その他',
      row: otherRow,
      status: '成功（集計）',
      amounts: excessAmounts,
      aggregatedCount: excessOther.length
    });

    Logger.log(`②-その他: ${excessOther.length}件を集計 (行${otherRow})`);
  }

  Logger.log(`②販売費及び一般管理費比較表: 人件費${sortedPersonnel.length}件、その他経費${sortedOther.length}件を転記完了`);

  return transferLog;
}

/**
 * ⑤ その他損益比較表への転記
 * カテゴリ別転記型：営業外収益/費用、特別利益/損失で分類して転記
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Object} ocrItemsByCategory - {{営業外収益: [...], 営業外費用: [...], ...}}
 */
function transferToTable05_OtherPL(sheet, ocrItemsByCategory) {
  const tableConfig = CONFIG.EXCEL_TRANSFER_CONFIG.TABLES['⑤その他損益比較表'];
  const periodCols = CONFIG.EXCEL_TRANSFER_CONFIG.PERIOD_COLUMNS_EXTENDED;

  Logger.log(`【⑤その他損益比較表】への転記を開始します`);

  const transferLog = [];

  // CONFIG から カテゴリマッピングを取得
  const categories = tableConfig.categories || {};

  Object.keys(ocrItemsByCategory).forEach(category => {
    const items = ocrItemsByCategory[category];
    if (items && items.length > 0) {
      const totalAmounts = aggregateAmounts(items);
      const categoryConfig = categories[category];
      const row = categoryConfig ? categoryConfig.row : null;

      if (row) {
        writeExcelRowData(sheet, tableConfig.itemColumn, row, category, totalAmounts, periodCols);

        transferLog.push({
          tableNumber: '⑤',
          itemName: category,
          row: row,
          status: '成功（集計）',
          amounts: totalAmounts,
          aggregatedCount: items.length
        });

        Logger.log(`⑤-${category}: ${items.length}件を集計 (当期: ${totalAmounts['当期']})`);
      } else {
        Logger.log(`警告: カテゴリ「${category}」の行番号が見つかりません`);
      }
    }
  });

  Logger.log(`⑤その他損益比較表: カテゴリ別集計を完了`);

  return transferLog;
}

/**
 * 指定されたテーブル範囲のデータをクリアする
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Object} tableConfig - クリア対象のテーブル設定
 */
function clearExcelTable(sheet, tableConfig) {
  const periodCols = tableConfig.columnLayout === 'extended'
    ? CONFIG.EXCEL_TRANSFER_CONFIG.PERIOD_COLUMNS_EXTENDED
    : CONFIG.EXCEL_TRANSFER_CONFIG.PERIOD_COLUMNS;

  const itemCol = tableConfig.itemColumn;
  const startRow = tableConfig.dataStartRow;
  const endRow = tableConfig.dataEndRow;

  // 項目名の範囲をクリア
  sheet.getRange(`${itemCol}${startRow}:${itemCol}${endRow}`).clearContent();

  // 金額の範囲をクリア
  Object.values(periodCols).forEach(cols => {
    const amountCol = cols.金額;
    if (amountCol) {
      sheet.getRange(`${amountCol}${startRow}:${amountCol}${endRow}`).clearContent();
    }
  });
  Logger.log(`クリア完了: ${tableConfig.tableName} のデータ範囲`);
}


/**
 * BS（貸借対照表）への転記
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - BSシート
 * @param {Object} classifiedBsData - 分類済みのBSデータ
 * @param {Object} bsConfig - BSシートの設定
 */
function transferToBSTable(sheet, classifiedBsData, bsConfig) {
  const transferLog = [];
  const itemColumn = bsConfig.ITEM_COLUMN;
  const periodColumns = bsConfig.PERIOD_COLUMNS;

  Object.keys(classifiedBsData).forEach(groupKey => {
    const items = classifiedBsData[groupKey];
    if (items.length === 0) return;

    const groupName = groupKey.replace('BS:', ''); // "BS:流動資産" -> "流動資産"
    const groupConfig = bsConfig.GROUPS[groupName];
    if (!groupConfig) {
      Logger.log(`警告: BSグループ設定が見つかりません: ${groupName}`);
      return;
    }

    Logger.log(`【BS:${groupName}】への転記を開始します（${items.length}件）`);

    // データを当期金額で降順ソート
    const sortedItems = sortByCurrentPeriod(items);
    const maxRows = groupConfig.dataEndRow - groupConfig.dataStartRow + 1;

    // 個別項目を転記
    const itemsToTransfer = sortedItems.slice(0, maxRows);
    itemsToTransfer.forEach((item, i) => {
      const row = groupConfig.dataStartRow + i;
      // マッピングシートで指定された行・列を優先
      writeExcelRowData(sheet, itemColumn, row, item.itemName, item.amounts, periodColumns, item.targetRow, item.targetColumn);
      transferLog.push({
        tableNumber: 'BS',
        group: groupName,
        itemName: item.itemName,
        row: item.targetRow || row,
        status: '成功',
        amounts: item.amounts
      });
    });

    // 行数を超えた項目を「その他」として集計
    if (sortedItems.length > maxRows) {
      const excessItems = sortedItems.slice(maxRows);
      const otherAmounts = aggregateAmounts(excessItems);
      const otherRow = groupConfig.dataEndRow; // 最後の行を「その他」とする

      writeExcelRowData(sheet, itemColumn, otherRow, `その他${groupName}`, otherAmounts, periodColumns);
      transferLog.push({
        tableNumber: 'BS',
        group: groupName,
        itemName: `その他${groupName}`,
        row: otherRow,
        status: '成功（集計）',
        amounts: otherAmounts,
        aggregatedCount: excessItems.length
      });
      Logger.log(`  その他${groupName}: ${excessItems.length}件を集計`);
    }
  });

  return transferLog;
}

/**
 * マッピングシートから分類情報を読み込み、該当するOCR項目を整理
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @param {Array<Array>} ocrData - OCRデータ [[項目名, 金額], ...]
 * @return {Object} {{tableKey: [{itemName, amounts, mappedTable, mappedGroup}, ...]}, ...}
 */
function classifyOcrDataByMapping(ss, ocrData) {
  const classifiedData = {
    // PL
    '①売上高内訳表': [],
    '②販売費及び一般管理費比較表': [],
    '③変動費内訳比較表': [],
    '④製造経費比較表': [],
    '⑤その他損益比較表': {
      '営業外収益': [], '営業外費用': [], '特別利益': [], '特別損失': []
    },
    // BS
    'BS:流動資産': [],
    'BS:固定資産': [],
    'BS:流動負債': [],
    'BS:固定負債': [],
    'BS:純資産': []
  };

  const mappingSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET_NAME);
  if (!mappingSheet || mappingSheet.getLastRow() <= 1) {
    Logger.log(`警告: マッピングシート「${CONFIG.MAPPING_SHEET_NAME}」に有効なデータがありません`);
    return classifiedData;
  }

  const mappingRange = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, 5).getValues();  // 5列に拡張（D列・E列追加）
  const mappingData = {};
  const plMappingKeys = [];  // デバッグ用: PL項目のマッピング一覧

  mappingRange.forEach(row => {
    if (row[0]) {
      const normalizedKey = normalizeJapaneseText(row[0], true);  // 勘定科目名: 強い正規化
      let normalizedTable = normalizeJapaneseText(row[1] || '', false);  // テーブル名: 軽い正規化

      // テーブル名の正規化: 通常の数字を丸数字に変換
      // 例: "1売上高" → "①売上高"
      normalizedTable = normalizedTable
        .replace(/^1売上高/, '①売上高')
        .replace(/^2販売費/, '②販売費')
        .replace(/^3変動費/, '③変動費')
        .replace(/^4製造経費/, '④製造経費')
        .replace(/^5その他/, '⑤その他');

      // D列・E列から転記先行・列を取得（任意）
      const targetRow = row[3] ? parseInt(row[3]) : null;  // D列: 転記先行
      const targetColumn = row[4] ? String(row[4]).trim().toUpperCase() : null;  // E列: 転記先列

      mappingData[normalizedKey] = {
        table: normalizedTable,
        group: normalizeJapaneseText(row[2] || '', false),  // グループ名: 軽い正規化
        targetRow: targetRow,      // 転記先行（null = 自動）
        targetColumn: targetColumn  // 転記先列（null = 自動）
      };

      // PL項目のマッピングを記録（BS以外）
      if (normalizedTable && !normalizedTable.startsWith('BS:')) {
        plMappingKeys.push(`"${normalizedKey}" → ${normalizedTable}`);
      }
    }
  });

  // デバッグ: マッピングシート内のPL項目を確認
  Logger.log(`
========== マッピングシート内のPL項目数: ${plMappingKeys.length}件 ==========`);
  Logger.log(`先頭10件:`);
  plMappingKeys.slice(0, 10).forEach(mapping => Logger.log(`  ${mapping}`));

  const unmappedPlItems = [];  // 未マッピングのPL項目リスト
  const ocrPlItems = [];  // デバッグ用: OCRシートから読み込んだPL候補項目
  ocrData.forEach(row => {
    // OCRデータの検証（4列形式: 項目名, 前々期, 前期, 当期）
    if (!row[0]) {
      return;  // 項目名が空の行はスキップ
    }

    const itemName = normalizeJapaneseText(row[0], true);  // 勘定科目名: 強い正規化
    
    // 3期分の金額を取得（4列形式）
    const amount2PeriodsAgo = row[1] !== undefined ? parseJapaneseNumber(row[1]) : 0; // 前々期
    const amount1PeriodAgo = row[2] !== undefined ? parseJapaneseNumber(row[2]) : 0;  // 前期
    const currentAmount = row[3] !== undefined ? parseJapaneseNumber(row[3]) : 0;     // 当期

    // 安全性チェック
    if (!isSafeText(itemName)) {
      Logger.log(`⚠ スキップ: ${itemName} (安全でない項目名)`);
      return;
    }

    // 集計行の除外（より正確な判定）
    // 除外対象: 「〇〇計」「合計」「小計」「総計」で終わる、または明確な集計行
    const isAggregationRow = 
      itemName.endsWith('計') || 
      itemName.endsWith('合計') || 
      itemName.endsWith('小計') || 
      itemName.endsWith('総計') ||
      itemName === '合計' ||
      itemName === '小計' ||
      itemName === '総計' ||
      itemName.includes('の部合計') ||
      itemName.includes('及び純資産合計');

    if (isAggregationRow) {
      Logger.log(`⚠ スキップ: ${itemName} (集計行)`);
      return;
    }

    const itemData = {
      itemName: itemName,
      amounts: { 
        '前々期': amount2PeriodsAgo, 
        '前期': amount1PeriodAgo, 
        '当期': currentAmount 
      }
    };

    Logger.log(`📊 OCRデータ読込: ${itemName} (前々期:${amount2PeriodsAgo}, 前期:${amount1PeriodAgo}, 当期:${currentAmount})`);

    const userMapping = mappingData[itemName];
    if (!userMapping || !userMapping.table) {
      // BS項目でない場合はPL未マッピング項目として記録
      unmappedPlItems.push(itemName);
      Logger.log(`⚠ 未マッピング（または転記先が空）: "${itemName}"`);
      return;
    }

    // OCRから読み込んだPL項目を記録（デバッグ用）
    if (!userMapping.table.startsWith('BS:')) {
      ocrPlItems.push(`"${itemName}" → ${userMapping.table}`);
    }

    const { table, group, targetRow, targetColumn } = userMapping;

    // itemDataに転記先行・列情報を追加
    const enrichedItemData = {
      ...itemData,
      mappedTable: table,
      mappedGroup: group,
      targetRow: targetRow,        // マッピングシートで指定された行（null = 自動）
      targetColumn: targetColumn    // マッピングシートで指定された列（null = 自動）
    };

    if (table.startsWith('BS:')) {
      if (classifiedData[table]) {
        classifiedData[table].push(enrichedItemData);
        Logger.log(`✓ BSマッピング適用: ${itemName} → ${table} [${group}]`);
      }
    } else if (table.startsWith('⑤')) {
      const category = group || '営業外収益';
      if (classifiedData[table] && classifiedData[table][category]) {
        const enrichedItemData_withCategory = { ...enrichedItemData, mappedGroup: category };
        classifiedData[table][category].push(enrichedItemData_withCategory);
        Logger.log(`✓ PLマッピング適用: ${itemName} → ${table} [${category}]`);
      }
    } else if (classifiedData[table]) {
      classifiedData[table].push(enrichedItemData);
      Logger.log(`✓ PLマッピング適用: ${itemName} → ${table} [${group}]`);
    }
  });

  // ========== デバッグサマリー: PL項目のマッピング結果 ========== 
  Logger.log(`
========== PL項目マッピング結果サマリー ==========`);
  Logger.log(`OCRから読み込んだPL項目（マッピング成功）: ${ocrPlItems.length}件`);
  Logger.log(`先頭10件:`);
  ocrPlItems.slice(0, 10).forEach(item => Logger.log(`  ${item}`));

  Logger.log(`
未マッピングのPL候補項目: ${unmappedPlItems.length}件`);
  Logger.log(`先頭10件:`);
  unmappedPlItems.slice(0, 10).forEach(item => Logger.log(`  "${item}"`));
  Logger.log(`
【PLテーブル分類結果】`);
  Logger.log(`①売上高内訳表: ${classifiedData['①売上高内訳表'].length}件`);
  Logger.log(`②販売費及び一般管理費比較表: ${classifiedData['②販売費及び一般管理費比較表'].length}件`);
  Logger.log(`③変動費内訳比較表: ${classifiedData['③変動費内訳比較表'].length}件`);
  Logger.log(`④製造経費比較表: ${classifiedData['④製造経費比較表'].length}件`);
  Logger.log(`⑤その他損益比較表:
  営業外収益: ${classifiedData['⑤その他損益比較表']['営業外収益'].length}件
  営業外費用: ${classifiedData['⑤その他損益比較表']['営業外費用'].length}件
  特別利益: ${classifiedData['⑤その他損益比較表']['特別利益'].length}件
  特別損失: ${classifiedData['⑤その他損益比較表']['特別損失'].length}件`);
  const plTotal = classifiedData['①売上高内訳表'].length + classifiedData['②販売費及び一般管理費比較表'].length + classifiedData['③変動費内訳比較表'].length + classifiedData['④製造経費比較表'].length + classifiedData['⑤その他損益比較表']['営業外収益'].length + classifiedData['⑤その他損益比較表']['営業外費用'].length + classifiedData['⑤その他損益比較表']['特別利益'].length + classifiedData['⑤その他損益比較表']['特別損失'].length;
  Logger.log(`PL小計: ${plTotal}件
`);

  Logger.log(`【BSテーブル分類結果】`);
  Logger.log(`BS:流動資産: ${classifiedData['BS:流動資産'].length}件`);
  Logger.log(`BS:固定資産: ${classifiedData['BS:固定資産'].length}件`);
  Logger.log(`BS:流動負債: ${classifiedData['BS:流動負債'].length}件`);
  Logger.log(`BS:固定負債: ${classifiedData['BS:固定負債'].length}件`);
  Logger.log(`BS:純資産: ${classifiedData['BS:純資産'].length}件`);
  const bsTotal = classifiedData['BS:流動資産'].length + classifiedData['BS:固定資産'].length + classifiedData['BS:流動負債'].length + classifiedData['BS:固定負債'].length + classifiedData['BS:純資産'].length;
  Logger.log(`BS小計: ${bsTotal}件
`);

  Logger.log(`合計: ${(plTotal + bsTotal)}件`);
  Logger.log(`==================================================
`);

  return {
    classifiedData: classifiedData,
    debugInfo: {
      plMappingCount: plMappingKeys.length,
      plMappingSample: plMappingKeys.slice(0, 10),
      ocrPlCount: ocrPlItems.length,
      ocrPlSample: ocrPlItems.slice(0, 10),
      unmappedCount: unmappedPlItems.length,
      unmappedSample: unmappedPlItems.slice(0, 10)
    }
  };
}

/**
 * =================================================================================
 * 統合処理：マッピングシートのデータをエクセルに一括転記する
 * =================================================================================
 */
function startStep3_transferDataToExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ocrSheet = ss.getSheetByName(CONFIG.OCR_SHEET_NAME);
  if (!ocrSheet || ocrSheet.getLastRow() === 0) {
    SpreadsheetApp.getUi().alert(`「${CONFIG.OCR_SHEET_NAME}」にデータがありません。ステップ１から実行してください。`);
    return;
  }

  const plSheet = ss.getSheetByName(CONFIG.EXCEL_TRANSFER_CONFIG.SHEET_NAME);
  const bsSheet = ss.getSheetByName(CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET.SHEET_NAME);
  if (!plSheet || !bsSheet) {
    SpreadsheetApp.getUi().alert('転記先シートが見つかりません。');
    return;
  }

  SpreadsheetApp.getUi().alert('マッピングに基づき転記を開始します。');

  // 1. 転記前のデータクリア
  Object.values(CONFIG.EXCEL_TRANSFER_CONFIG.TABLES).forEach(tableConfig => {
    clearExcelTable(plSheet, tableConfig);
  });
  Object.values(CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET.GROUPS).forEach(groupConfig => {
     const bsTableConfig = {
        tableName: groupConfig.label,
        itemColumn: CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET.ITEM_COLUMN,
        dataStartRow: groupConfig.dataStartRow,
        dataEndRow: groupConfig.dataEndRow,
        columnLayout: 'bs'
     };
     // BSの金額列は特殊なため個別指定でクリア
     const colsToClear = CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET.PERIOD_COLUMNS;
     plSheet.getRange(`${bsTableConfig.itemColumn}${bsTableConfig.dataStartRow}:${bsTableConfig.itemColumn}${bsTableConfig.dataEndRow}`).clearContent();
     Object.values(colsToClear).forEach(cols => {
        bsSheet.getRange(`${cols.金額}${bsTableConfig.dataStartRow}:${cols.金額}${bsTableConfig.dataEndRow}`).clearContent();
     });
  });


  // 2. OCRデータとマッピング情報を読み込み、分類
  const ocrData = ocrSheet.getRange(2, 1, ocrSheet.getLastRow() - 1, 4).getValues();  // 4列（項目名、前々期、前期、当期）+ ヘッダー行をスキップ
  const result = classifyOcrDataByMapping(ss, ocrData);
  const classifiedData = result.classifiedData;
  const debugInfo = result.debugInfo;

  // ========== デバッグログ: 分類結果の確認 ========== 
  let debugMessage = '========== 分類結果サマリー ==========\n\n';

  // マッピングシート情報
  debugMessage += '【マッピングシート】\n';
  debugMessage += 'PL項目マッピング数: ' + debugInfo.plMappingCount + '件\n';
  if (debugInfo.plMappingSample.length > 0) {
    debugMessage += '先頭5件:\n';
    debugInfo.plMappingSample.slice(0, 5).forEach(item => {
      debugMessage += '  ' + item + '\n';
    });
  }
  debugMessage += '\n';

  // OCR PL項目情報
  debugMessage += '【OCRシートのPL項目（マッピング成功）】\n';
  debugMessage += 'マッピング成功: ' + debugInfo.ocrPlCount + '件\n';
  if (debugInfo.ocrPlSample.length > 0) {
    debugMessage += '先頭5件:\n';
    debugInfo.ocrPlSample.slice(0, 5).forEach(item => {
      debugMessage += '  ' + item + '\n';
    });
  }
  debugMessage += '\n';

  // 未マッピング項目
  debugMessage += '【未マッピング項目】\n';
  debugMessage += '未マッピング: ' + debugInfo.unmappedCount + '件\n';
  if (debugInfo.unmappedSample.length > 0) {
    debugMessage += '先頭10件:\n';
    debugInfo.unmappedSample.forEach(item => {
      debugMessage += '  "' + item + '"\n';
    });
  }
  debugMessage += '\n';

  debugMessage += '【PLテーブル分類結果】\n';
  debugMessage += '①売上高内訳表: ' + classifiedData['①売上高内訳表'].length + '件\n';
  debugMessage += '②販売費及び一般管理費比較表: ' + classifiedData['②販売費及び一般管理費比較表'].length + '件\n';
  debugMessage += '③変動費内訳比較表: ' + classifiedData['③変動費内訳比較表'].length + '件\n';
  debugMessage += '④製造経費比較表: ' + classifiedData['④製造経費比較表'].length + '件\n';
  debugMessage += '⑤その他損益比較表:\n';
  debugMessage += '  営業外収益: ' + classifiedData['⑤その他損益比較表']['営業外収益'].length + '件\n';
  debugMessage += '  営業外費用: ' + classifiedData['⑤その他損益比較表']['営業外費用'].length + '件\n';
  debugMessage += '  特別利益: ' + classifiedData['⑤その他損益比較表']['特別利益'].length + '件\n';
  debugMessage += '  特別損失: ' + classifiedData['⑤その他損益比較表']['特別損失'].length + '件\n';
  const plTotal = classifiedData['①売上高内訳表'].length + classifiedData['②販売費及び一般管理費比較表'].length + classifiedData['③変動費内訳比較表'].length + classifiedData['④製造経費比較表'].length + classifiedData['⑤その他損益比較表']['営業外収益'].length + classifiedData['⑤その他損益比較表']['営業外費用'].length + classifiedData['⑤その他損益比較表']['特別利益'].length + classifiedData['⑤その他損益比較表']['特別損失'].length;
  debugMessage += 'PL小計: ' + plTotal + '件\n\n';

  debugMessage += '【BSテーブル分類結果】\n';
  debugMessage += 'BS:流動資産: ' + classifiedData['BS:流動資産'].length + '件\n';
  debugMessage += 'BS:固定資産: ' + classifiedData['BS:固定資産'].length + '件\n';
  debugMessage += 'BS:流動負債: ' + classifiedData['BS:流動負債'].length + '件\n';
  debugMessage += 'BS:固定負債: ' + classifiedData['BS:固定負債'].length + '件\n';
  debugMessage += 'BS:純資産: ' + classifiedData['BS:純資産'].length + '件\n';
  const bsTotal = classifiedData['BS:流動資産'].length + classifiedData['BS:固定資産'].length + classifiedData['BS:流動負債'].length + classifiedData['BS:固定負債'].length + classifiedData['BS:純資産'].length;
  debugMessage += 'BS小計: ' + bsTotal + '件\n\n';

  debugMessage += '合計: ' + (plTotal + bsTotal) + '件\n';
  debugMessage += '========================================';

  Logger.log(debugMessage);
  SpreadsheetApp.getUi().alert('🔍 分類結果を確認してください', debugMessage, SpreadsheetApp.getUi().ButtonSet.OK);
  // ====================================================

  let allLogs = [];

  // 3. 各PLテーブルへ転記
  allLogs.push(...transferToTable01_SalesBreakdown(plSheet, classifiedData['①売上高内訳表']));
  allLogs.push(...transferToTable02_SGA(plSheet, classifiedData['②販売費及び一般管理費比較表']));
  allLogs.push(...transferToTable03_VariableCosts(plSheet, classifiedData['③変動費内訳比較表']));
  allLogs.push(...transferToTable04_ManufacturingExpenses(plSheet, classifiedData['④製造経費比較表']));
  allLogs.push(...transferToTable05_OtherPL(plSheet, classifiedData['⑤その他損益比較表']));

  // 4. BSテーブルへ転記
  const bsData = {
    'BS:流動資産': classifiedData['BS:流動資産'],
    'BS:固定資産': classifiedData['BS:固定資産'],
    'BS:流動負債': classifiedData['BS:流動負債'],
    'BS:固定負債': classifiedData['BS:固定負債'],
    'BS:純資産': classifiedData['BS:純資産']
  };
  allLogs.push(...transferToBSTable(bsSheet, bsData, CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET));


  // 5. ログを記録
  const logEntries = [['タイムスタンプ', '分類', '項目名', '転記先行', '金額（当期）', 'ステータス', '詳細']];
  const timestamp = new Date();
  allLogs.forEach(log => {
    logEntries.push([
      timestamp,
      log.tableNumber === 'BS' ? `BS:${log.group}` : `${log.tableNumber}:${log.group || ''}`,
      log.itemName,
      log.row,
      log.amounts ? log.amounts['当期'] : '',
      log.status,
      log.aggregatedCount ? `${log.aggregatedCount}件を集計` : ''
    ]);
  });
  writeReconciliationLog(ss, logEntries);

  SpreadsheetApp.getUi().alert(`転記が完了しました。\n\n${allLogs.length}件の処理を行いました。\n詳細は「${CONFIG.RECONCILIATION_LOG_SHEET}」シートをご確認ください。`);
}

/**
 * 転記ログシートを表示（メニュー用）
 */
function showTransferLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.RECONCILIATION_LOG_SHEET);

  if (!logSheet) {
    SpreadsheetApp.getUi().alert(`「${CONFIG.RECONCILIATION_LOG_SHEET}」シートが見つかりません。\n\nまずステップ３を実行してください。`);
    return;
  }

  SpreadsheetApp.setActiveSheet(logSheet);
  SpreadsheetApp.getUi().alert(`「${CONFIG.RECONCILIATION_LOG_SHEET}」シートを表示しました。`);
}

/**
 * デバッグ：基本的なログ出力確認
 */
function testBasicLogging() {
  Logger.log(`テスト開始`);
  Logger.log(`テスト: startStep3テスト開始`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`テスト: スプレッドシート取得成功`);
    Logger.log(`スプレッドシート名: ${ss.getName()}`);

    // シート一覧を取得
    const sheets = ss.getSheets();
    Logger.log(`シート数: ${sheets.length}`);
    sheets.forEach((sheet, idx) => {
      Logger.log(`  [${idx}] ${sheet.getName()}`);
    });

    const ocrSheet = ss.getSheetByName(CONFIG.OCR_SHEET_NAME);
    Logger.log(`テスト: OCRシート取得試行 - ${CONFIG.OCR_SHEET_NAME}`);
    if (ocrSheet) {
      Logger.log(`テスト: OCRシート取得成功 - 最終行: ${ocrSheet.getLastRow()}`);
    } else {
      Logger.log(`テスト: OCRシート取得失敗`);
    }

    Logger.log(`テスト: テスト終了`);
  } catch (e) {
    Logger.log(`テストエラー: ${e.message}`);
  }
}
