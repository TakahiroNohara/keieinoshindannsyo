/**
 * =================================================================================
 * グローバル設定
 * =================================================================================
 */

/** @typedef {{item: string, value: string}} MappingEntry */
/** @typedef {{name: string, id: string}} FileInfo */
/** @typedef {{name: string, id: string}} FolderInfo */

const VERSION = '3.3.0'; // スクリプトバージョン管理（v3.3.0: PL/BSマッピング統合）

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
        dataEndRow: 19,
        subtotalRow: 17,
        totalRow: 20,
        itemColumn: 'A',
        dataRowCount: 6,
        type: 'classified',  // グループ分類型に変更
        groups: Object.freeze({
          '変動費': {
            dataStartRow: 14,
            dataEndRow: 15,
            maxDataRows: 2,
            label: '変動費'
          },
          '期首棚卸高': {
            dataStartRow: 16,
            dataEndRow: 16,
            maxDataRows: 1,
            label: '期首棚卸高'
          },
          '期末棚卸高': {
            dataStartRow: 17,
            dataEndRow: 19,
            maxDataRows: 3,
            otherRow: 19,
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

  // ========================================================
  // 旧設定（互換性維持）- シート名を統一
  // ========================================================
  TARGET_SHEETS: {
    PL: {
      NAME: '４．３期比較表',
      ITEM_RANGES: ['B6:B7', 'B11:B16', 'B20:B22', 'B24:B25', 'B28:B32', 'B35:B37', 'B41:B44', 'B47:B54', 'B58:B62'],
      COORDINATES: Object.freeze({
        '売上高': 'C6', '売上原価': 'C7', '売上総利益': 'C8',
        '②役員報酬': 'C11', '③給料手当': 'C12', '④福利厚生費': 'C13', '⑤その他経費': 'C14', '⑥減価償却費': 'C15', '⑦支払家賃': 'C16', '販売費及び一般管理費 合計': 'C17',
        '営業利益': 'C18',
        '⑧受取利息': 'C20', '⑨受取配当金': 'C21', '⑩その他': 'C22', '営業外収益 合計': 'C23',
        '⑪支払利息': 'C24', '⑫その他': 'C25', '営業外費用 合計': 'C26',
        '経常利益': 'C27',
        '特別利益 合計': 'C32', '特別損失 合計': 'C38',
        '税引前当期純利益': 'C39', '法人税等': 'C44', '当期純利益': 'C45'
      }),
      FORMULAS_R1C1: Object.freeze({
        'C17': '=IFERROR(SUM(R[-6]C:R[-1]C), "")',
        'C23': '=IFERROR(SUM(R[-3]C:R[-1]C), "")',
        'C26': '=IFERROR(SUM(R[-2]C:R[-1]C), "")'
      })
    },
    BS: {
      NAME: '４．３期比較表',
      ITEM_RANGES: ['B7:B12', 'B20:B26', 'B31:B31', 'B37:B41', 'B48:B49', 'B55:B58'],
      COORDINATES: Object.freeze({
        '現預金': 'C7', '受取手形': 'C8', '売掛金': 'C9', '商品': 'C10', 'その他流動資産': 'C11', '貸倒引当金': 'C12', '流動資産 合計': 'C17',
        '建物': 'C20', '構築物': 'C21', '機械装置': 'C22', '車輛': 'C23', '工具器具備品': 'C24', '土地': 'C25', 'その他': 'C26', '有形固定資産 合計': 'C28',
        '投資その他': 'C31', '固定資産 合計': 'C34', '資産 合計': 'C35',
        '支払手形': 'C37', '買掛金': 'C38', '短期借入金': 'C39', '未払金': 'C40', 'その他流動負債': 'C41', '流動負債 合計': 'C45',
        '長期借入金': 'C48', 'その他': 'C49', '固定負債 合計': 'C52', '負債 合計': 'C53',
        '資本金': 'C55', '資本準備金': 'C56', '利益準備金': 'C57', 'その他利益剰余金': 'C58', '純資産 合計': 'C59',
        '負債純資産 合計': 'C60'
      }),
      FORMULAS_R1C1: Object.freeze({
        'C17': '=IFERROR(SUM(R[-10]C:R[-5]C), "")',
        'C28': '=IFERROR(SUM(R[-8]C:R[-2]C), "")',
        'C34': '=IFERROR(R[-6]C+R[-3]C, "")',
        'C35': '=IFERROR(R[-18]C+R[-1]C, "")',
        'C45': '=IFERROR(SUM(R[-8]C:R[-4]C), "")',
        'C52': '=IFERROR(SUM(R[-4]C:R[-3]C), "")',
        'C53': '=IFERROR(R[-8]C+R[-1]C, "")',
        'C59': '=IFERROR(SUM(R[-4]C:R[-1]C), "")',
        'C60': '=IFERROR(R[-7]C+R[-1]C, "")'
      }),
      BALANCE_CHECK: Object.freeze({
        'E35': '=IFERROR(IF(ABS(C35-(C53+C59))<1, "✓", "⚠差異: "&TEXT(C35-(C53+C59), "#,##0")), "")'
      })
    }
  },
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
    .addItem('【ステップ３】マッピングを適用して転記 (旧)', 'startStep3_transferData')
    .addItem('【ステップ３NEW】自動分類して各表に転記（推奨）', 'startStep3_transferDataToExcel')
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
      ocrSheet.getRange(1, 1, results.length, 2).setValues(results).setHorizontalAlignment('left').setNumberFormat('@');
      ocrSheet.setColumnWidths(1, 2, 250);
      SpreadsheetApp.getUi().alert('OCR抽出が完了し、「' + CONFIG.OCR_SHEET_NAME + '」に出力しました。\n\n次に「ステップ２」を実行してください。');
    } else {
      SpreadsheetApp.getUi().alert('有効なデータを抽出できませんでした。');
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + e.message);
  }
}

/**
 * =================================================================================
 * 【ステップ２】統合マッピングシートの準備
 * =================================================================================
 */
function startStep2_createUnifiedMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ocrSheet = ss.getSheetByName(CONFIG.OCR_SHEET_NAME);
  if (!ocrSheet || ocrSheet.getLastRow() === 0) {
    SpreadsheetApp.getUi().alert('先に「ステップ１」を実行して、「' + CONFIG.OCR_SHEET_NAME + '」にデータを抽出してください。');
    return;
  }
  const sourceItems = ocrSheet.getRange(1, 1, ocrSheet.getLastRow(), 1).getValues().flat().filter(String);

  // 転記先シートの確認
  const targetSheetPL = ss.getSheetByName(CONFIG.EXCEL_TRANSFER_CONFIG.SHEET_NAME);
  const targetSheetBS = ss.getSheetByName(CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET.SHEET_NAME);
  if (!targetSheetPL || !targetSheetBS) {
    SpreadsheetApp.getUi().alert('転記先のシート「' + CONFIG.EXCEL_TRANSFER_CONFIG.SHEET_NAME + '」または「' + CONFIG.EXCEL_TRANSFER_CONFIG.BS_SHEET.SHEET_NAME + '」が見つかりません。');
    return;
  }

  // マッピングシートを作成/クリア
  const mappingSheet = getOrCreateSheet(CONFIG.MAPPING_SHEET_NAME);
  mappingSheet.clear();
  mappingSheet.setFrozenRows(1);

  // ヘッダー行を設定
  const headerValues = [
    ['OCR抽出項目', '転記先分類（推奨）', '詳細グループ（推奨）', '説明']
  ];
  mappingSheet.getRange('A1:D1').setValues(headerValues)
    .setFontWeight('bold')
    .setBackground('#4A86E8')
    .setFontColor('#FFFFFF');

  // データ行を準備（自動分類を適用）
  const dataRows = [];
  const descriptions = [];

  sourceItems.forEach(item => {
    const normalized = normalizeJapaneseText(item);
    const classification = classifyItemToTable(normalized);

    let suggestedTable = '';
    let suggestedGroup = '';
    let description = '自動推奨なし（手動で設定）';

    if (classification && classification.tableKey) {
      suggestedTable = classification.tableKey;
      suggestedGroup = classification.categoryKey || '';
      description = '自動推奨（必要に応じて変更）';
    }

    dataRows.push([item, suggestedTable, suggestedGroup]);
    descriptions.push([description]);
  });

  // データ行に値を一括設定
  if (dataRows.length > 0) {
    mappingSheet.getRange(2, 1, dataRows.length, 3).setValues(dataRows);
    mappingSheet.getRange(2, 4, descriptions.length, 1).setValues(descriptions).setFontColor('#999999');
  }

  // B列：転記先分類のドロップダウン（PLとBSを統合）
  const tableOptions = [
    // PL
    '①売上高内訳表', '②販売費及び一般管理費比較表', '③変動費内訳比較表', '④製造経費比較表', '⑤その他損益比較表',
    // BS
    'BS:流動資産', 'BS:固定資産', 'BS:流動負債', 'BS:固定負債', 'BS:純資産'
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
      '人件費', 'その他', 'その他経費', '変動費', '期首棚卸高', '期末棚卸高', '労務費',
      '営業外収益', '営業外費用', '特別利益', '特別損失',
      // BS（参考用、手入力も可）
      '現金・預金', '売上債権', '棚卸資産', 'その他流動資産',
      '有形固定資産', '無形固定資産', '投資その他',
      '仕入債務', '短期借入金', 'その他流動負債',
      '長期借入金', 'その他固定負債',
      '資本金', '資本剰余金', '利益剰余金'
    ], true)
    .setAllowInvalid(true)
    .build();
  mappingSheet.getRange(2, 3, sourceItems.length, 1).setDataValidation(groupRule);

  // 列幅を調整
  mappingSheet.setColumnWidths(1, 1, 250);
  mappingSheet.setColumnWidths(2, 1, 220);
  mappingSheet.setColumnWidths(3, 1, 180);
  mappingSheet.setColumnWidths(4, 1, 250);

  SpreadsheetApp.setActiveSheet(mappingSheet);

  // ヘルプメッセージ
  const helpMessage = `「${CONFIG.MAPPING_SHEET_NAME}」を準備しました。

【✨ 自動分類機能】
B列「転記先分類」とC列「詳細グループ」に推奨値を自動設定しました！
正しい場合はそのまま、違う場合のみ変更してください。

【記入方法】
1. B列「転記先分類」：ドロップダウンから転記先を選択
   • PL項目 → ①〜⑤の表
   • BS項目 → 「BS:流動資産」などのBSカテゴリ

2. C列「詳細グループ」：B列の分類に応じて選択
   • PLの販管費 → 「人件費」/「その他」
   • PLの変動費 → 「変動費」/「期首棚卸高」/「期末棚卸高」
   • PLの製造経費 → 「労務費」/「その他経費」
   • PLのその他損益 → 「営業外収益」/「営業外費用」など
   • BS項目 → C列は空欄でOK（B列の分類で転記されます）

【処理フロー】
- PL項目は指定された表のルールに従って転記されます。
- BS項目は指定された大区分（流動資産など）に集計・転記されます。

✅ 内容を確認後、「ステップ３NEW」を実行してください。`;

  SpreadsheetApp.getUi().alert(helpMessage);
}

/**
 * =================================================================================
 * 【ステップ３】マッピングを適用して転記
 * =================================================================================
 */
function startStep3_transferData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ocrSheet = ss.getSheetByName(CONFIG.OCR_SHEET_NAME);
  const mappingSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET_NAME);
  if (!ocrSheet || !mappingSheet) {
    SpreadsheetApp.getUi().alert('作業用シートが見つかりません。ステップ１と２を先に実行してください。');
    return;
  }

  // OCRデータを読み込み（正規化を適用）
  const ocrData = ocrSheet.getRange(1, 1, ocrSheet.getLastRow(), 2).getValues().reduce((obj, row) => {
    if (row[0]) {
      const normalizedKey = normalizeJapaneseText(row[0]);
      obj[normalizedKey] = row[1];
    }
    return obj;
  }, {});

  const mappingRules = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, 2).getValues().filter(row => row[0] && row[1]);
  const plSheet = ss.getSheetByName(CONFIG.TARGET_SHEETS.PL.NAME);
  const bsSheet = ss.getSheetByName(CONFIG.TARGET_SHEETS.BS.NAME);

  // 調整ログの準備
  const logEntries = [['タイムスタンプ', '転記元項目', '転記先項目', '金額', 'ステータス', 'メッセージ']];
  const timestamp = new Date();
  const unmappedItems = new Set(Object.keys(ocrData));
  const usedCoordinates = new Set();

  // データ転記
  mappingRules.forEach(rule => {
    const sourceItem = normalizeJapaneseText(rule[0]);
    const destinationItem = rule[1];
    const amount = ocrData[sourceItem];

    unmappedItems.delete(sourceItem); // マッピングされた項目を記録

    if (amount === undefined || amount === '') {
      logEntries.push([timestamp, sourceItem, destinationItem, '', 'スキップ', '金額が見つかりません']);
      return;
    }

    let targetCell;
    let sheetType;
    if (CONFIG.TARGET_SHEETS.PL.COORDINATES[destinationItem]) {
      targetCell = plSheet.getRange(CONFIG.TARGET_SHEETS.PL.COORDINATES[destinationItem]);
      sheetType = 'PL';
      usedCoordinates.add(CONFIG.TARGET_SHEETS.PL.COORDINATES[destinationItem]);
    } else if (CONFIG.TARGET_SHEETS.BS.COORDINATES[destinationItem]) {
      targetCell = bsSheet.getRange(CONFIG.TARGET_SHEETS.BS.COORDINATES[destinationItem]);
      sheetType = 'BS';
      usedCoordinates.add(CONFIG.TARGET_SHEETS.BS.COORDINATES[destinationItem]);
    }

    if (targetCell) {
      targetCell.setValue(amount);
      logEntries.push([timestamp, sourceItem, destinationItem, amount, '成功', `${sheetType}シートに転記`]);
      Logger.log(`転記成功: ${sourceItem} → ${destinationItem} (${amount})`);
    } else {
      logEntries.push([timestamp, sourceItem, destinationItem, amount, 'エラー', '転記先が見つかりません']);
      Logger.log(`エラー: 転記先が見つかりません: ${destinationItem}`);
    }
  });

  // 未マッピング項目のログ
  if (unmappedItems.size > 0) {
    Logger.log(`警告: ${unmappedItems.size}件の項目がマッピングされていません:`);
    unmappedItems.forEach(item => {
      Logger.log(`  - ${item}`);
      logEntries.push([timestamp, item, '', ocrData[item], '未マッピング', 'マッピング設定が必要です']);
    });
  }

  // R1C1形式の数式を設定（データ整合性向上のため）
  if (CONFIG.TARGET_SHEETS.PL.FORMULAS_R1C1) {
    Object.keys(CONFIG.TARGET_SHEETS.PL.FORMULAS_R1C1).forEach(cellAddress => {
      plSheet.getRange(cellAddress).setFormulaR1C1(CONFIG.TARGET_SHEETS.PL.FORMULAS_R1C1[cellAddress]);
    });
    Logger.log('PLシートに数式を設定しました');
  }
  if (CONFIG.TARGET_SHEETS.BS.FORMULAS_R1C1) {
    Object.keys(CONFIG.TARGET_SHEETS.BS.FORMULAS_R1C1).forEach(cellAddress => {
      bsSheet.getRange(cellAddress).setFormulaR1C1(CONFIG.TARGET_SHEETS.BS.FORMULAS_R1C1[cellAddress]);
    });
    Logger.log('BSシートに数式を設定しました');
  }

  // BSバランスチェックを設定
  if (CONFIG.TARGET_SHEETS.BS.BALANCE_CHECK) {
    Object.keys(CONFIG.TARGET_SHEETS.BS.BALANCE_CHECK).forEach(cellAddress => {
      bsSheet.getRange(cellAddress).setFormula(CONFIG.TARGET_SHEETS.BS.BALANCE_CHECK[cellAddress]);
    });
    Logger.log('BSシートにバランスチェックを設定しました');
  }

  // 数式セルの保護
  protectFormulaCells(plSheet, CONFIG.TARGET_SHEETS.PL.FORMULAS_R1C1);
  protectFormulaCells(bsSheet, CONFIG.TARGET_SHEETS.BS.FORMULAS_R1C1);
  protectFormulaCells(bsSheet, CONFIG.TARGET_SHEETS.BS.BALANCE_CHECK);

  // 調整ログシートに記録
  writeReconciliationLog(ss, logEntries);

  // バージョン情報を記録
  plSheet.getRange('A1').setNote(`スクリプトバージョン: ${CONFIG.VERSION}\n更新日時: ${timestamp}`);
  bsSheet.getRange('A1').setNote(`スクリプトバージョン: ${CONFIG.VERSION}\n更新日時: ${timestamp}`);

  SpreadsheetApp.getUi().alert(`転記が完了しました。\n\n転記成功: ${logEntries.length - 1 - unmappedItems.size}件\n未マッピング: ${unmappedItems.size}件\n\n詳細は「${CONFIG.RECONCILIATION_LOG_SHEET}」シートをご確認ください。`);
}

/**
 * =================================================================================
 * ヘルパー関数群（内部処理用）
 * =================================================================================
 */

/**
 * 日本語文字列を正規化（全角/半角、空白、Unicode正規化）
 * @param {string} text - 正規化する文字列
 * @return {string} 正規化された文字列
 */
function normalizeJapaneseText(text) {
  try {
    if (!text) return '';
    // Unicode正規化（NFKC: 互換文字を標準形に統一）
    let normalized = String(text).normalize('NFKC');
    // 前後の空白を削除
    normalized = normalized.trim();
    // 連続する空白を1つに統一
    normalized = normalized.replace(/\s+/g, ' ');
    return normalized;
  } catch (e) {
    Logger.log(`警告: テキスト正規化エラー: ${e.message}`);
    return String(text || '').trim();
  }
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
  if (/[ -]/.test(str)) {
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
      if (attempt < maxRetries - 1) {
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
    const prompt = `このファイルには、「貸借対照表」「損益計算書」「販売費及び一般管理費内訳書」が含まれています。これらの書類から、記載されているすべての勘定科目と金額のペアを抽出してください。出力する際は、必ず以下の順番を守ってください。1. 「貸借対照表」のすべての項目を上から順に 2. 「損益計算書」のすべての項目を上から順に 3. 「販売費及び一般管理費内訳書」のすべての項目を上から順に。小計、合計、内訳項目も含め、勘定科目と金額が記載されている行は例外なくすべてを抽出対象とします。出力形式は「勘定科目,金額」のCSV形式にしてください。金額に通貨記号(¥)や桁区切りのカンマ(,)は含めず、半角数字のみで出力してください。タイトル、日付、会社名など、勘定科目ではないテキスト行は無視してください。例:\n流動資産,50000000\n現金及び預金,20000000\n売上高,123456789`;
    const requestBody = {
      "contents": [{"parts": [{ "text": prompt }, { "inline_data": { "mime_type": file.getMimeType(), "data": Utilities.base64Encode(file.getBlob().getBytes()) } }] }], "generationConfig": { "temperature": 0.0, "topP": 1, "maxOutputTokens": 8192 }
    };
    const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(requestBody), 'muteHttpExceptions': true };
    const response = UrlFetchApp.fetch(API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    if (responseCode === 200) {
      const json = JSON.parse(responseBody);
      if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0].text) {
        return json.candidates[0].content.parts[0].text;
      } else { return ""; }
    } else { throw new Error(`APIリクエストに失敗しました。ステータスコード: ${responseCode}\nレスポンス: ${responseBody}`); }
  }, CONFIG.API.MAX_RETRIES, CONFIG.API.RETRY_DELAY_MS);
}

/**
 * CSV形式のOCR結果をパースし、正規化とバリデーションを適用
 * @param {string} text - CSV形式のテキスト
 * @return {Array<Array<string>>} [[項目名, 金額], ...] の配列
 */
function parseItemValueCsvResult(text) {
  const cleanedText = text.replace(/```csv\n?/g, '').replace(/```/g, '').trim();
  if (!cleanedText) return [];

  const lines = cleanedText.split('\n');
  const results = [];
  const seenItems = new Set(); // 重複チェック用

  for (const line of lines) {
    const parts = line.split(',');
    if (parts.length >= 2) {
      let item = parts[0].trim();
      let value = parts.slice(1).join('').trim();

      // 日本語テキストを正規化
      item = normalizeJapaneseText(item);

      // 安全性チェック（数式インジェクション防止）
      if (!isSafeText(item)) {
        Logger.log(`警告: 安全でない項目名を検出してスキップしました: ${item}`);
        continue;
      }

      // 数値を堅牢に解析
      const parsedValue = parseJapaneseNumber(value);

      // 重複項目のログ記録
      if (seenItems.has(item)) {
        Logger.log(`警告: 重複項目を検出しました: ${item} (既存の値を保持します)`);
        continue;
      }

      seenItems.add(item);
      results.push([item, parsedValue]);
    }
  }

  Logger.log(`OCR結果: ${results.length}件の項目を抽出しました`);
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
      const itemName = normalizeJapaneseText(row[0]);
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

  // 集計行を除外（「計」「合計」「小計」「総計」を含む項目は転記不要）
  if (normalized.match(/計$|合計|小計|総計|^計/)) {
    Logger.log(`集計行のためスキップ: ${itemName}`);
    return null;
  }

  // ===== 優先度A: BS項目（キーワードが明確なものから） =====
  // 純資産
  if (normalized.match(/資本金|準備金|剰余金|自己株式/)) {
    return { tableKey: 'BS:純資産', categoryKey: null };
  }
  // 固定負債
  if (normalized.match(/長期借入金|社債|退職給付引当金/)) {
    return { tableKey: 'BS:固定負債', categoryKey: null };
  }
  // 流動負債
  if (normalized.match(/支払手形|買掛金|短期借入金|未払金|未払費用|前受金|預り金|賞与引当金/)) {
    return { tableKey: 'BS:流動負債', categoryKey: null };
  }
  // 固定資産
  if (normalized.match(/建物|構築物|機械|装置|車両|運搬具|工具|器具|備品|土地|投資有価証券|長期貸付金|繰延資産/)) {
    return { tableKey: 'BS:固定資産', categoryKey: null };
  }
  // 流動資産
  if (normalized.match(/現金|預金|売掛金|受取手形|商品|製品|仕掛品|原材料|貯蔵品|前渡金|前払費用|未収入金|貸倒引当金/)) {
    return { tableKey: 'BS:流動資産', categoryKey: null };
  }


  // ===== 優先度B: PL項目（その他損益など、より具体的なものから）===== 
  // ⑤ その他損益
  if (normalized.match(/受取利息|預金利息|貸付金利息|受取配当金|配当金|有価証券利息|雑収入|受取手数料/)) {
    return { tableKey: '⑤その他損益比較表', categoryKey: '営業外収益' };
  }
  if (normalized.match(/支払利息|借入金利息|社債利息|手形売却損|雑損失/)) {
    return { tableKey: '⑤その他損益比較表', categoryKey: '営業外費用' };
  }
  if (normalized.match(/固定資産売却益|投資有価証券売却益|特別利益/)) {
    return { tableKey: '⑤その他損益比較表', categoryKey: '特別利益' };
  }
  if (normalized.match(/固定資産売却損|固定資産除却損|減損損失|特別損失/)) {
    return { tableKey: '⑤その他損益比較表', categoryKey: '特別損失' };
  }

  // ③ 変動費・原価関連
  if (normalized.match(/期首.*棚卸|期首.*在庫|期首商品/)) {
    return { tableKey: '③変動費内訳比較表', categoryKey: '期首棚卸高' };
  }
  if (normalized.match(/期末.*棚卸|期末.*在庫|期末商品/)) {
    return { tableKey: '③変動費内訳比較表', categoryKey: '期末棚卸高' };
  }
  if (normalized.match(/商品仕入|期中仕入|当期.*仕入|仕入高|材料仕入|原材料費|材料費|副資材|外注.*費|外注加工費|製品仕入|商品.*原価/)) {
    return { tableKey: '③変動費内訳比較表', categoryKey: '変動費' };
  }

  // ④ 製造経費
  if (normalized.match(/製造.*給料|製造.*賃金|製造.*賞与|工場.*給料|工場.*賃金|作業員.*給料/)) {
    return { tableKey: '④製造経費比較表', categoryKey: '労務費' };
  }
  if (normalized.match(/製造.*経費|工場.*経費|製造.*費|動力費|燃料費|工具.*費|機械.*費|設備.*費|作業.*費/) &&
      !normalized.match(/販売費|管理費/)) {
    return { tableKey: '④製造経費比較表', categoryKey: 'その他経費' };
  }

  // ② 販管費（人件費）
  if (normalized.match(/役員報酬|役員.*給料|給料.*手当|給与.*手当|賃金|賞与|ボーナス|雑給|法定福利費|福利厚生費|厚生費|退職金|退職給付/) &&
      !normalized.match(/製造|工場|作業/)) {
    return { tableKey: '②販売費及び一般管理費比較表', categoryKey: '人件費' };
  }

  // ② 販管費（その他経費）
  if (normalized.match(/支払家賃|地代家賃|賃借料|水道.*費|光熱費|電気代|ガス代|通信費|電話代|旅費.*交通費|交通費|旅費|出張.*費|広告.*費|宣伝費|販促費|荷造.*費|運賃|保険料|租税.*公課|公租公課|修繕費|消耗品費|事務.*費|会議費|交際費|接待.*費|寄付金|諸会費|雑費|減価償却費|リース料/) &&
      !normalized.match(/製造|工場/)) {
    return { tableKey: '②販売費及び一般管理費比較表', categoryKey: 'その他' };
  }

  // ① 売上高
  if (normalized.match(/売上高|商品.*売上|製品.*売上|売上|収益|販売.*収入/) &&
      !normalized.match(/売上原価|売上総利益/)) {
    return { tableKey: '①売上高内訳表', categoryKey: null };
  }

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
    itemName: normalizeJapaneseText(itemName),
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
 * @param {string} itemColumn - 項目列
 * @param {number} row - 行番号
 * @param {string} itemName - 項目名
 * @param {Object} amounts - {{前々期: number, 前期: number, 当期: number}}
 * @param {Object} periodColumns - 期別列設定
 */
function writeExcelRowData(sheet, itemColumn, row, itemName, amounts, periodColumns) {
  try {
    // 項目名を書き込み
    if (itemColumn && row) {
      sheet.getRange(`${itemColumn}${row}`).setValue(itemName);
    }

    // 各期の金額を書き込み
    Object.keys(periodColumns).forEach(period => {
      const colObj = periodColumns[period];
      const amountCol = colObj.金額;
      const amount = amounts[period] || 0;

      if (amount !== 0) {
        if (amountCol && row) {
          sheet.getRange(`${amountCol}${row}`).setValue(amount);
        }
      }
    });

    Logger.log(`✓ 書き込み成功: "${itemName}" → ${itemColumn}${row}`);
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
      writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols);
      transferLog.push({
        tableNumber: '①',
        itemName: item.itemName,
        row: currentRow,
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

      writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols);

      transferLog.push({
        tableNumber: '③',
        group: groupName,
        itemName: item.itemName,
        row: currentRow,
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

      writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols);

      transferLog.push({
        tableNumber: '④',
        group: '労務費',
        itemName: item.itemName,
        row: currentRow,
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
    const maxOtherRows = otherGroup.maxDataRows - 1;  // 最後の1行は「その他」用に確保

    // 個別転記
    for (let i = 0; i < Math.min(sortedOther.length, maxOtherRows); i++) {
      const currentRow = otherGroup.dataStartRow + i;
      const item = sortedOther[i];

      writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols);

      transferLog.push({
        tableNumber: '④',
        group: 'その他経費',
        itemName: item.itemName,
        row: currentRow,
        status: '成功',
        amounts: item.amounts
      });

      Logger.log(`④-その他経費: ${item.itemName} (当期: ${item.amounts['当期']})`);
    }

    // 超過分は合計
    if (sortedOther.length > maxOtherRows && otherGroup.otherRow) {
      const excessOther = sortedOther.slice(maxOtherRows);
      const excessAmounts = aggregateAmounts(excessOther);
      const excessRow = otherGroup.otherRow;

      writeExcelRowData(sheet, tableConfig.itemColumn, excessRow, 'その他', excessAmounts, periodCols);

      transferLog.push({
        tableNumber: '④',
        group: 'その他経費',
        itemName: 'その他',
        row: excessRow,
        status: '成功（集計）',
        amounts: excessAmounts,
        aggregatedCount: excessOther.length
      });

      Logger.log(`④-その他: ${excessOther.length}件を集計`);
    }
  }

  Logger.log(`④製造経費比較表: 労務費${laborItems.length}件、その他経費${otherItems.length}件を転記完了`);

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

    writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols);

    transferLog.push({
      tableNumber: '②',
      group: '人件費',
      itemName: item.itemName,
      row: currentRow,
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

    writeExcelRowData(sheet, tableConfig.itemColumn, currentRow, item.itemName, item.amounts, periodCols);

    transferLog.push({
      tableNumber: '②',
      group: 'その他経費',
      itemName: item.itemName,
      row: currentRow,
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
      writeExcelRowData(sheet, itemColumn, row, item.itemName, item.amounts, periodColumns);
      transferLog.push({ tableNumber: 'BS', group: groupName, itemName: item.itemName, row: row, status: '成功', amounts: item.amounts });
    });

    // 行数を超えた項目を「その他」として集計
    if (sortedItems.length > maxRows) {
      const excessItems = sortedItems.slice(maxRows);
      const otherAmounts = aggregateAmounts(excessItems);
      const otherRow = groupConfig.dataEndRow; // 最後の行を「その他」とする

      writeExcelRowData(sheet, itemColumn, otherRow, `その他${groupName}`, otherAmounts, periodColumns);
      transferLog.push({ tableNumber: 'BS', group: groupName, itemName: `その他${groupName}`, row: otherRow, status: '成功（集計）', amounts: otherAmounts, aggregatedCount: excessItems.length });
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

  const mappingRange = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, 3).getValues();
  const mappingData = {};
  mappingRange.forEach(row => {
    if (row[0]) {
      mappingData[normalizeJapaneseText(row[0])] = {
        table: normalizeJapaneseText(row[1] || ''),
        group: normalizeJapaneseText(row[2] || '')
      };
    }
  });

  ocrData.forEach(row => {
    if (!row[0] || row[1] === undefined) return;

    const itemName = normalizeJapaneseText(row[0]);
    const amount = parseJapaneseNumber(row[1]);

    if (!isSafeText(itemName) || isNaN(amount) || itemName.includes('計') || itemName.includes('合計')) {
      return;
    }

    const itemData = {
      itemName: itemName,
      amounts: { '前々期': 0, '前期': 0, '当期': amount }
    };

    const userMapping = mappingData[itemName];
    if (!userMapping || !userMapping.table) {
      Logger.log(`未マッピング（または転記先が空）: "${itemName}"`);
      return;
    }

    const { table, group } = userMapping;

    if (table.startsWith('BS:')) {
      if (classifiedData[table]) {
        classifiedData[table].push({ ...itemData, mappedTable: table, mappedGroup: group });
        Logger.log(`✓ BSマッピング適用: ${itemName} → ${table}`);
      }
    } else if (table.startsWith('⑤')) {
      const category = group || '営業外収益';
      if (classifiedData[table] && classifiedData[table][category]) {
        classifiedData[table][category].push({ ...itemData, mappedTable: table, mappedGroup: category });
        Logger.log(`✓ PLマッピング適用: ${itemName} → ${table} [${category}]`);
      }
    } else if (classifiedData[table]) {
      classifiedData[table].push({ ...itemData, mappedTable: table, mappedGroup: group });
      Logger.log(`✓ PLマッピング適用: ${itemName} → ${table} [${group}]`);
    }
  });

  return classifiedData;
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
  const ocrData = ocrSheet.getRange(1, 1, ocrSheet.getLastRow(), 2).getValues();
  const classifiedData = classifyOcrDataByMapping(ss, ocrData);

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
