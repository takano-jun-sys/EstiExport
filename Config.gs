/**
 * 設定管理スクリプト（AppSheet見積書システム用）
 */

const CONFIG = {
  // タイムゾーン
  TIMEZONE: 'Asia/Tokyo',

  // スプレッドシートID（AppSheet DB）
  SPREADSHEET_ID: '1kxEuJFDufdV7aEutiOlhpCYA4_1DNLSrWVR0cAITxbU',

  // テンプレートスプレッドシートID
  TEMPLATE_SPREADSHEET_ID: '1H7_tPrwyXj1c9m6PXYem-RJB6edCIp9DD2sB2mzXSCo',

  // シート名
  SHEET_NAMES: {
    JOBS: 'Jobs',
    DETAILS: 'Details',
    PROJECTS: 'Projects'
  },

  // テンプレートシート名
  TEMPLATE_SHEETS: {
    UNDER_13: '13未満',
    OVER_14: '14以上',
    JOB_UNIT: 'Job単位'
  },

  // Jobsシートの列番号（1始まり）
  JOBS_COLUMNS: {
    PROJECT_ID: 2,           // B列
    JOB_ID: 3,               // C列（Job. ID）
    案件名: 4,                // D列
    ステータス: 5,            // E列
    進行管理費率: 6,          // F列
    受注日: 7,               // G列
    納期: 8,                 // H列
    共通東京大阪: 9,         // I列
    発行日時: 10,            // J列
    RELATED_DETAILS: 11,     // K列（仮想列）
    小計: 12,                // L列（仮想列）
    進行管理費額: 13,        // M列（仮想列）
    合計: 14,                // N列（仮想列）
    参考見積額: 15,          // O列（仮想列）
    入力チェック警告: 16,    // P列（仮想列）
    警告_不適切なマイナス: 17, // Q列（仮想列）

    // 新規追加列（これらをAppSheetで追加する必要がある）
    クライアント名: 18,
    品名: 19,
    仕様1: 20,
    仕様2: 21,
    仕様3: 22,
    仕様4: 23,
    担当: 24,
    見積書スプレッドシートURL: 25,
    見積書PDF_URL: 26,
    最終発行日時: 27
  },

  // Detailsシートの列番号
  DETAILS_COLUMNS: {
    明細ID: 2,               // B列
    JOB_ID: 3,               // C列（Job. ID）
    枝番: 4,                 // D列
    行番号: 5,               // E列（仮想列）
    作業項目: 6,             // F列
    担当部署: 7,             // G列
    数量: 8,                 // H列
    単位: 9,                 // I列
    単価: 10,                // J列
    内製外注: 11,            // K列
    制作単価: 12,            // L列
    仕入れ金額: 13,          // M列
    かけ率: 14,              // N列
    仕入先: 15,              // O列（Ref）
    単価調整: 16,            // P列
    カテゴリ: 17,            // Q列（Ref）
    メモ: 18,                // R列
    金額: 19,                // S列（仮想列）
    粗利: 20                 // T列（仮想列）
  },

  // テンプレートセル位置（13未満シート）
  TEMPLATE_CELLS_13: {
    クライアント名: 'A2:C2',
    品名: 'B3:C3',
    仕様1: 'B12:C12',
    仕様2: 'B13:C13',
    仕様3: 'B14:C14',
    仕様4: 'B15:C15',
    担当: 'H10',
    PROJECT_JOB_ID: 'G1:H1',
    明細開始行: 18,
    明細終了行: 30,
    列: {
      JOB_ID: 'A',
      作業項目: 'B',
      メモ: 'C',
      単価: 'D',
      数量: 'E',
      単位: 'F',
      金額: 'G:H'  // 結合セル
    }
  },

  // テンプレートセル位置（14以上シート）
  TEMPLATE_CELLS_14: {
    クライアント名: 'A2:C2',
    品名: 'B3:C3',
    仕様1: 'B12:C12',
    仕様2: 'B13:C13',
    仕様3: 'B14:C14',
    仕様4: 'B15:C15',
    担当: 'H10',
    PROJECT_JOB_ID: 'G1:H1',
    ページ1: {
      明細開始行: 18,
      明細終了行: 37
    },
    ページ2: {
      明細開始行: 39,
      明細終了行: 62
    },
    列: {
      JOB_ID: 'A',
      作業項目: 'B',
      メモ: 'C',
      単価: 'D',
      数量: 'E',
      単位: 'F',
      金額: 'G:H'  // 結合セル
    }
  },

  // テンプレートセル位置（Job単位シート）
  TEMPLATE_CELLS_JOB: {
    案件名: 'B1',
    PROJECT_JOB_ID: 'G1:H1',
    明細開始行: 3,
    列: {
      JOB_ID: 'A',
      作業項目: 'B',
      メモ: 'C',
      単価: 'D',
      数量: 'E',
      単位: 'F',
      金額: 'G:H'
    }
  },

  // PDFフォルダ名
  PDF_FOLDER_NAME: '見積書PDF'
};

/**
 * PDFフォルダを取得または作成
 */
function getPDFFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.PDF_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(CONFIG.PDF_FOLDER_NAME);
  }
}

/**
 * AppSheetスプレッドシートを取得
 */
function getAppSheetSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/**
 * テンプレートスプレッドシートを取得
 */
function getTemplateSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.TEMPLATE_SPREADSHEET_ID);
}
