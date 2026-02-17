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
    BETWEEN_14_20: '14-20',
    OVER_20: '20以上'
  },

  // Jobsシートの列番号（1始まり）
  JOBS_COLUMNS: {
    PROJECT_ID: 1,           // A列
    JOB_ID: 2,               // B列（Job. ID）
    案件名: 3,                // C列
    ステータス: 4,            // D列
    進行管理費率: 5,          // E列
    受注日: 6,               // F列
    納期: 7,                 // G列
    共通東京大阪: 8,         // H列
    発行日時: 9,            // I列
    クライアント名: 10,      // J列（クライアント）
    品名: 11,                // K列
    仕様1: 12,               // L列
    仕様2: 13,               // M列
    仕様3: 14,               // N列
    仕様4: 15,               // O列
    担当: 16,                // P列
    見積書スプレッドシートURL: 17,  // Q列
    見積書PDF_URL: 18,       // R列
    最終発行日時: 19         // S列
  },

  // Detailsシートの列番号
  DETAILS_COLUMNS: {
    明細ID: 1,               // A列
    JOB_ID: 2,               // B列（Job. ID）
    枝番: 3,                 // C列
    行番号: 4,               // D列
    作業項目: 5,             // E列
    作業詳細: 6,             // F列
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
    金額: 18,                // R列（仮想列 - 存在しない場合あり）
    見積に含める: 19         // S列
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
    生成日時: 'G16:H16',  // 生成日時を表示（結合セル）
    端数調整: 'G33:H33',  // 端数調整金額（結合セル）
    明細開始行: 18,
    明細終了行: 30,
    列: {
      JOB_ID: 'A',
      作業項目: 'B',
      作業詳細: 'C',
      単価: 'D',
      数量: 'E',
      単位: 'F',
      金額: 'G:H'  // 結合セル
    }
  },

  // テンプレートセル位置（14-20シート）
  TEMPLATE_CELLS_14_20: {
    クライアント名: 'A2:C2',
    品名: 'B3:C3',
    仕様1: 'B12:C12',
    仕様2: 'B13:C13',
    仕様3: 'B14:C14',
    仕様4: 'B15:C15',
    担当: 'H10',
    PROJECT_JOB_ID: 'G1:H1',
    生成日時: 'G16:H16',  // 生成日時を表示（結合セル）
    端数調整: 'G40:H40',  // 端数調整金額（結合セル）
    明細開始行: 18,
    明細終了行: 37,
    列: {
      JOB_ID: 'A',
      作業項目: 'B',
      作業詳細: 'C',
      単価: 'D',
      数量: 'E',
      単位: 'F',
      金額: 'G:H'  // 結合セル
    }
  },

  // テンプレートセル位置（20以上シート）
  TEMPLATE_CELLS_20_OVER: {
    クライアント名: 'A2:C2',
    品名: 'B3:C3',
    仕様1: 'B12:C12',
    仕様2: 'B13:C13',
    仕様3: 'B14:C14',
    仕様4: 'B15:C15',
    担当: 'H10',
    PROJECT_JOB_ID: 'G1:H1',
    生成日時: 'G16:H16',  // 生成日時を表示（結合セル）
    端数調整: 'G65:H65',  // 端数調整金額（結合セル）
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
      作業詳細: 'C',
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
      作業詳細: 'C',
      単価: 'D',
      数量: 'E',
      単位: 'F',
      金額: 'G:H'
    }
  },

  // 保存先フォルダID
  QUOTE_FOLDER_ID: '15O79II-3SKSD1UsG033lakI2ypnlq0Ym',

  // ステータス設定
  STATUS: {
    初回生成時: '見積中1',
    見積中1: '見積中2',
    見積中2: '見積中3',
    見積中3: '見積中4',
    見積中4: '見積中4'  // 見積中4が最大
  }
};

/**
 * 見積書保存先フォルダを取得
 */
function getQuoteFolder() {
  return DriveApp.getFolderById(CONFIG.QUOTE_FOLDER_ID);
}

/**
 * PDFフォルダを取得または作成（スプレッドシート名と同じフォルダ）
 * @param {string} folderName - フォルダ名（ProjectID-JobID）
 * @return {Folder} - PDFフォルダ
 */
function getPDFFolder(folderName) {
  const parentFolder = getQuoteFolder();

  // 既存のフォルダを検索
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }

  // なければ作成
  return parentFolder.createFolder(folderName);
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
