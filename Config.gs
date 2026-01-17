/**
 * 設定管理スクリプト
 * 各種設定値とシート構成を定義
 */

const CONFIG = {
  // タイムゾーン
  TIMEZONE: 'Asia/Tokyo',

  // シート名
  SHEET_NAMES: {
    DATA: '見積データ',
    TEMPLATE: '見積書テンプレート',
    SETTINGS: '設定'
  },

  // 見積データシートの列番号
  COLUMNS: {
    DATA: {
      QUOTE_NUMBER: 1,      // A: 見積番号
      DATE: 2,              // B: 日付
      CUSTOMER_NAME: 3,     // C: 顧客名
      CUSTOMER_ZIP: 4,      // D: 郵便番号
      CUSTOMER_ADDRESS: 5,  // E: 住所
      CUSTOMER_PHONE: 6,    // F: 電話番号
      SUBJECT: 7,           // G: 件名
      ITEMS_JSON: 8,        // H: 明細JSON
      NOTE: 9,              // I: 備考
      TAX_RATE: 10,         // J: 消費税率
      EXPIRY_DATE: 11,      // K: 有効期限
      STATUS: 12,           // L: ステータス
      SPREADSHEET_URL: 13,  // M: スプレッドシートURL
      PDF_URL: 14           // N: PDF URL
    }
  },

  // テンプレートシートのセル位置
  TEMPLATE_CELLS: {
    QUOTE_NUMBER: 'G2',     // 見積番号
    DATE: 'G3',             // 見積日
    EXPIRY_DATE: 'G4',      // 有効期限
    CUSTOMER_NAME: 'B6',    // 顧客名
    CUSTOMER_ZIP: 'B7',     // 郵便番号
    CUSTOMER_ADDRESS: 'B8', // 住所
    CUSTOMER_PHONE: 'B9',   // 電話番号
    SUBJECT: 'B11',         // 件名
    ITEMS_START_ROW: 14,    // 明細開始行
    SUBTOTAL: 'G50',        // 小計
    TAX: 'G51',             // 消費税
    TOTAL: 'G52',           // 合計
    NOTE: 'B54',            // 備考
    COMPANY_NAME: 'F6',     // 自社名
    COMPANY_ZIP: 'F7',      // 自社郵便番号
    COMPANY_ADDRESS: 'F8',  // 自社住所
    COMPANY_PHONE: 'F9',    // 自社電話番号
    COMPANY_FAX: 'F10',     // 自社FAX
    COMPANY_EMAIL: 'F11'    // 自社メール
  },

  // 明細行のフォーマット
  ITEM_COLUMNS: {
    NO: 0,          // 番号
    ITEM: 1,        // 品名・サービス名
    QUANTITY: 2,    // 数量
    UNIT: 3,        // 単位
    UNIT_PRICE: 4,  // 単価
    AMOUNT: 5,      // 金額
    NOTE: 6         // 備考
  },

  // デフォルト値
  DEFAULTS: {
    TAX_RATE: 0.10,
    CURRENCY_FORMAT: '¥#,##0',
    DATE_FORMAT: 'yyyy年MM月dd日',
    ITEM_UNIT: '式',
    QUOTE_VALID_DAYS: 30  // 見積有効期限（日数）
  },

  // PDF設定
  PDF: {
    FOLDER_NAME: '見積書PDF',  // PDFを保存するフォルダ名
    FILE_PREFIX: '見積書_'     // PDFファイル名のプレフィックス
  }
};

/**
 * 設定シートから会社情報を取得
 * @return {Object} - 会社情報オブジェクト
 */
function getCompanyInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);

  if (!settingsSheet) {
    Logger.log('設定シートが見つかりません。デフォルト値を使用します。');
    return getDefaultCompanyInfo();
  }

  try {
    // 設定シートの構造: A列にキー、B列に値
    const data = settingsSheet.getRange('A2:B20').getValues();
    const settings = {};

    data.forEach(row => {
      if (row[0]) {
        settings[row[0]] = row[1];
      }
    });

    return {
      name: settings['会社名'] || '',
      zip: settings['郵便番号'] || '',
      address: settings['住所'] || '',
      phone: settings['電話番号'] || '',
      fax: settings['FAX'] || '',
      email: settings['メールアドレス'] || '',
      bankName: settings['銀行名'] || '',
      branchName: settings['支店名'] || '',
      accountType: settings['口座種別'] || '',
      accountNumber: settings['口座番号'] || '',
      accountName: settings['口座名義'] || ''
    };

  } catch (error) {
    Logger.log('会社情報の取得に失敗しました: ' + error.message);
    return getDefaultCompanyInfo();
  }
}

/**
 * デフォルトの会社情報を返す
 * @return {Object} - デフォルト会社情報
 */
function getDefaultCompanyInfo() {
  return {
    name: '株式会社サンプル',
    zip: '100-0001',
    address: '東京都千代田区千代田1-1-1',
    phone: '03-1234-5678',
    fax: '03-1234-5679',
    email: 'info@example.com',
    bankName: 'サンプル銀行',
    branchName: '本店',
    accountType: '普通',
    accountNumber: '1234567',
    accountName: 'カ）サンプル'
  };
}

/**
 * 設定シートに会社情報を保存
 * @param {Object} companyInfo - 会社情報オブジェクト
 */
function saveCompanyInfo(companyInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);

  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(CONFIG.SHEET_NAMES.SETTINGS);
  }

  const data = [
    ['設定項目', '値'],
    ['会社名', companyInfo.name || ''],
    ['郵便番号', companyInfo.zip || ''],
    ['住所', companyInfo.address || ''],
    ['電話番号', companyInfo.phone || ''],
    ['FAX', companyInfo.fax || ''],
    ['メールアドレス', companyInfo.email || ''],
    ['銀行名', companyInfo.bankName || ''],
    ['支店名', companyInfo.branchName || ''],
    ['口座種別', companyInfo.accountType || ''],
    ['口座番号', companyInfo.accountNumber || ''],
    ['口座名義', companyInfo.accountName || '']
  ];

  settingsSheet.getRange(1, 1, data.length, 2).setValues(data);
  settingsSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#4A86E8').setFontColor('#FFFFFF');
}

/**
 * PDFフォルダを取得または作成
 * @return {Folder} - PDFフォルダ
 */
function getPDFFolder() {
  const folderName = CONFIG.PDF.FOLDER_NAME;
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}
