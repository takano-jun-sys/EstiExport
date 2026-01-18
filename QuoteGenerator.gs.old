/**
 * 見積書生成スクリプト
 * テンプレートから見積書スプレッドシートを作成
 */

/**
 * テンプレートから見積書スプレッドシートを生成
 * @param {Object} quoteData - 見積データオブジェクト
 * @return {Spreadsheet} - 生成された見積書スプレッドシート
 */
function generateQuoteSpreadsheet(quoteData) {
  try {
    Logger.log('見積書スプレッドシート生成開始: ' + quoteData.quoteNumber);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATE);

    if (!templateSheet) {
      throw new Error('テンプレートシートが見つかりません');
    }

    // テンプレートシートをコピーして新しいスプレッドシートを作成
    const newSpreadsheet = SpreadsheetApp.create('見積書_' + quoteData.quoteNumber);
    const newSheet = newSpreadsheet.getActiveSheet();
    newSheet.setName('見積書');

    // テンプレートの内容をコピー
    const templateData = templateSheet.getDataRange();
    const values = templateData.getValues();
    const formats = templateData.getNumberFormats();

    newSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    newSheet.getRange(1, 1, formats.length, formats[0].length).setNumberFormats(formats);

    // テンプレートのフォーマットをコピー
    copySheetFormatting(templateSheet, newSheet);

    // 見積データを埋め込み
    fillQuoteData(newSheet, quoteData);

    Logger.log('見積書スプレッドシート生成完了: ' + newSpreadsheet.getUrl());
    return newSpreadsheet;

  } catch (error) {
    Logger.log('見積書スプレッドシート生成エラー: ' + error.message);
    throw error;
  }
}

/**
 * シートのフォーマットをコピー
 * @param {Sheet} sourceSheet - コピー元シート
 * @param {Sheet} targetSheet - コピー先シート
 */
function copySheetFormatting(sourceSheet, targetSheet) {
  // 列幅をコピー
  for (let col = 1; col <= sourceSheet.getMaxColumns(); col++) {
    const width = sourceSheet.getColumnWidth(col);
    targetSheet.setColumnWidth(col, width);
  }

  // 行高をコピー
  for (let row = 1; row <= sourceSheet.getMaxRows(); row++) {
    const height = sourceSheet.getRowHeight(row);
    targetSheet.setRowHeight(row, height);
  }

  // 背景色、フォント、罫線などのスタイルをコピー
  const sourceRange = sourceSheet.getDataRange();
  const targetRange = targetSheet.getRange(1, 1, sourceRange.getNumRows(), sourceRange.getNumColumns());

  targetRange.setBackgrounds(sourceRange.getBackgrounds());
  targetRange.setFontColors(sourceRange.getFontColors());
  targetRange.setFontFamilies(sourceRange.getFontFamilies());
  targetRange.setFontSizes(sourceRange.getFontSizes());
  targetRange.setFontWeights(sourceRange.getFontWeights());
  targetRange.setFontStyles(sourceRange.getFontStyles());
  targetRange.setHorizontalAlignments(sourceRange.getHorizontalAlignments());
  targetRange.setVerticalAlignments(sourceRange.getVerticalAlignments());
}

/**
 * 見積データをシートに埋め込み
 * @param {Sheet} sheet - 対象シート
 * @param {Object} quoteData - 見積データ
 */
function fillQuoteData(sheet, quoteData) {
  // 会社情報を取得
  const companyInfo = getCompanyInfo();

  // 基本情報を設定
  sheet.getRange(CONFIG.TEMPLATE_CELLS.QUOTE_NUMBER).setValue(quoteData.quoteNumber);

  // 日付のフォーマット
  const formattedDate = Utilities.formatDate(
    quoteData.date,
    CONFIG.TIMEZONE,
    CONFIG.DEFAULTS.DATE_FORMAT
  );
  sheet.getRange(CONFIG.TEMPLATE_CELLS.DATE).setValue(formattedDate);

  // 有効期限
  if (quoteData.expiryDate) {
    const formattedExpiryDate = Utilities.formatDate(
      quoteData.expiryDate,
      CONFIG.TIMEZONE,
      CONFIG.DEFAULTS.DATE_FORMAT
    );
    sheet.getRange(CONFIG.TEMPLATE_CELLS.EXPIRY_DATE).setValue(formattedExpiryDate);
  } else {
    // デフォルト30日後
    const expiryDate = new Date(quoteData.date);
    expiryDate.setDate(expiryDate.getDate() + CONFIG.DEFAULTS.QUOTE_VALID_DAYS);
    const formattedExpiryDate = Utilities.formatDate(
      expiryDate,
      CONFIG.TIMEZONE,
      CONFIG.DEFAULTS.DATE_FORMAT
    );
    sheet.getRange(CONFIG.TEMPLATE_CELLS.EXPIRY_DATE).setValue(formattedExpiryDate);
  }

  // 顧客情報
  sheet.getRange(CONFIG.TEMPLATE_CELLS.CUSTOMER_NAME).setValue(quoteData.customerName + ' 御中');
  sheet.getRange(CONFIG.TEMPLATE_CELLS.CUSTOMER_ZIP).setValue('〒' + quoteData.customerZip);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.CUSTOMER_ADDRESS).setValue(quoteData.customerAddress);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.CUSTOMER_PHONE).setValue('TEL: ' + quoteData.customerPhone);

  // 件名
  sheet.getRange(CONFIG.TEMPLATE_CELLS.SUBJECT).setValue(quoteData.subject);

  // 自社情報
  sheet.getRange(CONFIG.TEMPLATE_CELLS.COMPANY_NAME).setValue(companyInfo.name);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.COMPANY_ZIP).setValue('〒' + companyInfo.zip);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.COMPANY_ADDRESS).setValue(companyInfo.address);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.COMPANY_PHONE).setValue('TEL: ' + companyInfo.phone);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.COMPANY_FAX).setValue('FAX: ' + companyInfo.fax);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.COMPANY_EMAIL).setValue('E-mail: ' + companyInfo.email);

  // 明細行を追加
  fillItems(sheet, quoteData.items);

  // 金額を計算して設定
  calculateAndSetAmounts(sheet, quoteData.items, quoteData.taxRate);

  // 備考
  if (quoteData.note) {
    sheet.getRange(CONFIG.TEMPLATE_CELLS.NOTE).setValue(quoteData.note);
  }

  // 振込先情報（備考欄の下に追加）
  if (companyInfo.bankName) {
    const bankInfo = `\n【お振込先】\n${companyInfo.bankName} ${companyInfo.branchName}\n${companyInfo.accountType} ${companyInfo.accountNumber}\n${companyInfo.accountName}`;
    const currentNote = sheet.getRange(CONFIG.TEMPLATE_CELLS.NOTE).getValue();
    sheet.getRange(CONFIG.TEMPLATE_CELLS.NOTE).setValue(currentNote + bankInfo);
  }
}

/**
 * 明細行をシートに追加
 * @param {Sheet} sheet - 対象シート
 * @param {Array} items - 明細配列
 */
function fillItems(sheet, items) {
  if (!items || items.length === 0) {
    return;
  }

  const startRow = CONFIG.TEMPLATE_CELLS.ITEMS_START_ROW;

  items.forEach((item, index) => {
    const row = startRow + index;

    // 行を挿入（最初の行以外）
    if (index > 0) {
      sheet.insertRowAfter(row - 1);
      // 前の行のフォーマットをコピー
      const sourceRange = sheet.getRange(row - 1, 1, 1, 7);
      const targetRange = sheet.getRange(row, 1, 1, 7);
      sourceRange.copyFormatToRange(sheet, 1, 7, row, row);
    }

    // 明細データを設定
    sheet.getRange(row, 1).setValue(index + 1); // 番号
    sheet.getRange(row, 2).setValue(item.item || ''); // 品名
    sheet.getRange(row, 3).setValue(item.quantity || 0); // 数量
    sheet.getRange(row, 4).setValue(item.unit || CONFIG.DEFAULTS.ITEM_UNIT); // 単位
    sheet.getRange(row, 5).setValue(item.unitPrice || 0); // 単価
    sheet.getRange(row, 6).setValue(item.amount || (item.quantity * item.unitPrice)); // 金額
    sheet.getRange(row, 7).setValue(item.note || ''); // 備考

    // 金額のフォーマットを設定
    sheet.getRange(row, 5).setNumberFormat(CONFIG.DEFAULTS.CURRENCY_FORMAT);
    sheet.getRange(row, 6).setNumberFormat(CONFIG.DEFAULTS.CURRENCY_FORMAT);
  });
}

/**
 * 金額を計算して設定
 * @param {Sheet} sheet - 対象シート
 * @param {Array} items - 明細配列
 * @param {number} taxRate - 消費税率
 */
function calculateAndSetAmounts(sheet, items, taxRate) {
  // 小計を計算
  let subtotal = 0;
  items.forEach(item => {
    subtotal += item.amount || (item.quantity * item.unitPrice) || 0;
  });

  // 消費税を計算
  const tax = Math.floor(subtotal * taxRate);

  // 合計を計算
  const total = subtotal + tax;

  // シートに設定
  sheet.getRange(CONFIG.TEMPLATE_CELLS.SUBTOTAL).setValue(subtotal);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.TAX).setValue(tax);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.TOTAL).setValue(total);

  // 金額のフォーマット
  sheet.getRange(CONFIG.TEMPLATE_CELLS.SUBTOTAL).setNumberFormat(CONFIG.DEFAULTS.CURRENCY_FORMAT);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.TAX).setNumberFormat(CONFIG.DEFAULTS.CURRENCY_FORMAT);
  sheet.getRange(CONFIG.TEMPLATE_CELLS.TOTAL).setNumberFormat(CONFIG.DEFAULTS.CURRENCY_FORMAT);
}
