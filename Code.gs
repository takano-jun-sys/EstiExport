/**
 * EstiExport - AppSheet連携見積書作成システム
 * メインエントリーポイント
 */

/**
 * メイン処理：未作成の見積書を処理
 * この関数をトリガーで定期実行することで自動化可能
 */
function processNewQuotes() {
  try {
    Logger.log('見積書処理を開始します');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DATA);

    if (!dataSheet) {
      throw new Error('見積データシートが見つかりません');
    }

    // データを取得（ヘッダー行を除く）
    const lastRow = dataSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('処理するデータがありません');
      return;
    }

    const dataRange = dataSheet.getRange(2, 1, lastRow - 1, CONFIG.COLUMNS.DATA.PDF_URL);
    const data = dataRange.getValues();

    let processedCount = 0;

    // 各行を処理
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowIndex = i + 2; // 実際のシート行番号
      const status = row[CONFIG.COLUMNS.DATA.STATUS - 1];

      // ステータスが「未作成」または空白の場合のみ処理
      if (status === '未作成' || status === '') {
        try {
          Logger.log(`行 ${rowIndex} の見積書を作成中...`);

          // 見積データを構造化
          const quoteData = parseQuoteData(row);

          // 見積書スプレッドシートを生成
          const newSpreadsheet = generateQuoteSpreadsheet(quoteData);

          // PDFを生成
          const pdfFile = exportToPDF(newSpreadsheet, quoteData.quoteNumber);

          // URLを元のシートに書き戻し
          dataSheet.getRange(rowIndex, CONFIG.COLUMNS.DATA.SPREADSHEET_URL).setValue(newSpreadsheet.getUrl());
          dataSheet.getRange(rowIndex, CONFIG.COLUMNS.DATA.PDF_URL).setValue(pdfFile.getUrl());
          dataSheet.getRange(rowIndex, CONFIG.COLUMNS.DATA.STATUS).setValue('作成済み');

          processedCount++;
          Logger.log(`行 ${rowIndex} の見積書を作成しました: ${quoteData.quoteNumber}`);

        } catch (error) {
          Logger.log(`行 ${rowIndex} の処理中にエラーが発生しました: ${error.message}`);
          dataSheet.getRange(rowIndex, CONFIG.COLUMNS.DATA.STATUS).setValue('エラー');
        }
      }
    }

    Logger.log(`処理完了: ${processedCount}件の見積書を作成しました`);

  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * 行データを見積データオブジェクトにパース
 * @param {Array} row - シートの行データ
 * @return {Object} - 構造化された見積データ
 */
function parseQuoteData(row) {
  // 明細JSONをパース
  let items = [];
  const itemsJson = row[CONFIG.COLUMNS.DATA.ITEMS_JSON - 1];

  if (itemsJson) {
    try {
      items = JSON.parse(itemsJson);
    } catch (error) {
      Logger.log('明細JSONのパースに失敗しました。デフォルト値を使用します。');
      items = [];
    }
  }

  // 日付の処理
  let quoteDate = row[CONFIG.COLUMNS.DATA.DATE - 1];
  if (!(quoteDate instanceof Date)) {
    quoteDate = new Date(quoteDate);
  }

  let expiryDate = row[CONFIG.COLUMNS.DATA.EXPIRY_DATE - 1];
  if (expiryDate && !(expiryDate instanceof Date)) {
    expiryDate = new Date(expiryDate);
  }

  return {
    quoteNumber: row[CONFIG.COLUMNS.DATA.QUOTE_NUMBER - 1] || generateQuoteNumber(),
    date: quoteDate,
    customerName: row[CONFIG.COLUMNS.DATA.CUSTOMER_NAME - 1] || '',
    customerZip: row[CONFIG.COLUMNS.DATA.CUSTOMER_ZIP - 1] || '',
    customerAddress: row[CONFIG.COLUMNS.DATA.CUSTOMER_ADDRESS - 1] || '',
    customerPhone: row[CONFIG.COLUMNS.DATA.CUSTOMER_PHONE - 1] || '',
    subject: row[CONFIG.COLUMNS.DATA.SUBJECT - 1] || '御見積書',
    items: items,
    note: row[CONFIG.COLUMNS.DATA.NOTE - 1] || '',
    taxRate: row[CONFIG.COLUMNS.DATA.TAX_RATE - 1] || 0.10,
    expiryDate: expiryDate || null
  };
}

/**
 * 見積番号を生成
 * 形式: QUOTE-YYYYMMDD-XXX
 * @return {string} - 見積番号
 */
function generateQuoteNumber() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyyMMdd');
  const randomNum = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return `QUOTE-${dateStr}-${randomNum}`;
}

/**
 * 手動実行用：選択した行の見積書を再生成
 */
function regenerateSelectedQuote() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DATA);
  const activeRange = dataSheet.getActiveRange();
  const row = activeRange.getRow();

  if (row < 2) {
    Browser.msgBox('データ行を選択してください');
    return;
  }

  const rowData = dataSheet.getRange(row, 1, 1, CONFIG.COLUMNS.DATA.PDF_URL).getValues()[0];

  try {
    const quoteData = parseQuoteData(rowData);
    const newSpreadsheet = generateQuoteSpreadsheet(quoteData);
    const pdfFile = exportToPDF(newSpreadsheet, quoteData.quoteNumber);

    dataSheet.getRange(row, CONFIG.COLUMNS.DATA.SPREADSHEET_URL).setValue(newSpreadsheet.getUrl());
    dataSheet.getRange(row, CONFIG.COLUMNS.DATA.PDF_URL).setValue(pdfFile.getUrl());
    dataSheet.getRange(row, CONFIG.COLUMNS.DATA.STATUS).setValue('作成済み');

    Browser.msgBox('見積書を再生成しました');
  } catch (error) {
    Browser.msgBox('エラー: ' + error.message);
  }
}

/**
 * カスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('見積書システム')
    .addItem('新規見積書を処理', 'processNewQuotes')
    .addItem('選択行の見積書を再生成', 'regenerateSelectedQuote')
    .addSeparator()
    .addItem('初期セットアップ', 'setupSheets')
    .addToUi();
}
