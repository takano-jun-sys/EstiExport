/**
 * 見積書Export機能
 */

/**
 * 見積書を生成（ダイアログから呼び出される）
 * @param {string} jobId - Job ID
 * @param {Object} formData - フォームデータ
 * @return {Object} - 結果（success, spreadsheetUrl, pdfUrl）
 */
function generateQuote(jobId, formData) {
  try {
    Logger.log('見積書生成開始: ' + jobId);

    // 1. Jobデータを更新
    const updateResult = updateJobData(jobId, formData);
    if (!updateResult.success) {
      throw new Error('Jobデータの更新に失敗: ' + updateResult.error);
    }

    // 2. 更新後のJobデータとDetailsを取得
    const jobData = getJobData(jobId);
    const details = getDetails(jobId);

    if (details.length === 0) {
      throw new Error('明細が見つかりません');
    }

    Logger.log('明細数: ' + details.length);

    // 3. テンプレートを選択（13未満 / 14-20 / 20以上）
    let templateName;
    if (details.length <= 13) {
      templateName = CONFIG.TEMPLATE_SHEETS.UNDER_13;
    } else if (details.length <= 20) {
      templateName = CONFIG.TEMPLATE_SHEETS.BETWEEN_14_20;
    } else {
      templateName = CONFIG.TEMPLATE_SHEETS.OVER_20;
    }

    Logger.log('使用テンプレート: ' + templateName);

    // 4. スプレッドシートを作成または取得
    const spreadsheet = getOrCreateQuoteSpreadsheet(jobData);

    // 5. 新しいシートを追加（発行日時が名前）
    const now = new Date();
    const sheetName = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const newSheet = addQuoteSheet(spreadsheet, templateName, sheetName);

    // 6. データを埋め込み
    fillQuoteData(newSheet, jobData, details, templateName, now);

    // 6.5. スプレッドシートへの書き込みを確実に反映させる
    Logger.log('スプレッドシートへの書き込みを反映中...');
    SpreadsheetApp.flush();
    Logger.log('スプレッドシートへの書き込み完了');

    // 7. PDFを生成（ページ番号付き）
    const pdfFile = exportSheetToPDF(spreadsheet, newSheet, jobData, now);

    // 8. 生成したシートを直接開くURLを作成
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/edit#gid=${newSheet.getSheetId()}`;
    Logger.log('シートURL: ' + sheetUrl);

    // 9. JobsテーブルにURLを保存
    saveQuoteUrls(jobData.rowIndex, sheetUrl, pdfFile.getUrl(), now);

    Logger.log('見積書生成完了');

    return {
      success: true,
      spreadsheetUrl: sheetUrl,
      pdfUrl: pdfFile.getUrl(),
      message: '見積書を生成しました'
    };

  } catch (error) {
    Logger.log('generateQuoteエラー: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 見積書スプレッドシートを取得または作成
 * @param {Object} jobData - Jobデータ
 * @return {Spreadsheet} - スプレッドシート
 */
function getOrCreateQuoteSpreadsheet(jobData) {
  const fileName = jobData.projectId + '-' + jobData.jobId;
  const quoteFolder = getQuoteFolder();

  // 既存のスプレッドシートURLがあれば開く
  if (jobData.見積書スプレッドシートURL) {
    try {
      return SpreadsheetApp.openByUrl(jobData.見積書スプレッドシートURL);
    } catch (error) {
      Logger.log('既存スプレッドシートが見つからないため再検索: ' + error.message);
    }
  }

  // フォルダ内に同名のスプレッドシートが既にあるかチェック
  const existingFiles = quoteFolder.getFilesByName(fileName);
  if (existingFiles.hasNext()) {
    const existingFile = existingFiles.next();
    Logger.log('既存のスプレッドシートを再利用: ' + fileName);
    return SpreadsheetApp.openById(existingFile.getId());
  }

  // 新規作成
  const newSpreadsheet = SpreadsheetApp.create(fileName);
  Logger.log('新規スプレッドシート作成: ' + fileName);

  // 保存先フォルダに移動
  const file = DriveApp.getFileById(newSpreadsheet.getId());
  file.moveTo(quoteFolder);
  Logger.log('スプレッドシートを保存先フォルダに移動: ' + quoteFolder.getName());

  // デフォルトのシートを削除
  const defaultSheet = newSpreadsheet.getSheets()[0];
  if (defaultSheet.getName() === 'シート1' || defaultSheet.getName() === 'Sheet1') {
    // 後で削除（少なくとも1つのシートが必要なため）
  }

  return newSpreadsheet;
}

/**
 * テンプレートから新しいシートを追加
 * @param {Spreadsheet} targetSpreadsheet - 追加先スプレッドシート
 * @param {string} templateName - テンプレート名
 * @param {string} newSheetName - 新しいシート名
 * @return {Sheet} - 追加されたシート
 */
function addQuoteSheet(targetSpreadsheet, templateName, newSheetName) {
  const sheets = targetSpreadsheet.getSheets();

  // デフォルトシート以外の既存シートを取得
  const existingSheets = sheets.filter(sheet => {
    const name = sheet.getName();
    return name !== 'シート1' && name !== 'Sheet1';
  });

  let copiedSheet;

  if (existingSheets.length > 0) {
    // 既存シートがある場合：最新のシート（配列の最後）をコピー
    const latestSheet = existingSheets[existingSheets.length - 1];
    Logger.log('既存シートをコピー: ' + latestSheet.getName());

    copiedSheet = latestSheet.copyTo(targetSpreadsheet);
    copiedSheet.setName(newSheetName);

    // 古いデータをクリア
    clearSheetData(copiedSheet, templateName);
    Logger.log('シート追加（既存シートからコピー）: ' + newSheetName);
  } else {
    // 既存シートがない場合：テンプレートからコピー
    const templateSS = getTemplateSpreadsheet();
    const templateSheet = templateSS.getSheetByName(templateName);

    if (!templateSheet) {
      throw new Error('テンプレートシートが見つかりません: ' + templateName);
    }

    copiedSheet = templateSheet.copyTo(targetSpreadsheet);
    copiedSheet.setName(newSheetName);
    Logger.log('シート追加（テンプレートからコピー）: ' + newSheetName);
  }

  // デフォルトシートを削除（最初のシート追加後）
  sheets.forEach(sheet => {
    if (sheet.getName() === 'シート1' || sheet.getName() === 'Sheet1') {
      targetSpreadsheet.deleteSheet(sheet);
    }
  });

  return copiedSheet;
}

/**
 * シートの古いデータをクリア
 * @param {Sheet} sheet - 対象シート
 * @param {string} templateName - テンプレート名
 */
function clearSheetData(sheet, templateName) {
  Logger.log('古いデータをクリア開始: ' + templateName);

  // テンプレート名で設定を選択
  let config;
  if (templateName === CONFIG.TEMPLATE_SHEETS.UNDER_13 || templateName === '13未満') {
    config = CONFIG.TEMPLATE_CELLS_13;
  } else if (templateName === CONFIG.TEMPLATE_SHEETS.BETWEEN_14_20 || templateName === '14-20') {
    config = CONFIG.TEMPLATE_CELLS_14_20;
  } else if (templateName === CONFIG.TEMPLATE_SHEETS.OVER_20 || templateName === '20以上') {
    config = CONFIG.TEMPLATE_CELLS_20_OVER;
  } else {
    throw new Error('不明なテンプレート名: ' + templateName);
  }

  // 固定情報をクリア
  sheet.getRange(config.クライアント名).clearContent();
  sheet.getRange(config.品名).clearContent();
  sheet.getRange(config.仕様1).clearContent();
  sheet.getRange(config.仕様2).clearContent();
  sheet.getRange(config.仕様3).clearContent();
  sheet.getRange(config.仕様4).clearContent();
  sheet.getRange(config.担当).clearContent();
  sheet.getRange(config.PROJECT_JOB_ID).clearContent();
  sheet.getRange(config.生成日時).clearContent();
  sheet.getRange(config.端数調整).clearContent();

  // 明細データをクリア
  if (templateName === CONFIG.TEMPLATE_SHEETS.UNDER_13 || templateName === '13未満') {
    // 13未満: 18-30行
    clearDetailRows(sheet, config.明細開始行, config.明細終了行, config.列);
  } else if (templateName === CONFIG.TEMPLATE_SHEETS.BETWEEN_14_20 || templateName === '14-20') {
    // 14-20: 18-37行
    clearDetailRows(sheet, config.明細開始行, config.明細終了行, config.列);
  } else {
    // 20以上: ページ1（18-37行）+ ページ2（39-62行）
    clearDetailRows(sheet, config.ページ1.明細開始行, config.ページ1.明細終了行, config.列);
    clearDetailRows(sheet, config.ページ2.明細開始行, config.ページ2.明細終了行, config.列);
  }

  Logger.log('古いデータをクリア完了');
}

/**
 * 明細行の範囲をクリア
 * @param {Sheet} sheet - 対象シート
 * @param {number} startRow - 開始行
 * @param {number} endRow - 終了行
 * @param {Object} columns - 列設定
 */
function clearDetailRows(sheet, startRow, endRow, columns) {
  for (let row = startRow; row <= endRow; row++) {
    sheet.getRange(row, 1).clearContent(); // A列: 行番号
    sheet.getRange(columns.作業項目 + row).clearContent();
    sheet.getRange(columns.作業詳細 + row).clearContent();
    sheet.getRange(columns.単価 + row).clearContent();
    sheet.getRange(columns.数量 + row).clearContent();
    sheet.getRange(columns.単位 + row).clearContent();
    // 金額は結合セル（G:H）の先頭をクリア
    const amountCell = columns.金額.split(':')[0] + row;
    sheet.getRange(amountCell).clearContent();
  }
}

/**
 * 見積データをシートに埋め込み
 * @param {Sheet} sheet - 対象シート
 * @param {Object} jobData - Jobデータ
 * @param {Array} details - Details配列
 * @param {string} templateName - テンプレート名
 * @param {Date} generatedDate - 生成日時
 */
function fillQuoteData(sheet, jobData, details, templateName, generatedDate) {
  Logger.log('=== fillQuoteData 開始 ===');
  Logger.log('テンプレート: "' + templateName + '"');
  Logger.log('明細数: ' + details.length);

  // デバッグ：最初の明細の内容を確認
  if (details.length > 0) {
    Logger.log('=== 最初の明細の内容 ===');
    Logger.log(JSON.stringify(details[0], null, 2));
  }

  // テンプレート名で設定を選択
  let config;
  if (templateName === CONFIG.TEMPLATE_SHEETS.UNDER_13 || templateName === '13未満') {
    config = CONFIG.TEMPLATE_CELLS_13;
    Logger.log('13未満テンプレートの設定を使用');
  } else if (templateName === CONFIG.TEMPLATE_SHEETS.BETWEEN_14_20 || templateName === '14-20') {
    config = CONFIG.TEMPLATE_CELLS_14_20;
    Logger.log('14-20テンプレートの設定を使用');
  } else if (templateName === CONFIG.TEMPLATE_SHEETS.OVER_20 || templateName === '20以上') {
    config = CONFIG.TEMPLATE_CELLS_20_OVER;
    Logger.log('20以上テンプレートの設定を使用');
  } else {
    throw new Error('不明なテンプレート名: ' + templateName);
  }

  if (!config) {
    throw new Error('設定が見つかりません。テンプレート名: ' + templateName);
  }

  Logger.log('設定取得完了');

  // 固定情報を設定
  Logger.log('固定情報を設定開始');
  sheet.getRange(config.クライアント名).setValue(jobData.クライアント名);
  Logger.log('クライアント名設定完了: ' + jobData.クライアント名);

  sheet.getRange(config.品名).setValue(jobData.品名);
  Logger.log('品名設定完了: ' + jobData.品名);

  sheet.getRange(config.仕様1).setValue(jobData.仕様1);
  sheet.getRange(config.仕様2).setValue(jobData.仕様2);
  sheet.getRange(config.仕様3).setValue(jobData.仕様3);
  sheet.getRange(config.仕様4).setValue(jobData.仕様4);
  Logger.log('仕様設定完了');

  sheet.getRange(config.担当).setValue(jobData.担当);
  Logger.log('担当設定完了: ' + jobData.担当);

  sheet.getRange(config.PROJECT_JOB_ID).setValue(jobData.projectId + '-' + jobData.jobId);
  Logger.log('PROJECT_JOB_ID設定完了: ' + jobData.projectId + '-' + jobData.jobId);

  // 生成日時を設定
  const formattedDateTime = Utilities.formatDate(generatedDate, CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm');
  sheet.getRange(config.生成日時).setValue(formattedDateTime);
  Logger.log('生成日時設定完了: ' + formattedDateTime);

  // 明細を設定
  Logger.log('明細設定開始');
  if (templateName === CONFIG.TEMPLATE_SHEETS.UNDER_13 || templateName === '13未満') {
    // 13未満テンプレート
    fillDetails13(sheet, details, config);
  } else if (templateName === CONFIG.TEMPLATE_SHEETS.BETWEEN_14_20 || templateName === '14-20') {
    // 14-20テンプレート
    fillDetails14_20(sheet, details, config);
  } else {
    // 20以上テンプレート
    fillDetails20Over(sheet, details, config);
  }
  Logger.log('=== fillQuoteData 完了 ===');
}

/**
 * 明細を設定（13未満テンプレート）
 */
function fillDetails13(sheet, details, config) {
  const startRow = config.明細開始行;
  Logger.log('fillDetails13: 開始行=' + startRow + ', 明細数=' + details.length);

  let rowOffset = 0;  // 通常明細の行オフセット

  details.forEach((detail, index) => {
    Logger.log(`明細${index + 1}: 作業項目=${detail.作業項目}, 金額=${detail.金額}`);

    // 作業項目が「端数調整」の場合は特別処理
    if (detail.作業項目 === '端数調整') {
      Logger.log('  → 端数調整を検出: G33:H33に金額を設定');
      sheet.getRange(config.端数調整).setValue(detail.金額);
      return;  // 通常の明細行としては表示しない
    }

    // 通常の明細処理
    if (rowOffset >= 13) return; // 13行まで

    const row = startRow + rowOffset;
    Logger.log(`  → 通常明細: 行=${row}`);

    sheet.getRange(row, 1).setValue(detail.行番号); // A列: 行番号
    sheet.getRange(config.列.作業項目 + row).setValue(detail.作業項目);
    sheet.getRange(config.列.作業詳細 + row).setValue(detail.作業詳細);
    sheet.getRange(config.列.単価 + row).setValue(detail.単価);
    sheet.getRange(config.列.数量 + row).setValue(detail.数量);
    sheet.getRange(config.列.単位 + row).setValue(detail.単位);

    // 金額は結合セル（G:H）の先頭に設定
    const amountCell = config.列.金額.split(':')[0] + row; // G列
    sheet.getRange(amountCell).setValue(detail.金額);
    Logger.log(`  → セル${amountCell}に金額${detail.金額}を設定`);

    rowOffset++;  // 通常明細の場合のみ行を進める
  });

  Logger.log('fillDetails13: 完了');
}

/**
 * 明細を設定（14-20テンプレート）
 */
function fillDetails14_20(sheet, details, config) {
  const startRow = config.明細開始行;
  Logger.log('fillDetails14_20: 開始行=' + startRow + ', 明細数=' + details.length);

  let rowOffset = 0;  // 通常明細の行オフセット

  details.forEach((detail, index) => {
    Logger.log(`明細${index + 1}: 作業項目=${detail.作業項目}, 金額=${detail.金額}`);

    // 作業項目が「端数調整」の場合は特別処理
    if (detail.作業項目 === '端数調整') {
      Logger.log('  → 端数調整を検出: G40:H40に金額を設定');
      sheet.getRange(config.端数調整).setValue(detail.金額);
      return;  // 通常の明細行としては表示しない
    }

    // 通常の明細処理
    if (rowOffset >= 20) return; // 20行まで

    const row = startRow + rowOffset;
    Logger.log(`  → 通常明細: 行=${row}`);

    sheet.getRange(row, 1).setValue(detail.行番号); // A列: 行番号
    sheet.getRange(config.列.作業項目 + row).setValue(detail.作業項目);
    sheet.getRange(config.列.作業詳細 + row).setValue(detail.作業詳細);
    sheet.getRange(config.列.単価 + row).setValue(detail.単価);
    sheet.getRange(config.列.数量 + row).setValue(detail.数量);
    sheet.getRange(config.列.単位 + row).setValue(detail.単位);

    // 金額は結合セル（G:H）の先頭に設定
    const amountCell = config.列.金額.split(':')[0] + row; // G列
    sheet.getRange(amountCell).setValue(detail.金額);
    Logger.log(`  → セル${amountCell}に金額${detail.金額}を設定`);

    rowOffset++;  // 通常明細の場合のみ行を進める
  });

  Logger.log('fillDetails14_20: 完了');
}

/**
 * 明細を設定（20以上テンプレート）
 */
function fillDetails20Over(sheet, details, config) {
  Logger.log('fillDetails20Over: 開始, 明細数=' + details.length);

  // 端数調整を除外した通常明細のみを配列化
  const normalDetails = [];
  let hasukasuchousei = null;

  details.forEach((detail) => {
    if (detail.作業項目 === '端数調整') {
      Logger.log('端数調整を検出: 金額=' + detail.金額);
      hasukasuchousei = detail;
    } else {
      normalDetails.push(detail);
    }
  });

  // 端数調整があれば専用セルに設定
  if (hasukasuchousei) {
    sheet.getRange(config.端数調整).setValue(hasukasuchousei.金額);
    Logger.log('  → G65:H65に端数調整を設定');
  }

  let detailIndex = 0;

  // ページ1（18-37行、20件）
  const page1Start = config.ページ1.明細開始行;
  const page1End = config.ページ1.明細終了行;
  const page1Count = page1End - page1Start + 1;

  for (let i = 0; i < page1Count && detailIndex < normalDetails.length; i++) {
    const detail = normalDetails[detailIndex];
    const row = page1Start + i;

    sheet.getRange(row, 1).setValue(detail.行番号);
    sheet.getRange(config.列.作業項目 + row).setValue(detail.作業項目);
    sheet.getRange(config.列.作業詳細 + row).setValue(detail.作業詳細);
    sheet.getRange(config.列.単価 + row).setValue(detail.単価);
    sheet.getRange(config.列.数量 + row).setValue(detail.数量);
    sheet.getRange(config.列.単位 + row).setValue(detail.単位);

    const amountCell = config.列.金額.split(':')[0] + row;
    sheet.getRange(amountCell).setValue(detail.金額);

    detailIndex++;
  }

  // ページ2（39-62行、24件）
  const page2Start = config.ページ2.明細開始行;
  const page2End = config.ページ2.明細終了行;
  const page2Count = page2End - page2Start + 1;

  for (let i = 0; i < page2Count && detailIndex < normalDetails.length; i++) {
    const detail = normalDetails[detailIndex];
    const row = page2Start + i;

    sheet.getRange(row, 1).setValue(detail.行番号);
    sheet.getRange(config.列.作業項目 + row).setValue(detail.作業項目);
    sheet.getRange(config.列.作業詳細 + row).setValue(detail.作業詳細);
    sheet.getRange(config.列.単価 + row).setValue(detail.単価);
    sheet.getRange(config.列.数量 + row).setValue(detail.数量);
    sheet.getRange(config.列.単位 + row).setValue(detail.単位);

    const amountCell = config.列.金額.split(':')[0] + row;
    sheet.getRange(amountCell).setValue(detail.金額);

    detailIndex++;
  }

  Logger.log('fillDetails20Over: 完了');
}

/**
 * シートをPDFとしてエクスポート
 * @param {Spreadsheet} spreadsheet - スプレッドシート
 * @param {Sheet} sheet - シート
 * @param {Object} jobData - Jobデータ
 * @param {Date} generatedDate - 生成日時
 * @return {File} - PDFファイル
 */
function exportSheetToPDF(spreadsheet, sheet, jobData, generatedDate) {
  const folderName = `${jobData.projectId}-${jobData.jobId}`;
  const pdfFolder = getPDFFolder(folderName);
  const pdfFileName = `${jobData.projectId}-${jobData.jobId}_${Utilities.formatDate(generatedDate, CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss')}.pdf`;

  // 既存の同名PDFファイルがあれば削除
  const existingPDFs = pdfFolder.getFilesByName(pdfFileName);
  if (existingPDFs.hasNext()) {
    const existingPDF = existingPDFs.next();
    Logger.log('既存のPDFファイルを削除: ' + pdfFileName);
    existingPDF.setTrashed(true);
  }

  const pdfBlob = generatePDFBlob(spreadsheet, sheet, generatedDate);
  const pdfFile = pdfFolder.createFile(pdfBlob.setName(pdfFileName));

  Logger.log('PDF生成完了: ' + pdfFileName);
  Logger.log('PDF保存先: ' + pdfFolder.getName());

  return pdfFile;
}

/**
 * PDFBlobを生成
 * @param {Spreadsheet} spreadsheet - スプレッドシート
 * @param {Sheet} sheet - シート
 * @param {Date} generatedDate - 生成日時
 * @return {Blob} - PDFデータ
 */
function generatePDFBlob(spreadsheet, sheet, generatedDate) {
  const sheetId = sheet.getSheetId();
  const spreadsheetId = spreadsheet.getId();

  // 生成日時をフォーマット
  const formattedDateTime = Utilities.formatDate(generatedDate, CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm');

  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export' +
    '?exportFormat=pdf' +
    '&format=pdf' +
    '&size=A4' +
    '&portrait=true' +
    '&fitw=true' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=false' +  // ページ番号を無効化
    '&gridlines=false' +
    '&fzr=false' +
    '&horizontal_alignment=CENTER' +
    '&vertical_alignment=TOP' +
    '&gid=' + sheetId;

  Logger.log('PDF URL: ' + url);

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });

  return response.getBlob();
}

/**
 * 見積URLをJobsテーブルに保存
 */
function saveQuoteUrls(rowIndex, spreadsheetUrl, pdfUrl, timestamp) {
  const ss = getAppSheetSpreadsheet();
  const jobsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.JOBS);

  jobsSheet.getRange(rowIndex, CONFIG.JOBS_COLUMNS.見積書スプレッドシートURL).setValue(spreadsheetUrl);
  jobsSheet.getRange(rowIndex, CONFIG.JOBS_COLUMNS.見積書PDF_URL).setValue(pdfUrl);
  jobsSheet.getRange(rowIndex, CONFIG.JOBS_COLUMNS.最終発行日時).setValue(timestamp);

  Logger.log('URLを保存しました');
}
