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
    let details = getDetails(jobId);

    if (details.length === 0) {
      throw new Error('明細が見つかりません');
    }

    Logger.log('明細数（取得時）: ' + details.length);

    // 3. 作業項目でソートして行番号を振り直す
    details = sortAndReorderDetails(details);
    Logger.log('明細数（ソート後）: ' + details.length);

    // 4. テンプレートを選択（13未満 / 14-20 / 20以上）
    let templateName;
    if (details.length <= 13) {
      templateName = CONFIG.TEMPLATE_SHEETS.UNDER_13;
    } else if (details.length <= 20) {
      templateName = CONFIG.TEMPLATE_SHEETS.BETWEEN_14_20;
    } else {
      templateName = CONFIG.TEMPLATE_SHEETS.OVER_20;
    }

    Logger.log('使用テンプレート: ' + templateName);

    // 5. スプレッドシートを作成または取得
    const spreadsheet = getOrCreateQuoteSpreadsheet(jobData);

    // 6. 新しいシートを追加（発行日時が名前）
    const now = new Date();
    const sheetName = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const newSheet = addQuoteSheet(spreadsheet, templateName, sheetName);

    // 7. データを埋め込み
    fillQuoteData(newSheet, jobData, details, templateName, now);

    // 8. スプレッドシートへの書き込みを確実に反映させる
    Logger.log('スプレッドシートへの書き込みを反映中...');
    SpreadsheetApp.flush();
    Logger.log('スプレッドシートへの書き込み完了');

    // 9. PDFを生成（ページ番号付き）
    const pdfFile = exportSheetToPDF(spreadsheet, newSheet, jobData, now);

    // 10. 生成したシートを直接開くURLを作成
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/edit#gid=${newSheet.getSheetId()}`;
    Logger.log('シートURL: ' + sheetUrl);

    // 11. JobsテーブルにURLを保存
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
 * 明細を作業項目でソートして行番号を振り直す
 * @param {Array} details - Details配列
 * @return {Array} - ソート済みDetails配列
 */
function sortAndReorderDetails(details) {
  Logger.log('=== 明細ソート開始 ===');

  // 端数調整とそれ以外を分離
  const normalDetails = [];
  let hasukasuchousei = null;

  details.forEach((detail) => {
    if (detail.作業項目 === '端数調整') {
      hasukasuchousei = detail;
      Logger.log('端数調整を検出: ' + JSON.stringify(detail));
    } else {
      normalDetails.push(detail);
    }
  });

  // 通常明細を作業項目でソート
  normalDetails.sort((a, b) => {
    const itemA = a.作業項目 || '';
    const itemB = b.作業項目 || '';
    return itemA.localeCompare(itemB, 'ja');
  });

  Logger.log('ソート後の作業項目順:');
  normalDetails.forEach((detail, index) => {
    Logger.log(`  ${index + 1}: ${detail.作業項目}`);
  });

  // 行番号を1から振り直す
  normalDetails.forEach((detail, index) => {
    detail.行番号 = index + 1;
  });

  // 端数調整を最後に追加
  const result = [...normalDetails];
  if (hasukasuchousei) {
    hasukasuchousei.行番号 = result.length + 1;
    result.push(hasukasuchousei);
    Logger.log('端数調整を最後に追加（行番号: ' + hasukasuchousei.行番号 + '）');
  }

  Logger.log('=== 明細ソート完了（全' + result.length + '件）===');
  return result;
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
  const templateSS = getTemplateSpreadsheet();
  const templateSheet = templateSS.getSheetByName(templateName);

  if (!templateSheet) {
    throw new Error('テンプレートシートが見つかりません: ' + templateName);
  }

  // テンプレートをコピー
  const copiedSheet = templateSheet.copyTo(targetSpreadsheet);
  copiedSheet.setName(newSheetName);

  Logger.log('シート追加: ' + newSheetName);

  // デフォルトシートを削除（最初のシート追加後）
  const sheets = targetSpreadsheet.getSheets();
  sheets.forEach(sheet => {
    if (sheet.getName() === 'シート1' || sheet.getName() === 'Sheet1') {
      targetSpreadsheet.deleteSheet(sheet);
    }
  });

  return copiedSheet;
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

  // 同じ作業項目のセルを結合
  mergeWorkItemCells(sheet, details, startRow, config.列.作業項目);

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

  // 同じ作業項目のセルを結合
  mergeWorkItemCells(sheet, details, startRow, config.列.作業項目);

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
  const page1Details = [];

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

    page1Details.push(detail);
    detailIndex++;
  }

  // ページ1の作業項目を結合
  mergeWorkItemCells(sheet, page1Details, page1Start, config.列.作業項目);

  // ページ2（39-62行、24件）
  const page2Start = config.ページ2.明細開始行;
  const page2End = config.ページ2.明細終了行;
  const page2Count = page2End - page2Start + 1;
  const page2Details = [];

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

    page2Details.push(detail);
    detailIndex++;
  }

  // ページ2の作業項目を結合
  mergeWorkItemCells(sheet, page2Details, page2Start, config.列.作業項目);

  Logger.log('fillDetails20Over: 完了');
}

/**
 * 同じ作業項目のセルを結合して上揃えにする
 * @param {Sheet} sheet - 対象シート
 * @param {Array} details - Details配列（端数調整を除く）
 * @param {number} startRow - 開始行
 * @param {string} columnLetter - 作業項目の列（例: 'B'）
 */
function mergeWorkItemCells(sheet, details, startRow, columnLetter) {
  if (details.length === 0) return;

  Logger.log('=== 作業項目セル結合開始 ===');

  let mergeStart = startRow;
  let currentItem = details[0].作業項目;
  let mergeCount = 1;

  for (let i = 1; i < details.length; i++) {
    const detail = details[i];

    if (detail.作業項目 === currentItem) {
      // 同じ作業項目なら結合範囲を拡大
      mergeCount++;
    } else {
      // 作業項目が変わったら、前の範囲を結合
      if (mergeCount > 1) {
        const range = sheet.getRange(`${columnLetter}${mergeStart}:${columnLetter}${mergeStart + mergeCount - 1}`);
        range.mergeVertically();
        range.setVerticalAlignment('top');
        Logger.log(`  ${currentItem}: ${columnLetter}${mergeStart}:${columnLetter}${mergeStart + mergeCount - 1} を結合`);
      }

      // 新しい作業項目の範囲を開始
      mergeStart = startRow + i;
      currentItem = detail.作業項目;
      mergeCount = 1;
    }
  }

  // 最後の範囲を結合
  if (mergeCount > 1) {
    const range = sheet.getRange(`${columnLetter}${mergeStart}:${columnLetter}${mergeStart + mergeCount - 1}`);
    range.mergeVertically();
    range.setVerticalAlignment('top');
    Logger.log(`  ${currentItem}: ${columnLetter}${mergeStart}:${columnLetter}${mergeStart + mergeCount - 1} を結合`);
  }

  Logger.log('=== 作業項目セル結合完了 ===');
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
