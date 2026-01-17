/**
 * PDF出力スクリプト
 * 見積書スプレッドシートをPDFに変換
 */

/**
 * スプレッドシートをPDFとしてエクスポート
 * @param {Spreadsheet} spreadsheet - 対象スプレッドシート
 * @param {string} quoteNumber - 見積番号（ファイル名に使用）
 * @return {File} - 生成されたPDFファイル
 */
function exportToPDF(spreadsheet, quoteNumber) {
  try {
    Logger.log('PDF出力開始: ' + quoteNumber);

    const sheet = spreadsheet.getSheets()[0];
    const pdfFolder = getPDFFolder();

    // PDFファイル名を生成
    const pdfFileName = CONFIG.PDF.FILE_PREFIX + quoteNumber + '.pdf';

    // スプレッドシートをPDFとして出力
    const pdfBlob = generatePDFBlob(spreadsheet, sheet);

    // 既存のPDFファイルがあれば削除
    const existingFiles = pdfFolder.getFilesByName(pdfFileName);
    while (existingFiles.hasNext()) {
      const file = existingFiles.next();
      file.setTrashed(true);
    }

    // PDFファイルを保存
    const pdfFile = pdfFolder.createFile(pdfBlob.setName(pdfFileName));

    Logger.log('PDF出力完了: ' + pdfFile.getUrl());
    return pdfFile;

  } catch (error) {
    Logger.log('PDF出力エラー: ' + error.message);
    throw error;
  }
}

/**
 * スプレッドシートからPDF Blobを生成
 * @param {Spreadsheet} spreadsheet - 対象スプレッドシート
 * @param {Sheet} sheet - 対象シート
 * @return {Blob} - PDF Blob
 */
function generatePDFBlob(spreadsheet, sheet) {
  const sheetId = sheet.getSheetId();
  const spreadsheetId = spreadsheet.getId();

  // PDF出力のURLパラメータ
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export' +
    '?exportFormat=pdf' +
    '&format=pdf' +
    '&size=A4' +                    // 用紙サイズ
    '&portrait=true' +              // 縦向き
    '&fitw=true' +                  // 幅を用紙に合わせる
    '&sheetnames=false' +           // シート名を非表示
    '&printtitle=false' +           // スプレッドシート名を非表示
    '&pagenumbers=false' +          // ページ番号を非表示
    '&gridlines=false' +            // グリッド線を非表示
    '&fzr=false' +                  // 固定行を繰り返さない
    '&gid=' + sheetId;              // シートID

  // OAuth認証トークンを取得
  const token = ScriptApp.getOAuthToken();

  // PDFをフェッチ
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  return response.getBlob();
}

/**
 * 複数の見積書をまとめてPDF化
 * @param {Array} spreadsheetUrls - スプレッドシートURLの配列
 * @return {Array} - 生成されたPDFファイルの配列
 */
function batchExportToPDF(spreadsheetUrls) {
  const pdfFiles = [];

  spreadsheetUrls.forEach((url, index) => {
    try {
      const spreadsheet = SpreadsheetApp.openByUrl(url);
      const quoteNumber = spreadsheet.getName().replace('見積書_', '');
      const pdfFile = exportToPDF(spreadsheet, quoteNumber);
      pdfFiles.push(pdfFile);

      Logger.log(`${index + 1}/${spreadsheetUrls.length} 完了: ${quoteNumber}`);

    } catch (error) {
      Logger.log(`エラー (${url}): ${error.message}`);
    }
  });

  return pdfFiles;
}

/**
 * PDFをメールで送信
 * @param {File} pdfFile - PDFファイル
 * @param {string} toEmail - 送信先メールアドレス
 * @param {string} subject - メール件名
 * @param {string} body - メール本文
 */
function sendPDFByEmail(pdfFile, toEmail, subject, body) {
  try {
    MailApp.sendEmail({
      to: toEmail,
      subject: subject || '見積書を送付いたします',
      body: body || '見積書をPDFで送付いたします。\nご確認のほどよろしくお願いいたします。',
      attachments: [pdfFile.getBlob()]
    });

    Logger.log('メール送信完了: ' + toEmail);

  } catch (error) {
    Logger.log('メール送信エラー: ' + error.message);
    throw error;
  }
}

/**
 * AppSheet用：PDFのダイレクトリンクを取得
 * AppSheetでPDFを表示するために使用
 * @param {File} pdfFile - PDFファイル
 * @return {string} - ダイレクトリンクURL
 */
function getPDFDirectLink(pdfFile) {
  // ファイルを「リンクを知っている全員が閲覧可」に設定
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // ダイレクトリンクを生成
  const fileId = pdfFile.getId();
  return `https://drive.google.com/uc?export=download&id=${fileId}`;
}

/**
 * PDFプレビューリンクを取得
 * @param {File} pdfFile - PDFファイル
 * @return {string} - プレビューURL
 */
function getPDFPreviewLink(pdfFile) {
  const fileId = pdfFile.getId();
  return `https://drive.google.com/file/d/${fileId}/view`;
}
