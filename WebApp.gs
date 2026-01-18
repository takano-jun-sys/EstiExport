/**
 * Web Appエントリーポイント
 * AppSheetから呼び出される
 */

/**
 * GETリクエストハンドラー
 * AppSheetから Job.ID をパラメータで受け取る
 *
 * URL例: https://script.google.com/macros/s/.../exec?jobId=001
 */
function doGet(e) {
  try {
    const jobId = e.parameter.jobId;

    if (!jobId) {
      return HtmlService.createHtmlOutput('エラー: Job IDが指定されていません');
    }

    Logger.log('見積書Export開始: Job ID = ' + jobId);

    // Jobデータを取得
    const jobData = getJobData(jobId);

    if (!jobData) {
      return HtmlService.createHtmlOutput('エラー: Job ID "' + jobId + '" が見つかりません');
    }

    // 初回かどうかを判定
    const isFirstTime = !jobData.見積書スプレッドシートURL;

    // 入力ダイアログHTMLを生成
    const template = HtmlService.createTemplateFromFile('DialogHTML');
    template.jobData = jobData;
    template.isFirstTime = isFirstTime;

    return template.evaluate()
      .setTitle('見積書 Export - ' + jobId)
      .setWidth(500)
      .setHeight(600);

  } catch (error) {
    Logger.log('doGetエラー: ' + error.message);
    return HtmlService.createHtmlOutput('エラーが発生しました: ' + error.message);
  }
}

/**
 * Jobデータを取得
 * @param {string} jobId - Job ID
 * @return {Object|null} - Jobデータ
 */
function getJobData(jobId) {
  const ss = getAppSheetSpreadsheet();
  const jobsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.JOBS);

  const data = jobsSheet.getDataRange().getValues();

  // ヘッダー行を除いて検索
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentJobId = row[CONFIG.JOBS_COLUMNS.JOB_ID - 1];

    if (currentJobId === jobId) {
      return {
        rowIndex: i + 1,  // 実際の行番号（1始まり）
        projectId: row[CONFIG.JOBS_COLUMNS.PROJECT_ID - 1],
        jobId: jobId,
        案件名: row[CONFIG.JOBS_COLUMNS.案件名 - 1],
        クライアント名: row[CONFIG.JOBS_COLUMNS.クライアント名 - 1] || '',
        品名: row[CONFIG.JOBS_COLUMNS.品名 - 1] || '',
        仕様1: row[CONFIG.JOBS_COLUMNS.仕様1 - 1] || '',
        仕様2: row[CONFIG.JOBS_COLUMNS.仕様2 - 1] || '',
        仕様3: row[CONFIG.JOBS_COLUMNS.仕様3 - 1] || '',
        仕様4: row[CONFIG.JOBS_COLUMNS.仕様4 - 1] || '',
        担当: row[CONFIG.JOBS_COLUMNS.担当 - 1] || '',
        見積書スプレッドシートURL: row[CONFIG.JOBS_COLUMNS.見積書スプレッドシートURL - 1] || '',
        見積書PDF_URL: row[CONFIG.JOBS_COLUMNS.見積書PDF_URL - 1] || '',
        最終発行日時: row[CONFIG.JOBS_COLUMNS.最終発行日時 - 1] || ''
      };
    }
  }

  return null;
}

/**
 * Details（明細）を取得
 * @param {string} jobId - Job ID
 * @return {Array} - Details配列
 */
function getDetails(jobId) {
  const ss = getAppSheetSpreadsheet();
  const detailsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DETAILS);

  const data = detailsSheet.getDataRange().getValues();
  const details = [];

  // ヘッダー行を除いて検索
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentJobId = row[CONFIG.DETAILS_COLUMNS.JOB_ID - 1];

    if (currentJobId === jobId) {
      details.push({
        明細ID: row[CONFIG.DETAILS_COLUMNS.明細ID - 1],
        jobId: jobId,
        枝番: row[CONFIG.DETAILS_COLUMNS.枝番 - 1],
        行番号: row[CONFIG.DETAILS_COLUMNS.行番号 - 1],
        作業項目: row[CONFIG.DETAILS_COLUMNS.作業項目 - 1] || '',
        担当部署: row[CONFIG.DETAILS_COLUMNS.担当部署 - 1] || '',
        数量: row[CONFIG.DETAILS_COLUMNS.数量 - 1] || 0,
        単位: row[CONFIG.DETAILS_COLUMNS.単位 - 1] || '',
        単価: row[CONFIG.DETAILS_COLUMNS.単価 - 1] || 0,
        メモ: row[CONFIG.DETAILS_COLUMNS.メモ - 1] || '',
        金額: row[CONFIG.DETAILS_COLUMNS.金額 - 1] || 0
      });
    }
  }

  // 枝番でソート
  details.sort((a, b) => a.枝番 - b.枝番);

  return details;
}

/**
 * Jobデータを更新（ダイアログから呼び出される）
 * @param {string} jobId - Job ID
 * @param {Object} formData - フォームデータ
 * @return {Object} - 結果
 */
function updateJobData(jobId, formData) {
  try {
    const ss = getAppSheetSpreadsheet();
    const jobsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.JOBS);

    // Jobを検索
    const jobData = getJobData(jobId);
    if (!jobData) {
      throw new Error('Job ID が見つかりません');
    }

    const rowIndex = jobData.rowIndex;

    // 各列を更新
    jobsSheet.getRange(rowIndex, CONFIG.JOBS_COLUMNS.クライアント名).setValue(formData.クライアント名);
    jobsSheet.getRange(rowIndex, CONFIG.JOBS_COLUMNS.品名).setValue(formData.品名);
    jobsSheet.getRange(rowIndex, CONFIG.JOBS_COLUMNS.仕様1).setValue(formData.仕様1);
    jobsSheet.getRange(rowIndex, CONFIG.JOBS_COLUMNS.仕様2).setValue(formData.仕様2);
    jobsSheet.getRange(rowIndex, CONFIG.JOBS_COLUMNS.仕様3).setValue(formData.仕様3);
    jobsSheet.getRange(rowIndex, CONFIG.JOBS_COLUMNS.仕様4).setValue(formData.仕様4);
    jobsSheet.getRange(rowIndex, CONFIG.JOBS_COLUMNS.担当).setValue(formData.担当);

    return { success: true };

  } catch (error) {
    Logger.log('updateJobDataエラー: ' + error.message);
    return { success: false, error: error.message };
  }
}
