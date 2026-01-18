/**
 * デバッグ用テスト関数
 */

/**
 * Jobsシートのデータ構造を確認
 */
function debugJobsSheet() {
  const ss = getAppSheetSpreadsheet();
  const jobsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.JOBS);

  // ヘッダー行を取得（1行目、30列分）
  const headers = jobsSheet.getRange(1, 1, 1, 30).getValues()[0];
  Logger.log('=== ヘッダー行 ===');
  headers.forEach((header, index) => {
    if (header) {
      Logger.log(`列${index + 1}（${String.fromCharCode(65 + index)}列）: ${header}`);
    }
  });

  // 2行目のデータを取得
  Logger.log('\n=== 2行目のデータ ===');
  const row2 = jobsSheet.getRange(2, 1, 1, 30).getValues()[0];
  row2.forEach((value, index) => {
    if (value) {
      Logger.log(`列${index + 1}（${String.fromCharCode(65 + index)}列）: ${value} (型: ${typeof value})`);
    }
  });

  // Job IDの列番号を確認
  const jobIdColumnIndex = CONFIG.JOBS_COLUMNS.JOB_ID;
  const jobIdValue = row2[jobIdColumnIndex - 1];
  Logger.log(`\n=== Job ID ===`);
  Logger.log(`設定上のJob ID列番号: ${jobIdColumnIndex}`);
  Logger.log(`実際のJob ID値: "${jobIdValue}" (型: ${typeof jobIdValue})`);
}

/**
 * 特定のJob IDを検索してみる
 */
function testFindJob() {
  const testJobId = '001';  // 文字列として検索
  Logger.log(`Job ID "${testJobId}" を検索中...`);

  const jobData = getJobData(testJobId);

  if (jobData) {
    Logger.log('✓ Job が見つかりました:');
    Logger.log(JSON.stringify(jobData, null, 2));
  } else {
    Logger.log('✗ Job が見つかりませんでした');

    // 全てのJob IDをリスト表示
    const ss = getAppSheetSpreadsheet();
    const jobsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.JOBS);
    const data = jobsSheet.getDataRange().getValues();

    Logger.log('\n=== 存在する全てのJob ID ===');
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const jobId = row[CONFIG.JOBS_COLUMNS.JOB_ID - 1];
      if (jobId) {
        Logger.log(`行${i + 1}: "${jobId}" (型: ${typeof jobId})`);
      }
    }
  }
}

/**
 * Detailsシートのデバッグ
 */
function debugDetailsSheet() {
  const ss = getAppSheetSpreadsheet();
  const detailsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DETAILS);

  Logger.log('=== Detailsシート構造 ===');

  // ヘッダー行を取得
  const headers = detailsSheet.getRange(1, 1, 1, 20).getValues()[0];
  headers.forEach((header, index) => {
    if (header) {
      Logger.log(`列${index + 1}（${String.fromCharCode(65 + index)}列）: ${header}`);
    }
  });

  Logger.log('\n=== 設定上のJob ID列番号 ===');
  Logger.log('CONFIG.DETAILS_COLUMNS.JOB_ID = ' + CONFIG.DETAILS_COLUMNS.JOB_ID);

  Logger.log('\n=== 最初の5行のデータ ===');
  const data = detailsSheet.getRange(2, 1, Math.min(5, detailsSheet.getLastRow() - 1), 20).getValues();
  data.forEach((row, rowIndex) => {
    Logger.log(`\n--- 行${rowIndex + 2} ---`);
    row.forEach((cell, colIndex) => {
      if (cell) {
        Logger.log(`  列${colIndex + 1}（${String.fromCharCode(65 + colIndex)}列）: ${cell}`);
      }
    });
  });

  Logger.log('\n=== Job ID "001" の明細を検索 ===');
  const allData = detailsSheet.getDataRange().getValues();
  let found = 0;
  for (let i = 1; i < allData.length; i++) {
    const jobId = allData[i][CONFIG.DETAILS_COLUMNS.JOB_ID - 1];
    Logger.log(`行${i + 1}: Job ID列の値 = "${jobId}" (型: ${typeof jobId})`);
    if (String(jobId) === '001') {
      found++;
      Logger.log(`  ✓ 見つかりました！`);
    }
  }
  Logger.log(`\n合計 ${found} 件の明細が見つかりました`);
}
