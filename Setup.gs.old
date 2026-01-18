/**
 * セットアップスクリプト
 * 初期設定とシートの自動生成
 */

/**
 * 初期セットアップ：必要なシートとヘッダーを作成
 * GASエディタから手動で実行してください
 */
function setupSheets() {
  try {
    Logger.log('初期セットアップを開始します');

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // データシートを作成
    setupDataSheet(ss);

    // テンプレートシートを作成
    setupTemplateSheet(ss);

    // 設定シートを作成
    setupSettingsSheet(ss);

    Browser.msgBox('初期セットアップが完了しました！\\n\\n以下のシートが作成されました：\\n・見積データ\\n・見積書テンプレート\\n・設定\\n\\n次のステップ：\\n1. 「設定」シートで会社情報を編集\\n2. 「見積書テンプレート」をカスタマイズ（任意）\\n3. AppSheetでデータソースとして接続');

    Logger.log('初期セットアップが完了しました');

  } catch (error) {
    Logger.log('セットアップエラー: ' + error.message);
    Browser.msgBox('エラーが発生しました: ' + error.message);
    throw error;
  }
}

/**
 * データシートをセットアップ
 * @param {Spreadsheet} ss - スプレッドシート
 */
function setupDataSheet(ss) {
  let dataSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DATA);

  if (!dataSheet) {
    dataSheet = ss.insertSheet(CONFIG.SHEET_NAMES.DATA);
  }

  // ヘッダー行を作成
  const headers = [
    '見積番号',
    '日付',
    '顧客名',
    '郵便番号',
    '住所',
    '電話番号',
    '件名',
    '明細JSON',
    '備考',
    '消費税率',
    '有効期限',
    'ステータス',
    'スプレッドシートURL',
    'PDF URL'
  ];

  dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行のスタイル設定
  const headerRange = dataSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4A86E8');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // 列幅を調整
  dataSheet.setColumnWidth(1, 150);  // 見積番号
  dataSheet.setColumnWidth(2, 120);  // 日付
  dataSheet.setColumnWidth(3, 200);  // 顧客名
  dataSheet.setColumnWidth(4, 100);  // 郵便番号
  dataSheet.setColumnWidth(5, 300);  // 住所
  dataSheet.setColumnWidth(6, 120);  // 電話番号
  dataSheet.setColumnWidth(7, 200);  // 件名
  dataSheet.setColumnWidth(8, 400);  // 明細JSON
  dataSheet.setColumnWidth(9, 200);  // 備考
  dataSheet.setColumnWidth(10, 80);  // 消費税率
  dataSheet.setColumnWidth(11, 120); // 有効期限
  dataSheet.setColumnWidth(12, 100); // ステータス
  dataSheet.setColumnWidth(13, 300); // スプレッドシートURL
  dataSheet.setColumnWidth(14, 300); // PDF URL

  // サンプルデータを追加
  const sampleData = [
    [
      'QUOTE-20260117-001',
      new Date(),
      '株式会社サンプル商事',
      '100-0001',
      '東京都千代田区千代田1-1-1',
      '03-1111-2222',
      'Webシステム開発のお見積り',
      JSON.stringify([
        { item: 'Webサイト設計', quantity: 1, unit: '式', unitPrice: 200000, amount: 200000, note: 'レスポンシブ対応' },
        { item: 'フロントエンド開発', quantity: 1, unit: '式', unitPrice: 500000, amount: 500000, note: 'React使用' },
        { item: 'バックエンド開発', quantity: 1, unit: '式', unitPrice: 600000, amount: 600000, note: 'Node.js + PostgreSQL' }
      ]),
      '納期：契約後3ヶ月\n保守サポート：月額50,000円（任意）',
      0.10,
      new Date(new Date().getTime() + 30 * 24 * 60 * 60 * 1000),
      '未作成',
      '',
      ''
    ]
  ];

  dataSheet.getRange(2, 1, 1, headers.length).setValues(sampleData);

  // ステータス列にデータ検証を追加
  const statusRange = dataSheet.getRange(2, 12, 1000, 1);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未作成', '作成済み', 'エラー'], true)
    .build();
  statusRange.setDataValidation(statusRule);

  // 消費税率列の書式設定
  dataSheet.getRange(2, 10, 1000, 1).setNumberFormat('0.00');

  Logger.log('データシートのセットアップ完了');
}

/**
 * テンプレートシートをセットアップ
 * @param {Spreadsheet} ss - スプレッドシート
 */
function setupTemplateSheet(ss) {
  let templateSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATE);

  if (!templateSheet) {
    templateSheet = ss.insertSheet(CONFIG.SHEET_NAMES.TEMPLATE);
  }

  // シート全体の設定
  templateSheet.setColumnWidth(1, 40);   // A列: No
  templateSheet.setColumnWidth(2, 250);  // B列: 品名
  templateSheet.setColumnWidth(3, 60);   // C列: 数量
  templateSheet.setColumnWidth(4, 60);   // D列: 単位
  templateSheet.setColumnWidth(5, 100);  // E列: 単価
  templateSheet.setColumnWidth(6, 120);  // F列: 金額
  templateSheet.setColumnWidth(7, 150);  // G列: 備考

  // タイトル行
  templateSheet.getRange('A1:G1').merge();
  templateSheet.getRange('A1').setValue('御 見 積 書');
  templateSheet.getRange('A1').setFontSize(20);
  templateSheet.getRange('A1').setFontWeight('bold');
  templateSheet.getRange('A1').setHorizontalAlignment('center');

  // 見積情報エリア（右上）
  templateSheet.getRange('F2').setValue('見積番号:');
  templateSheet.getRange('F3').setValue('見積日:');
  templateSheet.getRange('F4').setValue('有効期限:');

  // 顧客情報エリア
  templateSheet.getRange('A6').setValue('顧客名:');
  templateSheet.getRange('A7').setValue('郵便番号:');
  templateSheet.getRange('A8').setValue('住所:');
  templateSheet.getRange('A9').setValue('電話番号:');

  // 件名
  templateSheet.getRange('A11').setValue('件名:');

  // 合計金額エリア
  templateSheet.getRange('F50').setValue('小計:');
  templateSheet.getRange('F51').setValue('消費税:');
  templateSheet.getRange('F52').setValue('合計:');
  templateSheet.getRange('F52').setFontWeight('bold');
  templateSheet.getRange('F52').setFontSize(12);

  // 明細ヘッダー
  const itemHeaders = ['No', '品名・サービス名', '数量', '単位', '単価', '金額', '備考'];
  templateSheet.getRange(13, 1, 1, 7).setValues([itemHeaders]);
  templateSheet.getRange(13, 1, 1, 7).setFontWeight('bold');
  templateSheet.getRange(13, 1, 1, 7).setBackground('#E8E8E8');
  templateSheet.getRange(13, 1, 1, 7).setHorizontalAlignment('center');

  // 明細サンプル行
  templateSheet.getRange(14, 1, 1, 7).setValues([
    [1, 'サンプル商品', 1, '個', 10000, 10000, '']
  ]);

  // 備考エリア
  templateSheet.getRange('A54').setValue('【備考】');
  templateSheet.getRange('A54').setFontWeight('bold');

  // 罫線を追加
  const borderRange = templateSheet.getRange('A1:G60');
  borderRange.setBorder(true, true, true, true, true, true);

  // 保護設定（テンプレートの構造を保護）
  const protection = templateSheet.protect().setDescription('テンプレートシート');
  protection.setWarningOnly(true);

  Logger.log('テンプレートシートのセットアップ完了');
}

/**
 * 設定シートをセットアップ
 * @param {Spreadsheet} ss - スプレッドシート
 */
function setupSettingsSheet(ss) {
  let settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);

  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(CONFIG.SHEET_NAMES.SETTINGS);
  }

  // デフォルトの会社情報を設定
  const defaultCompanyInfo = getDefaultCompanyInfo();
  saveCompanyInfo(defaultCompanyInfo);

  // 列幅を調整
  settingsSheet.setColumnWidth(1, 200);
  settingsSheet.setColumnWidth(2, 300);

  Logger.log('設定シートのセットアップ完了');
}

/**
 * トリガーを自動設定
 * 5分ごとにprocessNewQuotesを実行
 */
function setupTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processNewQuotes') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを作成
  ScriptApp.newTrigger('processNewQuotes')
    .timeBased()
    .everyMinutes(5)
    .create();

  Browser.msgBox('トリガーを設定しました。\\n\\n5分ごとに未作成の見積書を自動処理します。');
  Logger.log('トリガーのセットアップ完了');
}

/**
 * トリガーを削除
 */
function removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processNewQuotes') {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });

  Browser.msgBox(`${count}個のトリガーを削除しました。`);
  Logger.log('トリガーの削除完了');
}
