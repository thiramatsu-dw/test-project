/**
 * SetupSheet.js
 * スプレッドシートの初期設定を行う
 * 初回実行時に一度だけ実行してください
 */

/**
 * スプレッドシートに必要なシートとヘッダーを初期設定する
 */
function setupSpreadsheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  setupStoreConfigSheet(ss);
  setupUploadHistorySheet(ss);

  SpreadsheetApp.flush();
  Logger.log('スプレッドシートのセットアップが完了しました。');
  showSetupInstructions();
}

/**
 * 店舗設定シートを作成・初期化する
 */
function setupStoreConfigSheet(ss) {
  var sheetName = CONFIG.SHEETS.STORE_CONFIG;
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(sheetName + ' シートを作成しました。');
  }

  // ヘッダー行の設定
  var headers = [
    '店舗名',
    'GBPアカウントID\n(accounts/XXXXXX)',
    'GBPロケーションID\n(locations/XXXXXX)',
    'Driveフォルダ ID',
    '写真カテゴリ\n(ADDITIONAL等)',
    '最終アップロード日時',
    'ステータス\n(有効/無効)'
  ];

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setWrap(true);
  sheet.setRowHeight(1, 50);

  // 列幅の調整
  sheet.setColumnWidth(1, 150); // 店舗名
  sheet.setColumnWidth(2, 200); // アカウントID
  sheet.setColumnWidth(3, 200); // ロケーションID
  sheet.setColumnWidth(4, 220); // フォルダID
  sheet.setColumnWidth(5, 150); // カテゴリ
  sheet.setColumnWidth(6, 180); // 最終アップロード
  sheet.setColumnWidth(7, 100); // ステータス

  // データ入力例（2行目）
  var exampleData = [
    '渋谷店',
    'accounts/123456789',
    'locations/987654321',
    '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms', // DriveフォルダのID
    'ADDITIONAL',
    '',
    '有効'
  ];
  sheet.getRange(2, 1, 1, exampleData.length).setValues([exampleData]);
  sheet.getRange(2, 1, 1, exampleData.length).setBackground('#e8f0fe');

  // ステータス列のドロップダウン検証
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['有効', '無効'], true)
    .build();
  sheet.getRange(2, CONFIG.STORE_COLS.STATUS + 1, sheet.getMaxRows() - 1, 1)
    .setDataValidation(statusRule);

  // カテゴリ列のドロップダウン検証
  var categoryValues = Object.values(CONFIG.PHOTO_CATEGORIES);
  var categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categoryValues, true)
    .build();
  sheet.getRange(2, CONFIG.STORE_COLS.CATEGORY + 1, sheet.getMaxRows() - 1, 1)
    .setDataValidation(categoryRule);

  // 行のフリーズ
  sheet.setFrozenRows(1);
}

/**
 * アップロード履歴シートを作成・初期化する
 */
function setupUploadHistorySheet(ss) {
  var sheetName = CONFIG.SHEETS.UPLOAD_HISTORY;
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(sheetName + ' シートを作成しました。');
  }

  // ヘッダー行の設定
  var headers = ['日時', '店舗名', 'ファイル名', 'DriveファイルID', '結果', 'メッセージ'];
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅の調整
  sheet.setColumnWidth(1, 180); // 日時
  sheet.setColumnWidth(2, 150); // 店舗名
  sheet.setColumnWidth(3, 200); // ファイル名
  sheet.setColumnWidth(4, 220); // ファイルID
  sheet.setColumnWidth(5, 80);  // 結果
  sheet.setColumnWidth(6, 300); // メッセージ

  // 行のフリーズ
  sheet.setFrozenRows(1);
}

/**
 * セットアップ後の操作手順をログ表示する
 */
function showSetupInstructions() {
  var instructions = [
    '=== セットアップ完了 ===',
    '',
    '【次の手順で設定を行ってください】',
    '',
    '1. Google ビジネスプロフィール のアカウントID・ロケーションIDの確認:',
    '   - https://business.google.com/ にアクセス',
    '   - 管理している店舗の URL から取得',
    '   - 例: .../accounts/123456789/locations/987654321',
    '',
    '2. 「店舗設定」シートへの入力:',
    '   - 各店舗の情報を入力してください（例データの2行目を参考に）',
    '   - GBPアカウントID: "accounts/XXXXXX" 形式',
    '   - GBPロケーションID: "locations/XXXXXX" 形式',
    '   - DriveフォルダID: アップロードする画像を置くフォルダのID',
    '   - 写真カテゴリ: ADDITIONAL（その他）、COVER（カバー）など',
    '',
    '3. Driveフォルダの準備:',
    '   - 各店舗用のGoogle Driveフォルダを作成',
    '   - フォルダに画像ファイル（JPG/PNG）を配置',
    '   - フォルダIDをシートに入力',
    '',
    '4. 定期実行の設定:',
    '   - setupDailyTrigger() を実行して毎日自動実行を設定',
    '',
    '5. 手動テスト:',
    '   - uploadAllStoreImages() を実行してテストアップロード'
  ].join('\n');

  Logger.log(instructions);
}

/**
 * GBPのアカウント一覧を取得してログに表示する（ID確認用ユーティリティ）
 */
function listGBPAccounts() {
  var result = GBPApi.listAccounts();
  if (result.accounts) {
    Logger.log('=== GBP アカウント一覧 ===');
    result.accounts.forEach(function(account) {
      Logger.log('名前: ' + account.accountName + ' | ID: ' + account.name);
    });
  } else {
    Logger.log('アカウントが見つかりません: ' + JSON.stringify(result));
  }
}

/**
 * 指定アカウントのロケーション一覧を取得してログに表示する（ID確認用ユーティリティ）
 * @param {string} accountName - "accounts/XXXXXX" 形式のアカウント名
 */
function listGBPLocations(accountName) {
  if (!accountName) {
    Logger.log('使用方法: listGBPLocations("accounts/XXXXXX")');
    return;
  }
  var result = GBPApi.listLocations(accountName);
  if (result.locations) {
    Logger.log('=== ロケーション一覧 (' + accountName + ') ===');
    result.locations.forEach(function(location) {
      Logger.log('店舗名: ' + location.locationName + ' | ID: ' + location.name);
    });
  } else {
    Logger.log('ロケーションが見つかりません: ' + JSON.stringify(result));
  }
}
