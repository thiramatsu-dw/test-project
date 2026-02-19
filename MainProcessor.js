/**
 * MainProcessor.js
 * 入稿フォルダを監視して商品情報を一括登録するメインオーケストレーター
 *
 * 【処理フロー】
 *  1. 入稿フォルダ内の新規ファイルを検索
 *  2. InputProcessor で入稿ファイルをパース
 *  3. 各明細行について:
 *     a. 店舗コードから GBP ロケーションを特定
 *     b. 画像ファイルを Drive から取得
 *     c. ProductsApi で商品ローカルポストを作成
 *  4. ExecutionLogger でログシートを出力
 *  5. 入稿ファイルを処理済み / エラーフォルダへ移動
 */

// ===== エントリーポイント（手動 or トリガーから呼ぶ） =====

/**
 * 入稿フォルダ内の全ファイルを処理するメイン関数
 */
function processInboxFiles() {
  Logger.log('=== 入稿処理開始: ' + new Date().toLocaleString('ja-JP') + ' ===');

  var folders = _getWorkFolders();
  var files = _listInboxFiles(folders.inbox);

  if (files.length === 0) {
    Logger.log('入稿フォルダに処理対象ファイルがありません。');
    return;
  }

  Logger.log('処理対象ファイル数: ' + files.length);

  files.forEach(function(file, idx) {
    Logger.log('\n--- ファイル ' + (idx + 1) + '/' + files.length + ': ' + file.getName() + ' ---');

    try {
      _processOneFile(file, folders);
    } catch (e) {
      Logger.log('予期せぬエラー: ' + e.message);
      _moveFile(file, folders.error);
    }

    // ファイル間でインターバル（API レートリミット対策）
    if (idx < files.length - 1) Utilities.sleep(2000);
  });

  Logger.log('\n=== 入稿処理完了 ===');
}

/**
 * 特定のファイルIDを指定して処理する（テスト・再処理用）
 * @param {string} fileId - 処理する Drive ファイルの ID
 */
function processFileById(fileId) {
  if (!fileId) {
    Logger.log('使用方法: processFileById("DriveファイルID")');
    return;
  }

  var file = DriveApp.getFileById(fileId);
  Logger.log('ファイルを処理します: ' + file.getName());

  var folders = _getWorkFolders();
  _processOneFile(file, folders);
}

// ===== 内部処理 =====

/**
 * 作業用フォルダ（入稿・処理済み・エラー・処理結果）を取得または作成する
 */
function _getWorkFolders() {
  var inbox;

  if (CONFIG.INBOX_FOLDER_ID) {
    inbox = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  } else {
    var rootFolders = DriveApp.getRootFolder().getFoldersByName(CONFIG.INBOX_FOLDER_NAME);
    inbox = rootFolders.hasNext()
      ? rootFolders.next()
      : DriveApp.getRootFolder().createFolder(CONFIG.INBOX_FOLDER_NAME);
  }

  return {
    inbox:     inbox,
    processed: _getOrCreateSubFolder(inbox, CONFIG.PROCESSED_FOLDER_NAME),
    error:     _getOrCreateSubFolder(inbox, CONFIG.ERROR_FOLDER_NAME),
    results:   _getOrCreateSubFolder(inbox, CONFIG.RESULTS_FOLDER_NAME)
  };
}

/**
 * 入稿フォルダ内の処理対象ファイル一覧を返す
 * （処理済み・エラー・処理結果 サブフォルダを除く）
 */
function _listInboxFiles(inboxFolder) {
  var skip = [
    CONFIG.PROCESSED_FOLDER_NAME,
    CONFIG.ERROR_FOLDER_NAME,
    CONFIG.RESULTS_FOLDER_NAME
  ];
  var files = [];
  var it = inboxFolder.getFiles();

  while (it.hasNext()) {
    var f = it.next();
    var mime = f.getMimeType();
    // スプレッドシート・Excel・CSV のみ対象
    if (
      mime === MimeType.GOOGLE_SHEETS ||
      mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      mime === 'application/vnd.ms-excel' ||
      mime === 'text/csv'
    ) {
      files.push(f);
    }
  }

  return files;
}

/**
 * 1ファイルの入稿処理を行う
 */
function _processOneFile(file, folders) {
  var fileName = file.getName();

  // ====== 1. ファイルのパース ======
  var parsed = InputProcessor.parse(file);
  var logger = ExecutionLogger.create(folders.results, fileName);
  var hasAnyError = false;

  // ヘッダーレベルのエラーがある場合
  if (parsed.errors.length > 0) {
    parsed.errors.forEach(function(err) {
      Logger.log('入稿ヘッダーエラー: ' + err);
      logger.addHeaderError(err, fileName);
    });

    if (!parsed.header) {
      // ヘッダーが読めない致命的エラー
      logger.finalize();
      _moveFile(file, folders.error);
      Logger.log('ヘッダーが読み取れないためスキップしました: ' + fileName);
      return;
    }
    hasAnyError = true;
  }

  var header = parsed.header;
  Logger.log('ヘッダー情報 | グループ: ' + header.businessGroupName +
    ' | アカウントID: ' + header.googleAccountId);

  // GBP アカウント名を組み立てる
  // googleAccountId が "accounts/XXXXX" 形式の場合はそのまま、数字のみの場合は prefix を付ける
  var accountName = _normalizeAccountName(header.googleAccountId);

  // ====== 2. 店舗コードマップの取得 ======
  var storeCodeMap = {};
  try {
    storeCodeMap = ProductsApi.buildStoreCodeMap(accountName);
    Logger.log('ロケーション取得完了: ' + Object.keys(storeCodeMap).length + ' 件');
  } catch (e) {
    var errMsg = 'GBP ロケーション取得エラー: ' + e.message;
    Logger.log(errMsg);
    logger.addHeaderError(errMsg, fileName);
    logger.finalize();
    _moveFile(file, folders.error);
    return;
  }

  // 画像ルートフォルダIDの解決
  var imageRootFolderId = _resolveImageRootFolderId();

  // ====== 3. 明細行の処理 ======
  parsed.details.forEach(function(item) {
    var detail = item.data;
    var rowNum  = item.rowNum;

    Logger.log('  行' + rowNum + ': ' + detail.businessName + ' / ' + detail.productName);

    // 行レベルのバリデーションエラー
    if (item.errors.length > 0) {
      var errMsg = item.errors.join(' / ');
      Logger.log('    バリデーションエラー: ' + errMsg);
      logger.addRow(header, detail, 'ERROR', '', errMsg, fileName);
      hasAnyError = true;
      return;
    }

    // 店舗コードからロケーションを特定
    var locationName = ProductsApi.findLocationName(accountName, detail.storeCode, storeCodeMap);
    if (!locationName) {
      var msg = '店舗コード「' + detail.storeCode + '」に対応するGBPロケーションが見つかりません';
      Logger.log('    ' + msg);
      logger.addRow(header, detail, 'ERROR', '', msg, fileName);
      hasAnyError = true;
      return;
    }

    // GBP に商品を登録
    try {
      var result = ProductsApi.createProduct(
        accountName,
        locationName,
        detail,
        imageRootFolderId
      );

      var postId = result.name || '';
      Logger.log('    登録成功: ' + postId);
      logger.addRow(header, detail, 'SUCCESS', postId, '', fileName);

    } catch (e) {
      Logger.log('    登録失敗: ' + e.message);
      logger.addRow(header, detail, 'ERROR', '', e.message, fileName);
      hasAnyError = true;
    }

    // 明細間のインターバル
    Utilities.sleep(300);
  });

  // ====== 4. ログの確定 & ファイル移動 ======
  logger.finalize();
  Logger.log('実行ログURL: ' + logger.getSpreadsheetUrl());

  if (hasAnyError) {
    // エラーありでも部分成功の場合は処理済みフォルダへ
    // （全件エラーの場合はエラーフォルダ）
    var details = parsed.details || [];
    var allError = details.length > 0 && details.every(function(item) {
      return item.errors.length > 0;
    });

    _moveFile(file, allError ? folders.error : folders.processed);
  } else {
    _moveFile(file, folders.processed);
  }
}

// ===== ユーティリティ =====

/**
 * アカウント名を "accounts/XXXXXX" 形式に正規化する
 */
function _normalizeAccountName(googleAccountId) {
  if (!googleAccountId) return '';
  var s = googleAccountId.trim();
  if (s.startsWith('accounts/')) return s;
  return 'accounts/' + s;
}

/**
 * 画像ルートフォルダIDを設定またはフォルダ名から解決する
 */
function _resolveImageRootFolderId() {
  // 入稿フォルダ内に「商品画像」フォルダが存在すればそのIDを使用
  try {
    var inboxFolder = CONFIG.INBOX_FOLDER_ID
      ? DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID)
      : DriveApp.getRootFolder();

    var it = inboxFolder.getFoldersByName(CONFIG.IMAGE_ROOT_FOLDER_NAME);
    if (it.hasNext()) return it.next().getId();
  } catch (e) {}

  return null; // 見つからない場合は null（imagePath をファイルIDとして解釈）
}

/**
 * ファイルを指定フォルダへ移動する
 */
function _moveFile(file, targetFolder) {
  try {
    file.moveTo(targetFolder);
  } catch (e) {
    Logger.log('ファイル移動に失敗しました (' + file.getName() + '): ' + e.message);
  }
}

/**
 * 親フォルダ内にサブフォルダを取得、なければ作成する
 */
function _getOrCreateSubFolder(parentFolder, name) {
  var it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : parentFolder.createFolder(name);
}

// ===== 入稿テンプレート作成ユーティリティ =====

/**
 * 入稿用テンプレートスプレッドシートを入稿フォルダに作成する
 * 初回セットアップ時に実行
 */
function createInputTemplate() {
  var folders = _getWorkFolders();
  var ss = SpreadsheetApp.create('入稿テンプレート_商品登録');
  var sheet = ss.getActiveSheet();
  sheet.setName('入稿シート');

  // ヘッダー部
  var headerData = [
    ['ビジネスグループID', '（ここに値を入力）', '', '', '', '', '', '', ''],
    ['ビジネスグループ名', '（ここに値を入力）', '', '', '', '', '', '', ''],
    ['GoogleアカウントID', 'accounts/XXXXXX', '', '', '', '', '', '', ''],
    ['PASS',              '（ここに値を入力）', '', '', '', '', '', '', ''],
    [],  // 空行
    // 明細ヘッダー
    ['ビジネス名', '店舗コード', '商品カテゴリ', '商品・サービス名', '商品の説明', '商品価格', 'ボタン追加', '商品のランディングページURL', '画像ファイルパス']
  ];

  sheet.getRange(1, 1, headerData.length, 9).setValues(headerData);

  // スタイル: ヘッダーキー列
  sheet.getRange(1, 1, 4, 1).setFontWeight('bold').setBackground('#e8f0fe');
  // スタイル: 明細ヘッダー行（6行目）
  var detailHeaderRange = sheet.getRange(6, 1, 1, 9);
  detailHeaderRange.setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(6);

  // サンプル明細（7行目）
  var sampleData = [
    ['渋谷店', 'SHOP-001', '食品・飲料', 'おすすめランチセット',
     '当店自慢のランチセットです。11:00〜14:00限定。', 1200,
     '詳細', 'https://example.com/lunch', '商品画像/渋谷店/lunch.jpg']
  ];
  sheet.getRange(7, 1, 1, 9).setValues(sampleData);
  sheet.getRange(7, 1, 1, 9).setBackground('#fff9c4');

  // 列幅の調整
  [150, 120, 130, 200, 250, 80, 100, 250, 250].forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  // ボタン追加のドロップダウン
  var buttonOptions = Object.keys(CONFIG.BUTTON_TYPE_MAP);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(buttonOptions, true).build();
  sheet.getRange(7, 7, 100, 1).setDataValidation(rule);

  // Drive フォルダへ移動
  var file = DriveApp.getFileById(ss.getId());
  folders.inbox.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  Logger.log('入稿テンプレートを作成しました: ' + ss.getUrl());
  Logger.log('入稿フォルダに保存されました。このファイルを複製して入稿に使用してください。');
}
