/**
 * MainProcessor.js
 * 入稿フォルダを監視して商品情報を一括登録するメインオーケストレーター
 *
 * 【処理フロー】
 *  1. 入稿フォルダ内の新規ファイルを検索
 *  2. InputProcessor で入稿ファイルをパース
 *  3. 各明細行について:
 *     a. 店舗コードから GBP ロケーションを特定
 *     b. ProductsApi で商品ローカルポストを作成（画像は API 内で一時公開・復元）
 *  4. ExecutionLogger でログシートを出力
 *  5. 入稿ファイルを処理済み / エラーフォルダへ移動
 *
 * ★ フォルダ振り分けルール:
 *    - 1件でも成功 → 処理済みフォルダ
 *    - 全件失敗（成功0件） → エラーフォルダ
 */

// ===== エントリーポイント =====

/**
 * 入稿フォルダ内の全ファイルを処理するメイン関数
 * トリガーからも手動実行からも呼び出し可能
 */
function processInboxFiles() {
  Logger.log('=== 入稿処理開始: ' + new Date().toLocaleString('ja-JP') + ' ===');

  var folders = _getWorkFolders();
  var files   = _listInboxFiles(folders.inbox);

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
      Logger.log('予期せぬエラーが発生しました: ' + e.message);
      _moveFile(file, folders.error);
    }

    // ファイル間でインターバル（API レートリミット対策）
    if (idx < files.length - 1) Utilities.sleep(2000);
  });

  Logger.log('\n=== 入稿処理完了 ===');
}

/**
 * 特定のファイル ID を指定して処理する（再処理・テスト用）
 * @param {string} fileId - 処理する Drive ファイルの ID
 */
function processFileById(fileId) {
  if (!fileId) {
    Logger.log('使用方法: processFileById("DriveファイルID")');
    return;
  }
  var file = DriveApp.getFileById(fileId);
  Logger.log('ファイルを処理します: ' + file.getName());
  _processOneFile(file, _getWorkFolders());
}

// ===== 内部処理 =====

/**
 * 1ファイルの入稿処理を行う
 * @returns {{ successCount: number, errorCount: number }}
 */
function _processOneFile(file, folders) {
  var fileName = file.getName();

  // ====== 1. ファイルのパース ======
  var parsed = InputProcessor.parse(file);
  var logger = ExecutionLogger.create(folders.results, fileName);

  // ヘッダーレベルのエラー処理
  if (parsed.errors.length > 0) {
    parsed.errors.forEach(function(err) {
      Logger.log('  入稿エラー: ' + err);
      logger.addHeaderError(err, fileName);
    });

    if (!parsed.header) {
      // ヘッダーが読めない致命的エラー → ログ確定後にエラーフォルダへ
      logger.finalize();
      _moveFile(file, folders.error);
      return { successCount: 0, errorCount: 1 };
    }
  }

  var header = parsed.header;
  Logger.log('  グループ: ' + header.businessGroupName +
             ' | アカウントID: ' + header.googleAccountId);

  var accountName = _normalizeAccountName(header.googleAccountId);

  // ====== 2. 店舗コードマップの取得 ======
  var storeCodeMap = {};
  try {
    storeCodeMap = ProductsApi.buildStoreCodeMap(accountName);
  } catch (e) {
    var mapErrMsg = 'GBP ロケーション一覧の取得に失敗しました: ' + e.message;
    Logger.log('  ' + mapErrMsg);
    logger.addHeaderError(mapErrMsg, fileName);
    logger.finalize();
    _moveFile(file, folders.error);
    return { successCount: 0, errorCount: 1 };
  }

  // 画像ルートフォルダ ID の解決
  var imageRootFolderId = _resolveImageRootFolderId(folders.inbox);

  // ====== 3. 明細行の処理 ======
  var successCount = 0;
  var errorCount   = 0;

  parsed.details.forEach(function(item) {
    var detail = item.data;
    Logger.log('  行' + item.rowNum + ': ' + detail.businessName + ' / ' + detail.productName);

    // バリデーションエラー
    if (item.errors.length > 0) {
      var valErrMsg = item.errors.join(' / ');
      Logger.log('    バリデーションエラー: ' + valErrMsg);
      logger.addRow(header, detail, 'ERROR', '', valErrMsg, fileName);
      errorCount++;
      return;
    }

    // 店舗コードからロケーションを特定
    var locationName = ProductsApi.findLocationName(accountName, detail.storeCode, storeCodeMap);
    if (!locationName) {
      var notFoundMsg = '店舗コード「' + detail.storeCode + '」に対応する GBP ロケーションが見つかりません' +
                        '（GBP 管理画面で storeCode が設定されているか確認してください）';
      Logger.log('    ' + notFoundMsg);
      logger.addRow(header, detail, 'ERROR', '', notFoundMsg, fileName);
      errorCount++;
      return;
    }

    // GBP に商品を登録
    try {
      var result  = ProductsApi.createProduct(accountName, locationName, detail, imageRootFolderId);
      var postId  = result.name || '';
      Logger.log('    登録成功: ' + postId);
      logger.addRow(header, detail, 'SUCCESS', postId, '', fileName);
      successCount++;
    } catch (e) {
      Logger.log('    登録失敗: ' + e.message);
      logger.addRow(header, detail, 'ERROR', '', e.message, fileName);
      errorCount++;
    }

    // 明細間インターバル（API レートリミット対策）
    Utilities.sleep(300);
  });

  // ====== 4. ログ確定・ファイル振り分け ======
  logger.finalize();
  Logger.log('  実行ログ URL: ' + logger.getSpreadsheetUrl());
  Logger.log('  結果: 成功 ' + successCount + ' 件, 失敗 ' + errorCount + ' 件');

  // ★ 1件でも成功があれば処理済み、全件失敗ならエラーフォルダへ
  _moveFile(file, successCount > 0 ? folders.processed : folders.error);

  return { successCount: successCount, errorCount: errorCount };
}

// ===== ユーティリティ =====

/**
 * 作業用フォルダ（入稿・処理済み・エラー・処理結果）を取得または作成する
 */
function _getWorkFolders() {
  var inbox;

  if (CONFIG.INBOX_FOLDER_ID) {
    inbox = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  } else {
    var rootIt = DriveApp.getRootFolder().getFoldersByName(CONFIG.INBOX_FOLDER_NAME);
    inbox = rootIt.hasNext()
      ? rootIt.next()
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
 * 入稿フォルダ内の処理対象ファイル（Sheets / Excel / CSV）を返す
 *
 * DriveApp.Folder.getFiles() は直接の子ファイルのみ返すため、
 * サブフォルダ（処理済み・エラー等）内のファイルは自動的に除外される。
 */
function _listInboxFiles(inboxFolder) {
  var files = [];
  var it    = inboxFolder.getFiles();

  while (it.hasNext()) {
    var f    = it.next();
    var mime = f.getMimeType();

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
 * アカウント名を "accounts/XXXXXX" 形式に正規化する
 * @param {string} googleAccountId - "accounts/123" または "123" 形式
 */
function _normalizeAccountName(googleAccountId) {
  if (!googleAccountId) return '';
  var s = googleAccountId.trim();
  return s.startsWith('accounts/') ? s : 'accounts/' + s;
}

/**
 * 入稿フォルダ内の「商品画像」フォルダ ID を返す
 * 見つからない場合は null（imagePath をファイル ID / Drive URL として解釈）
 * @param {Folder} inboxFolder
 */
function _resolveImageRootFolderId(inboxFolder) {
  try {
    var it = inboxFolder.getFoldersByName(CONFIG.IMAGE_ROOT_FOLDER_NAME);
    if (it.hasNext()) return it.next().getId();
  } catch (e) {}
  return null;
}

/** ファイルを指定フォルダへ移動する */
function _moveFile(file, targetFolder) {
  try {
    file.moveTo(targetFolder);
  } catch (e) {
    Logger.log('ファイルの移動に失敗しました (' + file.getName() + '): ' + e.message);
  }
}

/** 親フォルダ内のサブフォルダを取得、なければ作成する */
function _getOrCreateSubFolder(parentFolder, name) {
  var it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : parentFolder.createFolder(name);
}

// ===== テンプレート作成ユーティリティ =====

/**
 * 入稿用テンプレートスプレッドシートを入稿フォルダに作成する
 * 初回セットアップ時に実行してください
 */
function createInputTemplate() {
  var folders = _getWorkFolders();
  var ss      = SpreadsheetApp.create('入稿テンプレート_商品登録');
  var sheet   = ss.getActiveSheet();
  sheet.setName('入稿シート');

  // ===== ヘッダー部（行1〜5）=====
  var headerSection = [
    ['ビジネスグループID', '',             '', '', '', '', '', '', ''],
    ['ビジネスグループ名', '',             '', '', '', '', '', '', ''],
    ['GoogleアカウントID', 'accounts/',    '', '', '', '', '', '', ''],
    ['PASS',              '',             '', '', '', '', '', '', ''],
    []   // 空行
  ];
  sheet.getRange(1, 1, headerSection.length, 9).setValues(headerSection);

  // ヘッダーキー列のスタイル
  sheet.getRange(1, 1, 4, 1)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  // ヘッダー値列のガイドスタイル
  sheet.getRange(1, 2, 4, 1)
    .setBackground('#f8f9fa')
    .setFontColor('#5f6368')
    .setFontStyle('italic');

  // ===== 明細ヘッダー行（行6）=====
  var colNames = Object.keys(CONFIG.INPUT_DETAIL_COLS);
  sheet.getRange(6, 1, 1, colNames.length).setValues([colNames]);
  sheet.getRange(6, 1, 1, colNames.length)
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(6);

  // ===== サンプルデータ（行7）=====
  var sample = [
    '渋谷店',         // ビジネス名
    'SHOP-001',       // 店舗コード（GBP の storeCode と一致させること）
    '食品・飲料',      // 商品カテゴリ
    'おすすめランチセット', // 商品・サービス名
    '当店自慢のランチセット。11:00〜14:00限定。旬の食材を使った日替わりメニューです。', // 商品の説明
    1200,             // 商品価格
    '詳細',           // ボタン追加
    'https://example.com/lunch', // ランディングページURL
    '商品画像/渋谷店/lunch.jpg'  // 画像ファイルパス（Driveパス または ファイルID）
  ];
  sheet.getRange(7, 1, 1, sample.length).setValues([sample]);
  sheet.getRange(7, 1, 1, sample.length).setBackground('#fff9c4');

  // ===== 列幅 =====
  var colWidths = [120, 110, 120, 200, 300, 80, 90, 250, 250];
  colWidths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.setRowHeight(7, 21);

  // ===== ボタン追加のドロップダウン（8行目以降に適用）=====
  var btnOptions = Object.keys(CONFIG.BUTTON_TYPE_MAP);
  var btnRule    = SpreadsheetApp.newDataValidation()
    .requireValueInList(btnOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(7, 7, 1000, 1).setDataValidation(btnRule);

  // ===== ガイドメモ =====
  var noteRange = sheet.getRange(1, 4);
  noteRange.setNote(
    '【使い方】\n' +
    '1. 1〜4行目のB列に各値を入力してください\n' +
    '2. 7行目以降に商品情報を入力してください（サンプルを参考に）\n' +
    '3. 画像ファイルパスは Driveパス（例: 商品画像/店舗名/image.jpg）\n' +
    '   または Drive ファイル ID を入力してください\n' +
    '4. 入稿フォルダにこのファイルを保存して processInboxFiles() を実行してください'
  );

  // ===== Drive 入稿フォルダへ移動 =====
  var file = DriveApp.getFileById(ss.getId());
  folders.inbox.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}

  Logger.log('入稿テンプレートを作成しました: ' + ss.getUrl());
  Logger.log('入稿フォルダ: ' + folders.inbox.getUrl());
}
