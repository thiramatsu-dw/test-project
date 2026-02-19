/**
 * ImageUploader.js
 * 全店舗の画像アップロードを統括するメインモジュール
 */

/**
 * 全店舗の画像を一括アップロードするメイン関数
 * トリガーからも手動からも実行可能
 */
function uploadAllStoreImages() {
  Logger.log('=== 画像アップロード開始: ' + new Date().toLocaleString('ja-JP') + ' ===');

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var configSheet = ss.getSheetByName(CONFIG.SHEETS.STORE_CONFIG);

  if (!configSheet) {
    Logger.log('エラー: 「' + CONFIG.SHEETS.STORE_CONFIG + '」シートが見つかりません。setupSpreadsheet() を先に実行してください。');
    return;
  }

  var stores = _getActiveStores(configSheet);
  if (stores.length === 0) {
    Logger.log('有効な店舗データがありません。');
    return;
  }

  Logger.log('処理対象店舗数: ' + stores.length);

  var totalUploaded = 0;
  var totalFailed = 0;

  stores.forEach(function(store, index) {
    Logger.log('\n--- 店舗 ' + (index + 1) + '/' + stores.length + ': ' + store.storeName + ' ---');

    var result = _processStore(ss, configSheet, store);
    totalUploaded += result.uploaded;
    totalFailed += result.failed;

    // 店舗間で少し待機（API レートリミット対策）
    if (index < stores.length - 1) {
      Utilities.sleep(1000);
    }
  });

  Logger.log('\n=== アップロード完了 ===');
  Logger.log('成功: ' + totalUploaded + ' 件');
  Logger.log('失敗: ' + totalFailed + ' 件');
}

/**
 * 店舗設定シートから有効な店舗データを取得する
 * @param {Sheet} configSheet - 店舗設定シート
 * @returns {Array} 店舗データの配列
 */
function _getActiveStores(configSheet) {
  var lastRow = configSheet.getLastRow();
  if (lastRow < 2) return [];

  var data = configSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var stores = [];

  data.forEach(function(row, i) {
    var storeName   = String(row[CONFIG.STORE_COLS.STORE_NAME]).trim();
    var accountId   = String(row[CONFIG.STORE_COLS.ACCOUNT_ID]).trim();
    var locationId  = String(row[CONFIG.STORE_COLS.LOCATION_ID]).trim();
    var folderId    = String(row[CONFIG.STORE_COLS.FOLDER_ID]).trim();
    var category    = String(row[CONFIG.STORE_COLS.CATEGORY]).trim() || CONFIG.PHOTO_CATEGORIES.ADDITIONAL;
    var status      = String(row[CONFIG.STORE_COLS.STATUS]).trim();

    // 必須フィールドのバリデーション
    if (!storeName || storeName === '' || storeName === 'undefined') return;
    if (status !== '有効') return;
    if (!accountId || !locationId || !folderId) {
      Logger.log('警告: ' + storeName + ' の設定が不完全です（行 ' + (i + 2) + '）。スキップします。');
      return;
    }

    stores.push({
      rowIndex: i + 2,  // シートの実際の行番号
      storeName: storeName,
      accountId: accountId,
      locationId: locationId,
      folderId: folderId,
      category: category
    });
  });

  return stores;
}

/**
 * 1店舗分の画像アップロード処理を行う
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Sheet} configSheet - 店舗設定シート
 * @param {Object} store - 店舗データ
 * @returns {Object} {uploaded: number, failed: number}
 */
function _processStore(ss, configSheet, store) {
  var uploaded = 0;
  var failed = 0;

  // Drive フォルダから画像ファイルを取得
  var imageFiles = _getImagesFromFolder(store.folderId);

  if (imageFiles.length === 0) {
    Logger.log(store.storeName + ': アップロード対象の画像がありません。');
    return { uploaded: 0, failed: 0 };
  }

  Logger.log(store.storeName + ': ' + imageFiles.length + ' 件の画像を処理します。');

  imageFiles.forEach(function(file) {
    var fileName = file.getName();
    var fileId = file.getId();

    try {
      // GBP API で画像をアップロード
      GBPApi.uploadMedia(
        store.accountId,
        store.locationId,
        fileId,
        store.category
      );

      Logger.log('  ✓ アップロード成功: ' + fileName);

      // アップロード済みフォルダへ移動
      _moveToUploadedFolder(file, store.folderId);

      // 履歴シートに記録
      _logHistory(ss, store.storeName, fileName, fileId, '成功', 'アップロード完了');

      uploaded++;

    } catch (e) {
      var errMsg = e.message || String(e);
      Logger.log('  ✗ アップロード失敗: ' + fileName + ' - ' + errMsg);

      // 履歴シートに記録
      _logHistory(ss, store.storeName, fileName, fileId, '失敗', errMsg);

      failed++;
    }

    // ファイル間で少し待機
    Utilities.sleep(500);
  });

  // 店舗設定シートの最終アップロード日時を更新
  if (uploaded > 0 || failed > 0) {
    configSheet.getRange(store.rowIndex, CONFIG.STORE_COLS.LAST_UPLOAD + 1)
      .setValue(new Date());
  }

  return { uploaded: uploaded, failed: failed };
}

/**
 * 指定フォルダから対象画像ファイルを取得する
 * アップロード済みサブフォルダ内のファイルは除外する
 * @param {string} folderId - Drive フォルダ ID
 * @returns {Array} File オブジェクトの配列
 */
function _getImagesFromFolder(folderId) {
  var imageFiles = [];

  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();

    while (files.hasNext()) {
      var file = files.next();
      var mimeType = file.getMimeType();
      var name = file.getName().toLowerCase();

      // 画像ファイルのみ対象
      var isImage = (
        mimeType === 'image/jpeg' ||
        mimeType === 'image/png' ||
        mimeType === 'image/gif'
      );

      // 拡張子チェック（MIMEタイプが不明な場合のフォールバック）
      var hasImageExt = CONFIG.ALLOWED_EXTENSIONS.some(function(ext) {
        return name.endsWith('.' + ext);
      });

      if (isImage || hasImageExt) {
        imageFiles.push(file);
      }
    }
  } catch (e) {
    Logger.log('フォルダの読み込みに失敗しました (folderId: ' + folderId + '): ' + e.message);
  }

  return imageFiles;
}

/**
 * アップロード済みファイルを「uploaded」サブフォルダへ移動する
 * @param {File} file - 移動対象のファイル
 * @param {string} parentFolderId - 親フォルダ ID
 */
function _moveToUploadedFolder(file, parentFolderId) {
  try {
    var parentFolder = DriveApp.getFolderById(parentFolderId);

    // 「uploaded」フォルダを取得または作成
    var uploadedFolder;
    var subFolders = parentFolder.getFoldersByName(CONFIG.UPLOADED_FOLDER_NAME);

    if (subFolders.hasNext()) {
      uploadedFolder = subFolders.next();
    } else {
      uploadedFolder = parentFolder.createFolder(CONFIG.UPLOADED_FOLDER_NAME);
    }

    // ファイルを移動
    file.moveTo(uploadedFolder);

  } catch (e) {
    Logger.log('ファイルの移動に失敗しました (' + file.getName() + '): ' + e.message);
    // 移動失敗はアップロード成否に影響しないので例外を再スローしない
  }
}

/**
 * アップロード履歴をシートに記録する
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} storeName - 店舗名
 * @param {string} fileName - ファイル名
 * @param {string} fileId - Drive ファイル ID
 * @param {string} result - 「成功」or「失敗」
 * @param {string} message - 詳細メッセージ
 */
function _logHistory(ss, storeName, fileName, fileId, result, message) {
  try {
    var historySheet = ss.getSheetByName(CONFIG.SHEETS.UPLOAD_HISTORY);
    if (!historySheet) return;

    var row = [
      new Date(),
      storeName,
      fileName,
      fileId,
      result,
      message
    ];

    historySheet.appendRow(row);

    // 結果に応じて行の背景色を変更
    var lastRow = historySheet.getLastRow();
    var resultCell = historySheet.getRange(lastRow, CONFIG.HISTORY_COLS.RESULT + 1);
    if (result === '成功') {
      resultCell.setBackground('#d9ead3');
    } else {
      resultCell.setBackground('#fce5cd');
    }

  } catch (e) {
    Logger.log('履歴の記録に失敗しました: ' + e.message);
  }
}

/**
 * 特定の店舗だけを手動でアップロードするためのユーティリティ関数
 * @param {string} storeName - 処理する店舗名
 */
function uploadSingleStore(storeName) {
  if (!storeName) {
    Logger.log('使用方法: uploadSingleStore("店舗名")');
    return;
  }

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var configSheet = ss.getSheetByName(CONFIG.SHEETS.STORE_CONFIG);

  if (!configSheet) {
    Logger.log('エラー: 店舗設定シートが見つかりません。');
    return;
  }

  var stores = _getActiveStores(configSheet);
  var target = stores.filter(function(s) { return s.storeName === storeName; });

  if (target.length === 0) {
    Logger.log('店舗「' + storeName + '」が見つかりません（有効な店舗として設定されているか確認してください）。');
    return;
  }

  var result = _processStore(ss, configSheet, target[0]);
  Logger.log('完了: 成功 ' + result.uploaded + ' 件、失敗 ' + result.failed + ' 件');
}
