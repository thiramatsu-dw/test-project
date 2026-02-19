/**
 * GBPApi.js
 * Google Business Profile API との通信を担当するモジュール
 */

var GBPApi = (function() {

  /**
   * 認証ヘッダーを含むデフォルトリクエストオプションを返す
   */
  function _getRequestOptions(method, payload) {
    var options = {
      method: method || 'GET',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    if (payload) {
      options.payload = JSON.stringify(payload);
    }
    return options;
  }

  /**
   * APIレスポンスをパースして返す。エラーの場合は例外をスロー
   */
  function _parseResponse(response, context) {
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code >= 200 && code < 300) {
      return body ? JSON.parse(body) : {};
    }

    var errorMsg = 'API エラー [' + context + ']: HTTP ' + code + ' - ' + body;
    Logger.log(errorMsg);
    throw new Error(errorMsg);
  }

  /**
   * GBP に登録されているアカウント一覧を取得する
   * @returns {Object} accounts 配列を含むレスポンス
   */
  function listAccounts() {
    var url = CONFIG.GBP_ACCOUNT_API_BASE + '/accounts';
    var response = UrlFetchApp.fetch(url, _getRequestOptions('GET'));
    return _parseResponse(response, 'listAccounts');
  }

  /**
   * 指定アカウントのロケーション一覧を取得する
   * @param {string} accountName - "accounts/XXXXXX" 形式
   * @returns {Object} locations 配列を含むレスポンス
   */
  function listLocations(accountName) {
    var url = CONFIG.GBP_API_BASE + '/' + accountName + '/locations';
    var response = UrlFetchApp.fetch(url, _getRequestOptions('GET'));
    return _parseResponse(response, 'listLocations');
  }

  /**
   * ロケーションに既にアップロードされているメディア一覧を取得する
   * @param {string} accountName - "accounts/XXXXXX" 形式
   * @param {string} locationName - "locations/XXXXXX" 形式
   * @returns {Object} mediaItems 配列を含むレスポンス
   */
  function listMedia(accountName, locationName) {
    var url = CONFIG.GBP_API_BASE + '/' + accountName + '/' + locationName + '/media';
    var response = UrlFetchApp.fetch(url, _getRequestOptions('GET'));
    return _parseResponse(response, 'listMedia');
  }

  /**
   * Google Drive のファイルを指定ロケーションの GBP 画像としてアップロードする
   *
   * 仕組み:
   *  1. Drive ファイルを一時的に「リンクを知っている全員が閲覧可能」に設定
   *  2. 公開URLを sourceUrl として GBP API に送信
   *  3. アップロード後、Drive ファイルの共有設定を元に戻す
   *
   * @param {string} accountName  - "accounts/XXXXXX" 形式
   * @param {string} locationName - "locations/XXXXXX" 形式
   * @param {string} driveFileId  - Google Drive のファイルID
   * @param {string} category     - 写真カテゴリ（CONFIG.PHOTO_CATEGORIES のいずれか）
   * @returns {Object} 作成されたメディアアイテムの情報
   */
  function uploadMedia(accountName, locationName, driveFileId, category) {
    var file = DriveApp.getFileById(driveFileId);
    var originalAccess = file.getSharingAccess();
    var originalPermission = file.getSharingPermission();

    try {
      // 一時的に公開設定にして URL を取得
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      // GBP API が参照できる直接ダウンロードURL
      var sourceUrl = 'https://drive.google.com/uc?export=download&id=' + driveFileId;

      var mediaItem = {
        mediaFormat: 'PHOTO',
        locationAssociation: {
          category: category || CONFIG.PHOTO_CATEGORIES.ADDITIONAL
        },
        sourceUrl: sourceUrl
      };

      var url = CONFIG.GBP_API_BASE + '/' + accountName + '/' + locationName + '/media';
      var response = UrlFetchApp.fetch(url, _getRequestOptions('POST', mediaItem));
      var result = _parseResponse(response, 'uploadMedia');

      return result;

    } finally {
      // 必ず元の共有設定に戻す
      try {
        file.setSharing(originalAccess, originalPermission);
      } catch (e) {
        Logger.log('共有設定の復元に失敗しました (fileId: ' + driveFileId + '): ' + e.message);
      }
    }
  }

  // Public API
  return {
    listAccounts: listAccounts,
    listLocations: listLocations,
    listMedia: listMedia,
    uploadMedia: uploadMedia
  };

})();
