/**
 * ProductsApi.js
 * Google Business Profile への商品・サービス登録 API モジュール
 *
 * GBP LocalPosts API (topicType: PRODUCT) を使用して商品情報を投稿する
 */

var ProductsApi = (function() {

  /**
   * 認証済みリクエストオプションを生成する
   */
  function _getOptions(method, payload) {
    var opts = {
      method: method || 'GET',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    if (payload) opts.payload = JSON.stringify(payload);
    return opts;
  }

  /**
   * API レスポンスをパースし、エラー時は例外をスロー
   */
  function _parse(response, context) {
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code >= 200 && code < 300) {
      return body ? JSON.parse(body) : {};
    }

    // エラーレスポンスから詳細を取得
    var detail = body;
    try {
      var parsed = JSON.parse(body);
      detail = (parsed.error && parsed.error.message) ? parsed.error.message : body;
    } catch (e) {}

    throw new Error('[' + context + '] HTTP ' + code + ': ' + detail);
  }

  /**
   * 指定アカウント配下の全ロケーション（店舗）を取得し、
   * 店舗コードをキーとするマップを返す
   *
   * @param {string} accountName - "accounts/XXXXXX" 形式
   * @returns {Object} { storeCode: locationName, ... }
   */
  function buildStoreCodeMap(accountName) {
    var map = {};
    var pageToken = '';

    do {
      var url = CONFIG.GBP_API_BASE + '/' + accountName + '/locations?pageSize=100';
      if (pageToken) url += '&pageToken=' + encodeURIComponent(pageToken);

      var resp = UrlFetchApp.fetch(url, _getOptions('GET'));
      var data = _parse(resp, 'listLocations');

      (data.locations || []).forEach(function(loc) {
        if (loc.storeCode) {
          map[loc.storeCode] = loc.name; // "accounts/X/locations/Y"
        }
      });

      pageToken = data.nextPageToken || '';
    } while (pageToken);

    return map;
  }

  /**
   * GBP の LocationName を Google Business Account 配下の
   * "accounts/{id}/locations/{id}" 形式で取得する
   *
   * @param {string} accountName  - "accounts/XXXXXX"
   * @param {string} storeCode    - 入稿ファイルの店舗コード
   * @param {Object} storeCodeMap - buildStoreCodeMap() で取得したマップ
   * @returns {string|null} locationName or null if not found
   */
  function findLocationName(accountName, storeCode, storeCodeMap) {
    return storeCodeMap[storeCode] || null;
  }

  /**
   * 商品画像を Drive から取得し、GBP API が参照できる URL を返す
   * Drive ファイルを一時公開して URL を発行し、呼び出し元が戻した後に復元すること
   *
   * @param {string} imagePath - Google Drive ファイルID または "フォルダ名/ファイル名" 形式
   * @param {string} imageRootFolderId - 画像ルートフォルダID（省略可）
   * @returns {{ file: File, url: string } | null}
   */
  function prepareImageUrl(imagePath, imageRootFolderId) {
    if (!imagePath) return null;

    var file = null;

    // パスがDriveのURL形式 (https://drive.google.com/...) の場合
    var urlMatch = imagePath.match(/(?:id=|\/d\/)([a-zA-Z0-9_-]{25,})/);
    if (urlMatch) {
      try { file = DriveApp.getFileById(urlMatch[1]); } catch (e) {}
    }

    // まずファイルIDとして直接解決を試みる
    if (!file) {
      try { file = DriveApp.getFileById(imagePath); } catch (e) {}
    }

    // "フォルダ名/ファイル名" パス形式
    if (!file && imagePath.indexOf('/') !== -1) {
      file = _resolveByPath(imagePath, imageRootFolderId);
    }

    if (!file) {
      throw new Error('画像ファイルが見つかりません: ' + imagePath);
    }

    // 一時公開してURLを返す
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/uc?export=download&id=' + file.getId();
    return { file: file, url: url };
  }

  /**
   * "フォルダ名/ファイル名" 形式のパスをファイルオブジェクトに解決する
   */
  function _resolveByPath(imagePath, rootFolderId) {
    var parts = imagePath.split('/').map(function(p) { return p.trim(); }).filter(Boolean);
    var folder = rootFolderId
      ? DriveApp.getFolderById(rootFolderId)
      : DriveApp.getRootFolder();

    for (var i = 0; i < parts.length - 1; i++) {
      var it = folder.getFoldersByName(parts[i]);
      if (!it.hasNext()) return null;
      folder = it.next();
    }

    var fileName = parts[parts.length - 1];
    var files = folder.getFilesByName(fileName);
    return files.hasNext() ? files.next() : null;
  }

  /**
   * GBP に商品ローカルポストを作成する
   *
   * @param {string} accountName  - "accounts/XXXXXX"
   * @param {string} locationName - "accounts/X/locations/Y"（フル形式）
   * @param {Object} productData  - 明細フィールドオブジェクト
   * @param {string|null} imageRootFolderId - 画像ルートフォルダID
   * @returns {Object} 作成されたローカルポストオブジェクト
   */
  function createProduct(accountName, locationName, productData, imageRootFolderId) {
    var imageRef = null;

    // 画像の準備
    try {
      if (productData.imagePath) {
        imageRef = prepareImageUrl(productData.imagePath, imageRootFolderId);
      }
    } catch (e) {
      throw new Error('画像の準備に失敗しました: ' + e.message);
    }

    try {
      // LocalPost ペイロードを構築
      var post = _buildLocalPost(productData, imageRef ? imageRef.url : null);

      var url = CONFIG.GBP_API_BASE + '/' + locationName + '/localPosts';
      var resp = UrlFetchApp.fetch(url, _getOptions('POST', post));
      var result = _parse(resp, 'createProduct');

      return result;

    } finally {
      // 画像の一時公開を必ず解除
      if (imageRef && imageRef.file) {
        try {
          imageRef.file.setSharing(
            DriveApp.Access.PRIVATE,
            DriveApp.Permission.NONE
          );
        } catch (e) {
          Logger.log('画像共有設定の復元に失敗: ' + e.message);
        }
      }
    }
  }

  /**
   * GBP LocalPost ペイロードを構築する
   */
  function _buildLocalPost(d, imageUrl) {
    var post = {
      languageCode: 'ja',
      topicType: 'PRODUCT',
      summary: d.description || '',
      product: {
        name: d.productName || '',
        category: d.productCategory || '',
        description: d.description || ''
      }
    };

    // 価格
    if (d.price && d.price > 0) {
      post.product.price = {
        currencyCode: 'JPY',
        units: String(Math.floor(d.price))
      };
    }

    // CTA ボタン
    var actionType = d.buttonTypeApi || '';
    if (actionType && d.landingPageUrl) {
      post.callToAction = {
        actionType: actionType,
        url: d.landingPageUrl
      };
    } else if (d.landingPageUrl) {
      // ボタン種別未指定でURLあり → LEARN_MORE
      post.callToAction = {
        actionType: 'LEARN_MORE',
        url: d.landingPageUrl
      };
    }

    // 画像
    if (imageUrl) {
      post.media = [{
        mediaFormat: 'PHOTO',
        sourceUrl: imageUrl
      }];
    }

    return post;
  }

  // Public API
  return {
    buildStoreCodeMap: buildStoreCodeMap,
    findLocationName: findLocationName,
    createProduct: createProduct
  };

})();
