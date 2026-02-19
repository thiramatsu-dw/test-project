/**
 * ProductsApi.js
 * Google Business Profile への商品・サービス登録 API モジュール
 *
 * GBP LocalPosts API (topicType: PRODUCT) を使用して商品情報を投稿する。
 */

var ProductsApi = (function() {

  // ===== 内部ヘルパー =====

  /**
   * 認証済みリクエストオプションを生成する
   * @param {string} method  - HTTP メソッド
   * @param {Object} payload - リクエストボディ（オブジェクト）
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
   * @param {HTTPResponse} response
   * @param {string}       context  - エラーメッセージ用ラベル
   */
  function _parse(response, context) {
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code >= 200 && code < 300) {
      return body ? JSON.parse(body) : {};
    }

    var detail = body;
    try {
      var parsed = JSON.parse(body);
      detail = (parsed.error && parsed.error.message) ? parsed.error.message : body;
    } catch (e) {}

    throw new Error('[' + context + '] HTTP ' + code + ': ' + detail);
  }

  // ===== ロケーション解決 =====

  /**
   * 指定アカウント配下の全ロケーションを取得し、
   * 「店舗コード → ロケーション名（フルパス）」マップを返す。
   *
   * ページネーションに対応しており、100件超のロケーションも網羅する。
   *
   * @param {string} accountName - "accounts/XXXXXX" 形式
   * @returns {Object} { "店舗コード": "accounts/X/locations/Y", ... }
   * @throws ロケーション一覧取得に失敗した場合
   */
  function buildStoreCodeMap(accountName) {
    var map       = {};
    var pageToken = '';
    var pageNum   = 0;

    do {
      pageNum++;
      var url = CONFIG.GBP_API_BASE + '/' + accountName +
                '/locations?pageSize=100' +
                (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

      var resp = UrlFetchApp.fetch(url, _getOptions('GET'));
      var data = _parse(resp, 'listLocations(page=' + pageNum + ')');

      (data.locations || []).forEach(function(loc) {
        // storeCode が設定されているロケーションのみマップに追加
        if (loc.storeCode) {
          map[loc.storeCode] = loc.name; // "accounts/X/locations/Y"
        } else {
          // storeCode 未設定のロケーションはロケーション名で警告
          Logger.log('警告: storeCode が未設定のロケーション → ' + loc.name +
                     ' (locationName: ' + (loc.locationName || '-') + ')');
        }
      });

      pageToken = data.nextPageToken || '';

    } while (pageToken);

    Logger.log('ロケーションマップ構築完了: ' + Object.keys(map).length + ' 件（' +
               pageNum + ' ページ）');
    return map;
  }

  /**
   * 店舗コードに対応するロケーション名を返す。
   * 見つからない場合は null を返す。
   *
   * @param {string} accountName  - "accounts/XXXXXX" （将来の拡張のため引数として保持）
   * @param {string} storeCode    - 入稿ファイルの店舗コード
   * @param {Object} storeCodeMap - buildStoreCodeMap() で生成したマップ
   * @returns {string|null}
   */
  function findLocationName(accountName, storeCode, storeCodeMap) {
    return storeCodeMap[storeCode] || null;
  }

  // ===== 画像 URL 準備 =====

  /**
   * Google Drive の画像ファイルを特定し、GBP API が参照できる一時公開 URL を発行する。
   *
   * ★ 共有設定の保存・復元:
   *   公開前の Access / Permission を保存し、戻り値に含める。
   *   呼び出し元は必ず restoreImageSharing() を呼ぶこと。
   *
   * imagePath の解釈順序:
   *   1. Google Drive URL 形式 (https://drive.google.com/...)
   *   2. Drive ファイル ID（直接）
   *   3. "フォルダ名/ファイル名" 形式のパス（imageRootFolderId 配下）
   *
   * @param {string}      imagePath        - ファイル指定文字列
   * @param {string|null} imageRootFolderId - 画像ルートフォルダID（パス形式で使用）
   * @returns {{ file: File, url: string, origAccess, origPermission } | null}
   * @throws 画像が見つからない場合
   */
  function prepareImageUrl(imagePath, imageRootFolderId) {
    if (!imagePath || imagePath.trim() === '') return null;

    var file = _resolveFile(imagePath.trim(), imageRootFolderId);
    if (!file) {
      throw new Error('画像ファイルが見つかりません: ' + imagePath);
    }

    // ★ 変更前の共有設定を保存
    var origAccess     = file.getSharingAccess();
    var origPermission = file.getSharingPermission();

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/uc?export=download&id=' + file.getId();

    return {
      file:           file,
      url:            url,
      origAccess:     origAccess,
      origPermission: origPermission
    };
  }

  /**
   * prepareImageUrl() で変更した共有設定を元の状態に戻す。
   * @param {{ file, origAccess, origPermission }} imageRef
   */
  function restoreImageSharing(imageRef) {
    if (!imageRef || !imageRef.file) return;
    try {
      imageRef.file.setSharing(imageRef.origAccess, imageRef.origPermission);
    } catch (e) {
      Logger.log('画像共有設定の復元に失敗しました (ID: ' + imageRef.file.getId() + '): ' + e.message);
    }
  }

  /**
   * imagePath 文字列からファイルオブジェクトを解決する内部関数
   */
  function _resolveFile(imagePath, rootFolderId) {
    var file = null;

    // 1. Google Drive URL 形式
    var urlMatch = imagePath.match(/(?:id=|\/d\/)([a-zA-Z0-9_-]{25,})/);
    if (urlMatch) {
      try { file = DriveApp.getFileById(urlMatch[1]); } catch (e) {}
    }

    // 2. Drive ファイル ID（25文字以上の英数字ハイフン）
    if (!file && /^[a-zA-Z0-9_-]{25,}$/.test(imagePath)) {
      try { file = DriveApp.getFileById(imagePath); } catch (e) {}
    }

    // 3. "フォルダ名/ファイル名" 形式
    if (!file) {
      file = _resolveByPath(imagePath, rootFolderId);
    }

    return file;
  }

  /**
   * "フォルダA/フォルダB/ファイル名.jpg" 形式のパスを解決する
   */
  function _resolveByPath(imagePath, rootFolderId) {
    var parts  = imagePath.split('/').map(function(p) { return p.trim(); }).filter(Boolean);
    if (parts.length === 0) return null;

    var folder = rootFolderId
      ? DriveApp.getFolderById(rootFolderId)
      : DriveApp.getRootFolder();

    // フォルダ階層をたどる
    for (var i = 0; i < parts.length - 1; i++) {
      var it = folder.getFoldersByName(parts[i]);
      if (!it.hasNext()) {
        Logger.log('フォルダが見つかりません: ' + parts.slice(0, i + 1).join('/'));
        return null;
      }
      folder = it.next();
    }

    // ファイルを取得
    var fileName = parts[parts.length - 1];
    var files    = folder.getFilesByName(fileName);
    if (!files.hasNext()) {
      Logger.log('ファイルが見つかりません: ' + imagePath);
      return null;
    }
    return files.next();
  }

  // ===== 商品登録 =====

  /**
   * GBP に商品ローカルポストを作成する。
   *
   * 処理の流れ:
   *   1. 画像が指定されている場合、Drive から一時公開 URL を取得
   *   2. LocalPost ペイロードを構築して POST
   *   3. finally ブロックで画像の共有設定を必ず元に戻す
   *
   * @param {string}      accountName       - "accounts/XXXXXX"
   * @param {string}      locationName      - "accounts/X/locations/Y"（フルパス形式）
   * @param {Object}      productData       - 明細フィールドオブジェクト
   * @param {string|null} imageRootFolderId - 画像ルートフォルダID
   * @returns {Object} 作成されたローカルポストオブジェクト
   * @throws API 呼び出し失敗、または画像準備失敗の場合
   */
  function createProduct(accountName, locationName, productData, imageRootFolderId) {
    var imageRef = null;

    // 画像の準備（エラー時はここで例外をスロー）
    if (productData.imagePath) {
      imageRef = prepareImageUrl(productData.imagePath, imageRootFolderId);
    }

    try {
      var post = _buildLocalPost(productData, imageRef ? imageRef.url : null);
      var url  = CONFIG.GBP_API_BASE + '/' + locationName + '/localPosts';
      var resp = UrlFetchApp.fetch(url, _getOptions('POST', post));
      return _parse(resp, 'createProduct');

    } finally {
      // ★ 成功・失敗に関わらず元の共有設定に戻す
      restoreImageSharing(imageRef);
    }
  }

  /**
   * GBP LocalPost リクエストボディを構築する
   *
   * 対応フィールド:
   *   - topicType: PRODUCT（固定）
   *   - product.name, product.category, product.description
   *   - product.price（0円より大きい場合のみ設定）
   *   - callToAction（ボタン種別と URL が両方ある場合のみ設定）
   *   - media（画像 URL がある場合のみ設定）
   *
   * @param {Object}      d        - 明細フィールドオブジェクト
   * @param {string|null} imageUrl - 公開済み画像 URL
   */
  function _buildLocalPost(d, imageUrl) {
    var post = {
      languageCode: 'ja',
      topicType:    'PRODUCT',
      summary:      d.description || '',
      product: {
        name:        d.productName    || '',
        category:    d.productCategory || '',
        description: d.description    || ''
      }
    };

    // 価格（1円以上の場合のみ）
    var priceVal = parseFloat(d.price) || 0;
    if (priceVal > 0) {
      post.product.price = {
        currencyCode: 'JPY',
        units:        String(Math.floor(priceVal))
        // nanos は整数円のため省略
      };
    }

    // CTA ボタン
    var actionType = d.buttonTypeApi || '';
    if (actionType && actionType !== '' && d.landingPageUrl) {
      post.callToAction = { actionType: actionType, url: d.landingPageUrl };
    } else if (!actionType && d.landingPageUrl) {
      // ボタン種別未指定でURLあり → LEARN_MORE にフォールバック
      post.callToAction = { actionType: 'LEARN_MORE', url: d.landingPageUrl };
    }

    // 画像
    if (imageUrl) {
      post.media = [{ mediaFormat: 'PHOTO', sourceUrl: imageUrl }];
    }

    return post;
  }

  // ===== Public API =====
  return {
    buildStoreCodeMap: buildStoreCodeMap,
    findLocationName:  findLocationName,
    createProduct:     createProduct
  };

})();
