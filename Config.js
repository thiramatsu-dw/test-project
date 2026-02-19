/**
 * Config.js
 * システム全体の設定定数
 */

var CONFIG = {
  // スプレッドシートID（このスクリプトと同じスプレッドシートを使用する場合は SpreadsheetApp.getActiveSpreadsheet() を使用）
  SPREADSHEET_ID: '1nZFiS6wFj3Jk9W7QdKzzMfLkr2PY5hVjre5ZvPoglfc',

  // シート名
  SHEETS: {
    STORE_CONFIG: '店舗設定',    // 店舗・ロケーション設定シート
    UPLOAD_HISTORY: 'アップロード履歴' // アップロード履歴シート
  },

  // Google Business Profile API ベースURL
  GBP_API_BASE: 'https://mybusiness.googleapis.com/v4',
  GBP_ACCOUNT_API_BASE: 'https://mybusinessaccountmanagement.googleapis.com/v1',

  // 画像カテゴリの選択肢（GBP APIで使用）
  PHOTO_CATEGORIES: {
    COVER: 'COVER',               // カバー写真
    PROFILE: 'PROFILE',           // プロフィール写真
    EXTERIOR: 'EXTERIOR',         // 外観
    INTERIOR: 'INTERIOR',         // 内観
    PRODUCT: 'PRODUCT',           // 商品
    AT_WORK: 'AT_WORK',           // 作業中
    FOOD_AND_DRINK: 'FOOD_AND_DRINK', // 食べ物・飲み物
    MENU: 'MENU',                 // メニュー
    ADDITIONAL: 'ADDITIONAL'      // その他（デフォルト）
  },

  // アップロード対象のファイル拡張子
  ALLOWED_EXTENSIONS: ['jpg', 'jpeg', 'png'],

  // アップロード済みファイルを移動するフォルダ名
  UPLOADED_FOLDER_NAME: 'uploaded',

  // ========== 入稿処理システム ==========

  // 入稿フォルダの名前（Driveのルート直下または任意フォルダID）
  INBOX_FOLDER_ID: '',  // ★ 入稿フォルダのDrive IDを設定してください（空の場合はルートを使用）

  // Driveフォルダ名（フォルダIDが空の場合に使用）
  INBOX_FOLDER_NAME: '入稿',
  RESULTS_FOLDER_NAME: '処理結果',
  IMAGE_ROOT_FOLDER_NAME: '商品画像',

  // 入稿ファイルのヘッダ行の構造（キー: 行番号(1始まり), 値: フィールド名）
  INPUT_HEADER_ROWS: {
    1: 'businessGroupId',    // ビジネスグループID
    2: 'businessGroupName',  // ビジネスグループ名
    3: 'googleAccountId',    // GoogleアカウントID
    4: 'pass'               // PASS
  },

  // 入稿ファイルの明細ヘッダー行番号（この行がカラム名行）
  INPUT_DETAIL_HEADER_ROW: 6,

  // 入稿ファイルの明細カラム名と内部フィールドのマッピング
  INPUT_DETAIL_COLS: {
    'ビジネス名':               'businessName',
    '店舗コード':               'storeCode',
    '商品カテゴリ':              'productCategory',
    '商品・サービス名':          'productName',
    '商品の説明':               'description',
    '商品価格':                 'price',
    'ボタン追加':               'buttonType',
    '商品のランディングページURL': 'landingPageUrl',
    '画像ファイルパス':           'imagePath'
  },

  // GBP ローカルポスト「ボタン追加」の種別マッピング
  BUTTON_TYPE_MAP: {
    '予約':      'BOOK',
    '注文':      'ORDER',
    '詳細':      'LEARN_MORE',
    '購入':      'BUY',
    '特典を入手': 'GET_OFFER',
    '登録':      'SIGN_UP',
    '電話':      'CALL',
    'なし':      ''
  },

  // ログシートのカラム定義
  LOG_COLS: [
    '処理日時',
    'ビジネスグループID',
    'ビジネスグループ名',
    'ビジネス名',
    '店舗コード',
    '商品・サービス名',
    '結果',          // 成功 / 失敗 / スキップ
    'GBP投稿ID',
    'エラー詳細',
    '入稿ファイル名'
  ],

  // 処理済みファイルの保管フォルダ名（入稿フォルダ内に作成）
  PROCESSED_FOLDER_NAME: '処理済み',
  ERROR_FOLDER_NAME: 'エラー',

  // 店舗設定シートの列インデックス（0始まり）
  STORE_COLS: {
    STORE_NAME: 0,       // A: 店舗名
    ACCOUNT_ID: 1,       // B: GBPアカウントID (accounts/XXXXX)
    LOCATION_ID: 2,      // C: GBPロケーションID (locations/XXXXX)
    FOLDER_ID: 3,        // D: Driveフォルダ ID
    CATEGORY: 4,         // E: 写真カテゴリ
    LAST_UPLOAD: 5,      // F: 最終アップロード日時
    STATUS: 6            // G: ステータス（有効/無効）
  },

  // アップロード履歴シートの列インデックス（0始まり）
  HISTORY_COLS: {
    TIMESTAMP: 0,   // A: 日時
    STORE_NAME: 1,  // B: 店舗名
    FILE_NAME: 2,   // C: ファイル名
    FILE_ID: 3,     // D: DriveファイルID
    RESULT: 4,      // E: 結果
    MESSAGE: 5      // F: メッセージ
  }
};
