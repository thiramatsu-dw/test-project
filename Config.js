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
