/**
 * ExecutionLogger.js
 * 処理結果のログファイル（Google Spreadsheet）を生成・管理するモジュール
 *
 * 出力先: 処理結果フォルダ内に「実行ログ_YYYYMMDD_HHMMSS.xlsx」相当のシートを作成
 */

var ExecutionLogger = (function() {

  var _ss = null;        // ログ用 Spreadsheet
  var _sheet = null;     // ログシート
  var _summarySheet = null; // サマリーシート
  var _stats = null;     // 統計カウンター

  /**
   * ログセッションを初期化する
   * 処理開始時に1回呼び出す
   *
   * @param {Folder} resultsFolder - ログファイルを作成するDriveフォルダ
   * @param {string} inputFileName - 入稿ファイル名（ログファイル名に使用）
   * @returns {Object} logger インスタンス（addRow, finalize を持つ）
   */
  function create(resultsFolder, inputFileName) {
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    var baseName = inputFileName.replace(/\.[^.]+$/, ''); // 拡張子除去
    var logName = '実行ログ_' + baseName + '_' + timestamp;

    _ss = SpreadsheetApp.create(logName);
    _sheet = _ss.getActiveSheet();
    _sheet.setName('実行詳細');
    _summarySheet = _ss.insertSheet('サマリー');

    _stats = { total: 0, success: 0, skip: 0, error: 0 };

    // 詳細シートのヘッダー設定
    _setupDetailHeader(_sheet);

    // ログファイルをDriveフォルダへ移動
    var logFile = DriveApp.getFileById(_ss.getId());
    resultsFolder.addFile(logFile);
    DriveApp.getRootFolder().removeFile(logFile); // マイドライブのルートから削除

    Logger.log('実行ログを作成しました: ' + logName + ' (ID: ' + _ss.getId() + ')');

    return {
      addRow: addRow,
      addHeaderError: addHeaderError,
      finalize: finalize,
      getSpreadsheetUrl: getSpreadsheetUrl
    };
  }

  /**
   * 詳細シートのヘッダー行を設定する
   */
  function _setupDetailHeader(sheet) {
    var headers = CONFIG.LOG_COLS;
    var range = sheet.getRange(1, 1, 1, headers.length);
    range.setValues([headers]);
    range.setBackground('#1a73e8');
    range.setFontColor('#ffffff');
    range.setFontWeight('bold');
    sheet.setFrozenRows(1);

    // 列幅の設定
    var widths = [160, 130, 160, 150, 100, 180, 70, 250, 350, 200];
    headers.forEach(function(_, i) {
      if (widths[i]) sheet.setColumnWidth(i + 1, widths[i]);
    });
  }

  /**
   * 1件の処理結果をログシートに追加する
   *
   * @param {Object} header   - 入稿ヘッダー情報
   * @param {Object} detail   - 明細フィールドオブジェクト
   * @param {string} result   - 'SUCCESS' | 'ERROR' | 'SKIP'
   * @param {string} postId   - 作成された GBP 投稿ID（成功時）
   * @param {string} message  - エラー詳細またはスキップ理由
   * @param {string} fileName - 入稿ファイル名
   */
  function addRow(header, detail, result, postId, message, fileName) {
    if (!_sheet) return;

    _stats.total++;
    if (result === 'SUCCESS') _stats.success++;
    else if (result === 'SKIP') _stats.skip++;
    else _stats.error++;

    var label = result === 'SUCCESS' ? '成功' : result === 'SKIP' ? 'スキップ' : '失敗';

    var row = [
      new Date(),
      (header && header.businessGroupId)   || '',
      (header && header.businessGroupName) || '',
      detail.businessName  || '',
      detail.storeCode     || '',
      detail.productName   || '',
      label,
      postId   || '',
      message  || '',
      fileName || ''
    ];

    var lastRow = _sheet.getLastRow() + 1;
    _sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);

    // 結果セルの色付け
    var resultCell = _sheet.getRange(lastRow, 7);
    if (result === 'SUCCESS') {
      resultCell.setBackground('#d9ead3');
    } else if (result === 'SKIP') {
      resultCell.setBackground('#fff2cc');
    } else {
      resultCell.setBackground('#fce5cd');
    }
  }

  /**
   * 入稿ヘッダーレベルのエラーをログシートに追加する
   */
  function addHeaderError(message, fileName) {
    if (!_sheet) return;

    _stats.total++;
    _stats.error++;

    var row = [
      new Date(),
      '', '', '', '', '',
      '失敗',
      '',
      '[ヘッダーエラー] ' + message,
      fileName || ''
    ];

    var lastRow = _sheet.getLastRow() + 1;
    _sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);
    _sheet.getRange(lastRow, 7).setBackground('#ea4335').setFontColor('#ffffff');
  }

  /**
   * ログセッションを終了し、サマリーシートを作成する
   */
  function finalize() {
    if (!_ss || !_summarySheet) return;

    _buildSummarySheet(_summarySheet);

    // シートの順序をサマリーを先頭に
    _ss.setActiveSheet(_summarySheet);
    _ss.moveActiveSheet(1);

    SpreadsheetApp.flush();
    Logger.log('実行ログを確定しました。成功: ' + _stats.success +
      ', スキップ: ' + _stats.skip + ', 失敗: ' + _stats.error + ', 合計: ' + _stats.total);
  }

  /**
   * サマリーシートを作成する
   */
  function _buildSummarySheet(sheet) {
    sheet.setName('サマリー');

    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    var successRate = _stats.total > 0
      ? Math.round((_stats.success / _stats.total) * 100) + '%'
      : '-';

    var summaryData = [
      ['Google ビジネスプロフィール 商品登録 実行ログ'],
      [],
      ['実行日時',   now],
      [],
      ['処理結果サマリー', ''],
      ['合計件数',   _stats.total],
      ['成功',       _stats.success],
      ['スキップ',   _stats.skip],
      ['失敗',       _stats.error],
      ['成功率',     successRate]
    ];

    sheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);

    // スタイル設定
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
    sheet.getRange(5, 1).setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange(6, 1, 5, 1).setFontWeight('bold');

    // 成功件数を緑、失敗件数を赤
    sheet.getRange(7, 2).setBackground('#d9ead3');
    sheet.getRange(9, 2).setBackground('#fce5cd');

    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 150);
  }

  /**
   * 実行ログスプレッドシートのURLを返す
   */
  function getSpreadsheetUrl() {
    return _ss ? _ss.getUrl() : '';
  }

  // Public API
  return { create: create };

})();
