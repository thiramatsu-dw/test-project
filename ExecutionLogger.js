/**
 * ExecutionLogger.js
 * 処理結果のログファイル（Google Spreadsheet）を生成・管理するモジュール
 *
 * ★ ファクトリーパターン採用:
 *   create() が呼ばれるたびに独立した状態を持つロガーインスタンスを返す。
 *   複数ファイルを処理しても、各ロガーが互いに干渉しない。
 */

var ExecutionLogger = (function() {

  /**
   * ログセッションを初期化し、独立したロガーインスタンスを返す
   *
   * @param {Folder} resultsFolder - ログファイルを作成するDriveフォルダ
   * @param {string} inputFileName - 入稿ファイル名（ログファイル名に使用）
   * @returns {{ addRow, addHeaderError, finalize, getSpreadsheetUrl }}
   */
  function create(resultsFolder, inputFileName) {
    // ===== ローカル状態（インスタンスごとに独立） =====
    var _ss           = null;
    var _detailSheet  = null;
    var _summarySheet = null;
    var _stats        = { total: 0, success: 0, skip: 0, error: 0 };

    // ===== 初期化 =====
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    var baseName  = inputFileName.replace(/\.[^.]+$/, '');
    var logName   = '実行ログ_' + baseName + '_' + timestamp;

    _ss           = SpreadsheetApp.create(logName);
    _detailSheet  = _ss.getActiveSheet();
    _detailSheet.setName('実行詳細');
    _summarySheet = _ss.insertSheet('サマリー');

    _setupDetailHeader(_detailSheet);

    // Drive の処理結果フォルダへ移動
    var logFile = DriveApp.getFileById(_ss.getId());
    resultsFolder.addFile(logFile);
    try { DriveApp.getRootFolder().removeFile(logFile); } catch (e) {}

    Logger.log('実行ログを作成しました: ' + logName + ' (ID: ' + _ss.getId() + ')');

    // ===== 詳細シートのヘッダー設定 =====
    function _setupDetailHeader(sheet) {
      var headers = CONFIG.LOG_COLS;
      var range   = sheet.getRange(1, 1, 1, headers.length);
      range.setValues([headers]);
      range.setBackground('#1a73e8');
      range.setFontColor('#ffffff');
      range.setFontWeight('bold');
      sheet.setFrozenRows(1);

      var widths = [160, 130, 160, 150, 100, 180, 70, 250, 350, 200];
      headers.forEach(function(_, i) {
        if (widths[i]) sheet.setColumnWidth(i + 1, widths[i]);
      });
    }

    // ===== 公開メソッド =====

    /**
     * 1件の処理結果をログシートに追記する
     *
     * @param {Object} header   - 入稿ヘッダー情報
     * @param {Object} detail   - 明細フィールドオブジェクト
     * @param {string} result   - 'SUCCESS' | 'ERROR' | 'SKIP'
     * @param {string} postId   - 作成された GBP 投稿ID（成功時）
     * @param {string} message  - エラー詳細またはスキップ理由
     * @param {string} fileName - 入稿ファイル名
     */
    function addRow(header, detail, result, postId, message, fileName) {
      _stats.total++;
      if      (result === 'SUCCESS') _stats.success++;
      else if (result === 'SKIP')    _stats.skip++;
      else                           _stats.error++;

      var label = result === 'SUCCESS' ? '成功' : result === 'SKIP' ? 'スキップ' : '失敗';

      var row = [
        new Date(),
        (header && header.businessGroupId)   || '',
        (header && header.businessGroupName) || '',
        (detail && detail.businessName)      || '',
        (detail && detail.storeCode)         || '',
        (detail && detail.productName)       || '',
        label,
        postId   || '',
        message  || '',
        fileName || ''
      ];

      var nextRow = _detailSheet.getLastRow() + 1;
      _detailSheet.getRange(nextRow, 1, 1, row.length).setValues([row]);

      var resultCell = _detailSheet.getRange(nextRow, CONFIG.LOG_COLS.indexOf('結果') + 1);
      if      (result === 'SUCCESS') resultCell.setBackground('#d9ead3');
      else if (result === 'SKIP')    resultCell.setBackground('#fff2cc');
      else                           resultCell.setBackground('#fce5cd');
    }

    /**
     * 入稿ヘッダーレベルのエラーをログシートに追記する
     *
     * @param {string} message  - エラー内容
     * @param {string} fileName - 入稿ファイル名
     */
    function addHeaderError(message, fileName) {
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

      var nextRow = _detailSheet.getLastRow() + 1;
      _detailSheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
      _detailSheet.getRange(nextRow, 7).setBackground('#ea4335').setFontColor('#ffffff');
    }

    /**
     * ログセッションを終了し、サマリーシートを完成させる
     */
    function finalize() {
      _buildSummarySheet();

      // サマリーシートを先頭に移動
      _ss.setActiveSheet(_summarySheet);
      _ss.moveActiveSheet(1);

      SpreadsheetApp.flush();

      Logger.log(
        '実行ログ確定 | 成功: ' + _stats.success +
        ', スキップ: ' + _stats.skip +
        ', 失敗: '     + _stats.error +
        ', 合計: '     + _stats.total
      );
    }

    function _buildSummarySheet() {
      var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
      var successRate = _stats.total > 0
        ? Math.round((_stats.success / _stats.total) * 100) + '%'
        : '-';

      var data = [
        ['Google ビジネスプロフィール 商品登録 実行ログ', ''],
        ['', ''],
        ['実行日時',         now],
        ['入稿ファイル名',   inputFileName],
        ['', ''],
        ['処理結果サマリー', ''],
        ['合計件数',   _stats.total],
        ['成功',       _stats.success],
        ['スキップ',   _stats.skip],
        ['失敗',       _stats.error],
        ['成功率',     successRate]
      ];

      _summarySheet.setName('サマリー');
      _summarySheet.getRange(1, 1, data.length, 2).setValues(data);

      // スタイル
      _summarySheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
      _summarySheet.getRange(6, 1).setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold');
      _summarySheet.getRange(7, 1, 5, 1).setFontWeight('bold');
      _summarySheet.getRange(8, 2).setBackground('#d9ead3');  // 成功: 緑
      _summarySheet.getRange(10, 2).setBackground('#fce5cd'); // 失敗: 橙
      _summarySheet.setColumnWidth(1, 160);
      _summarySheet.setColumnWidth(2, 200);
    }

    /**
     * 実行ログスプレッドシートの URL を返す
     */
    function getSpreadsheetUrl() {
      return _ss ? _ss.getUrl() : '';
    }

    // インスタンスを返す
    return {
      addRow:            addRow,
      addHeaderError:    addHeaderError,
      finalize:          finalize,
      getSpreadsheetUrl: getSpreadsheetUrl
    };
  }

  // Public API
  return { create: create };

})();
