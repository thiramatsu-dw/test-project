/**
 * InputProcessor.js
 * 入稿ファイル（Google Spreadsheet / Excel / CSV）のパース処理
 *
 * 入稿ファイルの構造:
 *   行1: ビジネスグループID | [値]
 *   行2: ビジネスグループ名  | [値]
 *   行3: GoogleアカウントID  | [値]
 *   行4: PASS               | [値]
 *   行5: (空行)
 *   行6: 明細ヘッダー行（ビジネス名, 店舗コード, ...）
 *   行7以降: 明細データ
 */

var InputProcessor = (function() {

  /**
   * Drive ファイルを読み込んでパースし、入稿データを返す
   * @param {File} file - Google Drive のファイルオブジェクト
   * @returns {{ header: Object, details: Array<Object>, errors: Array<string> }}
   */
  function parse(file) {
    var mimeType = file.getMimeType();
    var rows;

    try {
      if (mimeType === MimeType.GOOGLE_SHEETS) {
        rows = _readGoogleSheet(file);
      } else if (
        mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        mimeType === 'application/vnd.ms-excel'
      ) {
        rows = _readExcel(file);
      } else if (mimeType === 'text/csv' || mimeType === 'text/plain') {
        rows = _readCsv(file);
      } else {
        return {
          header: null,
          details: [],
          errors: ['未対応のファイル形式です: ' + mimeType + '（対応形式: Google Sheets / Excel / CSV）']
        };
      }
    } catch (e) {
      return {
        header: null,
        details: [],
        errors: ['ファイルの読み込みに失敗しました: ' + e.message]
      };
    }

    return _buildInputData(rows);
  }

  /**
   * Google Spreadsheet ファイルを2次元配列として読み込む
   */
  function _readGoogleSheet(file) {
    var ss = SpreadsheetApp.openById(file.getId());
    var sheet = ss.getSheets()[0]; // 先頭シートを対象
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) return [];
    return sheet.getRange(1, 1, lastRow, lastCol).getValues();
  }

  /**
   * Excel ファイルを一時的に Google Sheets に変換して読み込む
   */
  function _readExcel(file) {
    // Drive API で Excel → Google Sheets に変換
    var blob = file.getBlob();
    var tempFile = {
      title: 'temp_' + file.getName(),
      mimeType: MimeType.GOOGLE_SHEETS
    };
    var converted = Drive.Files.insert(tempFile, blob, { convert: true });
    var rows = [];

    try {
      var ss = SpreadsheetApp.openById(converted.id);
      var sheet = ss.getSheets()[0];
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow > 0 && lastCol > 0) {
        rows = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      }
    } finally {
      // 変換した一時ファイルを削除
      try { DriveApp.getFileById(converted.id).setTrashed(true); } catch (e) {}
    }
    return rows;
  }

  /**
   * CSV ファイルを2次元配列として読み込む
   */
  function _readCsv(file) {
    var content = file.getBlob().getDataAsString('UTF-8');
    return Utilities.parseCsv(content);
  }

  /**
   * 2次元配列からヘッダー・明細データを組み立てる
   * @param {Array<Array>} rows
   * @returns {{ header: Object, details: Array<Object>, errors: Array<string> }}
   */
  function _buildInputData(rows) {
    var errors = [];

    if (!rows || rows.length < CONFIG.INPUT_DETAIL_HEADER_ROW) {
      return {
        header: null,
        details: [],
        errors: ['ファイルの行数が不足しています（最低 ' + CONFIG.INPUT_DETAIL_HEADER_ROW + ' 行必要）']
      };
    }

    // ヘッダー部分の解析（行1〜4: "キー | 値" の形式）
    var header = {};
    for (var rowNum in CONFIG.INPUT_HEADER_ROWS) {
      var fieldName = CONFIG.INPUT_HEADER_ROWS[rowNum];
      var row = rows[parseInt(rowNum) - 1];
      // B列（インデックス1）に値があると想定。なければA列
      var value = (row && row.length > 1) ? String(row[1]).trim() : '';
      if (!value && row && row.length > 0) value = String(row[0]).trim();
      header[fieldName] = value;
    }

    // 必須ヘッダーの検証
    if (!header.businessGroupId) errors.push('ビジネスグループIDが未入力です');
    if (!header.googleAccountId) errors.push('GoogleアカウントIDが未入力です');

    // 明細ヘッダー行の解析（行6）
    var detailHeaderRow = rows[CONFIG.INPUT_DETAIL_HEADER_ROW - 1];
    var colIndexMap = _buildColIndexMap(detailHeaderRow);

    if (Object.keys(colIndexMap).length === 0) {
      errors.push('明細ヘッダー行（' + CONFIG.INPUT_DETAIL_HEADER_ROW + '行目）が読み取れません');
      return { header: header, details: [], errors: errors };
    }

    // 明細データの解析（行7以降）
    var details = [];
    for (var i = CONFIG.INPUT_DETAIL_HEADER_ROW; i < rows.length; i++) {
      var row = rows[i];

      // 全列が空の行はスキップ
      if (_isEmptyRow(row)) continue;

      var detail = _parseDetailRow(row, colIndexMap, i + 1);
      var rowErrors = _validateDetail(detail, i + 1);

      details.push({
        rowNum: i + 1,
        data: detail,
        errors: rowErrors
      });
    }

    if (details.length === 0) {
      errors.push('明細データが1件もありません');
    }

    return { header: header, details: details, errors: errors };
  }

  /**
   * 明細ヘッダー行からカラム名→列インデックスのマップを作成する
   */
  function _buildColIndexMap(headerRow) {
    var map = {};
    if (!headerRow) return map;

    var colDefs = CONFIG.INPUT_DETAIL_COLS;
    headerRow.forEach(function(cell, idx) {
      var name = String(cell).trim();
      if (colDefs.hasOwnProperty(name)) {
        map[colDefs[name]] = idx;
      }
    });
    return map;
  }

  /**
   * 1明細行をフィールドオブジェクトに変換する
   */
  function _parseDetailRow(row, colIndexMap, rowNum) {
    var detail = { rowNum: rowNum };

    Object.keys(colIndexMap).forEach(function(field) {
      var idx = colIndexMap[field];
      var raw = (idx < row.length) ? row[idx] : '';
      detail[field] = (raw === null || raw === undefined) ? '' : String(raw).trim();
    });

    // 価格を数値に変換（カンマや円マーク除去）
    if (detail.price) {
      detail.price = parseFloat(String(detail.price).replace(/[^\d.]/g, '')) || 0;
    }

    // ボタン種別を GBP API 形式に変換
    if (detail.buttonType) {
      var mapped = CONFIG.BUTTON_TYPE_MAP[detail.buttonType];
      detail.buttonTypeApi = (mapped !== undefined) ? mapped : detail.buttonType;
    }

    return detail;
  }

  /**
   * 1明細行のバリデーション
   * @returns {Array<string>} エラーメッセージ配列（空なら正常）
   */
  function _validateDetail(detail, rowNum) {
    var errors = [];
    var prefix = rowNum + '行目: ';

    if (!detail.businessName)  errors.push(prefix + 'ビジネス名が未入力です');
    if (!detail.storeCode)     errors.push(prefix + '店舗コードが未入力です');
    if (!detail.productName)   errors.push(prefix + '商品・サービス名が未入力です');
    if (!detail.description)   errors.push(prefix + '商品の説明が未入力です');

    if (detail.landingPageUrl && !/^https?:\/\//i.test(detail.landingPageUrl)) {
      errors.push(prefix + '商品のランディングページURLの形式が不正です: ' + detail.landingPageUrl);
    }

    if (detail.buttonType && detail.buttonType !== 'なし' && !detail.landingPageUrl) {
      errors.push(prefix + 'ボタン追加を指定した場合はランディングページURLが必須です');
    }

    return errors;
  }

  /**
   * 全列が空の行かどうかを判定する
   */
  function _isEmptyRow(row) {
    if (!row) return true;
    return row.every(function(cell) {
      return cell === null || cell === undefined || String(cell).trim() === '';
    });
  }

  // Public API
  return { parse: parse };

})();
