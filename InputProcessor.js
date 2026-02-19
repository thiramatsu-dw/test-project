/**
 * InputProcessor.js
 * 入稿ファイル（Google Spreadsheet / Excel / CSV）のパース処理
 *
 * 入稿ファイルの構造:
 *   行1: A="ビジネスグループID"  B=[値]
 *   行2: A="ビジネスグループ名"  B=[値]
 *   行3: A="GoogleアカウントID"  B=[値]
 *   行4: A="PASS"               B=[値]
 *   行5: (空行)
 *   行6: 明細ヘッダー行（ビジネス名 | 店舗コード | ...）
 *   行7以降: 明細データ
 */

var InputProcessor = (function() {

  /**
   * Drive ファイルを読み込んでパースし、入稿データを返す
   *
   * @param {File} file - Google Drive のファイルオブジェクト
   * @returns {{ header: Object|null, details: Array, errors: Array<string> }}
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
          errors: ['未対応のファイル形式です: ' + mimeType +
                   '（対応形式: Google Sheets / Excel (.xlsx) / CSV）']
        };
      }
    } catch (e) {
      return {
        header: null,
        details: [],
        errors: ['ファイルの読み込みに失敗しました: ' + e.message]
      };
    }

    return _buildInputData(rows, file.getName());
  }

  // ===== ファイル読み込み =====

  /** Google Spreadsheet の先頭シートを2次元配列で返す */
  function _readGoogleSheet(file) {
    var ss      = SpreadsheetApp.openById(file.getId());
    var sheet   = ss.getSheets()[0];
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) return [];
    return sheet.getRange(1, 1, lastRow, lastCol).getValues();
  }

  /**
   * Excel ファイルを Google Sheets に変換して読み込む
   * 変換後の一時ファイルは必ず削除する
   */
  function _readExcel(file) {
    var blob      = file.getBlob();
    var tempMeta  = { title: '__tmp_' + file.getId(), mimeType: MimeType.GOOGLE_SHEETS };
    var converted = Drive.Files.insert(tempMeta, blob, { convert: true });
    var rows      = [];

    try {
      var ss      = SpreadsheetApp.openById(converted.id);
      var sheet   = ss.getSheets()[0];
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow > 0 && lastCol > 0) {
        rows = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      }
    } finally {
      try { DriveApp.getFileById(converted.id).setTrashed(true); } catch (e) {}
    }
    return rows;
  }

  /** CSV / テキストファイルを2次元配列で返す（UTF-8 のみ対応） */
  function _readCsv(file) {
    var content = file.getBlob().getDataAsString('UTF-8');
    return Utilities.parseCsv(content);
  }

  // ===== データ構築 =====

  /**
   * 2次元配列からヘッダー・明細データを構築して返す
   *
   * @param {Array<Array>} rows     - ファイル全行
   * @param {string}       fileName - ログ用ファイル名
   * @returns {{ header, details, errors }}
   */
  function _buildInputData(rows, fileName) {
    var errors = [];

    if (!rows || rows.length < CONFIG.INPUT_DETAIL_HEADER_ROW) {
      return {
        header: null,
        details: [],
        errors: [
          'ファイルの行数が不足しています（最低 ' + CONFIG.INPUT_DETAIL_HEADER_ROW + ' 行必要）: ' + fileName
        ]
      };
    }

    // ===== ヘッダー部の解析（行1〜4） =====
    var header = _parseHeader(rows);

    // 必須ヘッダーの検証
    if (!header.businessGroupId) errors.push('ビジネスグループID（1行目B列）が未入力です');
    if (!header.googleAccountId) errors.push('GoogleアカウントID（3行目B列）が未入力です');

    // ===== 明細ヘッダー行の解析（行6） =====
    var detailHeaderRow = rows[CONFIG.INPUT_DETAIL_HEADER_ROW - 1];
    var colIndexMap     = _buildColIndexMap(detailHeaderRow);

    if (Object.keys(colIndexMap).length === 0) {
      errors.push(
        '明細ヘッダー行（' + CONFIG.INPUT_DETAIL_HEADER_ROW + '行目）が読み取れません。' +
        '列名がテンプレートと一致しているか確認してください。'
      );
      return { header: header, details: [], errors: errors };
    }

    // ===== 明細データの解析（行7以降） =====
    var details = [];
    for (var i = CONFIG.INPUT_DETAIL_HEADER_ROW; i < rows.length; i++) {
      var dataRow = rows[i];

      // 全列空の行はスキップ
      if (_isEmptyRow(dataRow)) continue;

      var detail    = _parseDetailRow(dataRow, colIndexMap, i + 1);
      var rowErrors = _validateDetail(detail, i + 1);

      details.push({ rowNum: i + 1, data: detail, errors: rowErrors });
    }

    if (details.length === 0) {
      errors.push('明細データが1件もありません（7行目以降に入力してください）');
    }

    return { header: header, details: details, errors: errors };
  }

  /**
   * ヘッダー部（行1〜4）を解析してオブジェクトを返す
   *
   * ★ B列（インデックス1）の値のみを読む。
   *   B列が空でも A列にフォールバックしない（誤ってラベルを値として読むことを防ぐ）。
   */
  function _parseHeader(rows) {
    var header = {};

    for (var rowNumStr in CONFIG.INPUT_HEADER_ROWS) {
      var fieldName = CONFIG.INPUT_HEADER_ROWS[rowNumStr];
      var rowIdx    = parseInt(rowNumStr) - 1;
      var rowData   = rows[rowIdx];

      // B列（インデックス1）の値を取得。列がなければ空文字
      var value = '';
      if (rowData && rowData.length > 1) {
        var raw = rowData[1];
        if (raw !== null && raw !== undefined) {
          value = String(raw).trim();
          // スプレッドシートのデフォルト値（例: "（ここに値を入力）"）は空扱い
          if (value.startsWith('（') || value === 'accounts/XXXXXX') value = '';
        }
      }

      header[fieldName] = value;
    }

    return header;
  }

  /**
   * 明細ヘッダー行から「カラム表示名 → 列インデックス」マップを構築する
   */
  function _buildColIndexMap(headerRow) {
    var map     = {};
    var colDefs = CONFIG.INPUT_DETAIL_COLS;
    if (!headerRow) return map;

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
  function _parseDetailRow(dataRow, colIndexMap, rowNum) {
    var detail = { rowNum: rowNum };

    Object.keys(colIndexMap).forEach(function(field) {
      var idx = colIndexMap[field];
      var raw = (idx < dataRow.length) ? dataRow[idx] : '';
      detail[field] = (raw === null || raw === undefined) ? '' : String(raw).trim();
    });

    // 商品価格を数値に変換（カンマ・円マーク・スペースを除去）
    var priceStr = detail.price || '';
    detail.price = parseFloat(priceStr.replace(/[^\d.]/g, '')) || 0;

    // ボタン種別を GBP API の actionType 値に変換
    var btnLabel = detail.buttonType || '';
    var mapped   = CONFIG.BUTTON_TYPE_MAP[btnLabel];
    detail.buttonTypeApi = (mapped !== undefined) ? mapped : btnLabel;

    return detail;
  }

  /**
   * 1明細行のバリデーションを行い、エラーメッセージ配列を返す（空なら正常）
   */
  function _validateDetail(detail, rowNum) {
    var errors = [];
    var prefix = rowNum + '行目: ';

    if (!detail.businessName)
      errors.push(prefix + 'ビジネス名が未入力です');
    if (!detail.storeCode)
      errors.push(prefix + '店舗コードが未入力です');
    if (!detail.productName)
      errors.push(prefix + '商品・サービス名が未入力です');
    if (!detail.description)
      errors.push(prefix + '商品の説明が未入力です');

    if (detail.landingPageUrl && !/^https?:\/\//i.test(detail.landingPageUrl)) {
      errors.push(prefix + 'ランディングページURLの形式が不正です（http/https で始まる必要があります）: ' +
                  detail.landingPageUrl);
    }

    if (detail.buttonType &&
        detail.buttonType !== 'なし' &&
        detail.buttonType !== '' &&
        !detail.landingPageUrl) {
      errors.push(prefix + 'ボタンを設定する場合はランディングページURLが必須です');
    }

    return errors;
  }

  /** 全セルが空の行かどうかを判定する */
  function _isEmptyRow(row) {
    if (!row || row.length === 0) return true;
    return row.every(function(cell) {
      return cell === null || cell === undefined || String(cell).trim() === '';
    });
  }

  // Public API
  return { parse: parse };

})();
