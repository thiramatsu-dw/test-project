/**
 * Triggers.js
 * 定期実行トリガーの設定・管理
 */

/**
 * 毎日深夜2時に uploadAllStoreImages を実行するトリガーを設定する
 * 既存のトリガーがある場合は削除してから再設定する
 */
function setupDailyTrigger() {
  _deleteTriggersByFunction('uploadAllStoreImages');

  ScriptApp.newTrigger('uploadAllStoreImages')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  Logger.log('毎日 2:00 (JST) に画像アップロードを実行するトリガーを設定しました。');
}

/**
 * 毎週月曜日の深夜2時に uploadAllStoreImages を実行するトリガーを設定する
 */
function setupWeeklyTrigger() {
  _deleteTriggersByFunction('uploadAllStoreImages');

  ScriptApp.newTrigger('uploadAllStoreImages')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(2)
    .create();

  Logger.log('毎週月曜 2:00 (JST) に画像アップロードを実行するトリガーを設定しました。');
}

/**
 * 指定した間隔（時間）で uploadAllStoreImages を実行するトリガーを設定する
 * @param {number} hours - 実行間隔（時間）。1, 2, 4, 6, 8, 12 のいずれか
 */
function setupHourlyTrigger(hours) {
  var validHours = [1, 2, 4, 6, 8, 12];
  if (!hours || validHours.indexOf(hours) === -1) {
    Logger.log('有効な時間を指定してください: ' + validHours.join(', '));
    return;
  }

  _deleteTriggersByFunction('uploadAllStoreImages');

  ScriptApp.newTrigger('uploadAllStoreImages')
    .timeBased()
    .everyHours(hours)
    .create();

  Logger.log(hours + '時間ごとに画像アップロードを実行するトリガーを設定しました。');
}

/**
 * uploadAllStoreImages に関連する全トリガーを削除する
 */
function deleteAllTriggers() {
  var deleted = _deleteTriggersByFunction('uploadAllStoreImages');
  Logger.log(deleted + ' 件のトリガーを削除しました。');
}

/**
 * 現在設定されているトリガーの一覧をログに出力する
 */
function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    Logger.log('トリガーは設定されていません。');
    return;
  }

  Logger.log('=== 現在のトリガー一覧 ===');
  triggers.forEach(function(trigger) {
    Logger.log(
      '関数: ' + trigger.getHandlerFunction() +
      ' | タイプ: ' + trigger.getEventType() +
      ' | ID: ' + trigger.getUniqueId()
    );
  });
}

/**
 * 指定した関数名のトリガーを全て削除するヘルパー関数
 * @param {string} functionName - 削除対象の関数名
 * @returns {number} 削除した件数
 */
function _deleteTriggersByFunction(functionName) {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });

  return count;
}
