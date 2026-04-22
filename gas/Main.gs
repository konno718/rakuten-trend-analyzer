// Main.gs - メインオーケストレーター

/**
 * メイン処理（毎日深夜に自動実行）
 * 設定: createDailyTrigger() を一度だけ実行してトリガーを登録
 */
function runDailyCollection() {
  var startTime = new Date();
  var dateStr   = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy-MM-dd');
  Logger.log('=== 楽天トレンドワード収集開始: ' + dateStr + ' ===');

  var config = getConfig();
  if (!config.rakutenAppId) {
    throw new Error('RAKUTEN_APP_ID が未設定です。setupApiKeys() を実行してください。');
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  initExcludeWordsSheet(ss);

  var genreConfigs = readGenreConfigs();
  if (genreConfigs.length === 0) {
    throw new Error('有効なジャンル設定がありません。設定シートを確認してください。');
  }

  var excludeMap    = loadExcludeWords();
  var pastKeywordMap = getPastKeywordMap(30);

  Logger.log('除外ワード数: ' + Object.keys(excludeMap).length);

  var allResults = [];

  for (var i = 0; i < genreConfigs.length; i++) {
    var gc = genreConfigs[i];
    try {
      Logger.log('処理中: ' + gc.genreName);

      var genreId = parseGenreIdFromUrl(gc.rakutenUrl);
      if (!genreId) {
        Logger.log('ジャンルID取得失敗: ' + gc.rakutenUrl);
        continue;
      }

      var items     = fetchRakutenRanking(config.rakutenAppId, genreId, RANKING_TOP_N);
      if (items.length === 0) { continue; }

      var processed = processItems(items, excludeMap);
      var results   = aggregateKeywords(processed, gc.genreName, dateStr);

      for (var j = 0; j < results.length; j++) {
        var key = results[j].genre + '::' + results[j].keyword;
        results[j].isNew = !pastKeywordMap[key];
      }

      for (var k = 0; k < results.length; k++) allResults.push(results[k]);
      Utilities.sleep(1000);

    } catch(e) {
      Logger.log(gc.genreName + ' エラー: ' + e);
      if (config.discordWebhook) sendErrorAlert(config.discordWebhook, gc.genreName + ' エラー: ' + e);
    }
  }

  if (allResults.length === 0) {
    Logger.log('結果が0件です');
    return;
  }

  var newEntrantMap = {};
  for (var n = 0; n < allResults.length; n++) {
    if (allResults[n].isNew) newEntrantMap[allResults[n].genre + '::' + allResults[n].keyword] = true;
  }

  writeWordStats(allResults, dateStr, newEntrantMap);
  writeProducts(allResults, dateStr);

  var candidates = detectHighFrequencyCandidates(allResults, genreConfigs.length);
  writeExcludeCandidates(candidates, dateStr);

  if (config.discordWebhook) sendDailyDigest(config.discordWebhook, allResults, dateStr);

  var elapsed = (new Date() - startTime) / 1000;
  Logger.log('=== 完了: ' + allResults.length + 'キーワード / ' + elapsed.toFixed(1) + '秒 ===');
}

// テスト実行（1ジャンル・30件・スプシ書き込みなし）
function runTest() {
  var config = getConfig();
  if (!config.rakutenAppId) throw new Error('setupApiKeys() を先に実行してください。');

  var genreConfigs = readGenreConfigs();
  if (genreConfigs.length === 0) throw new Error('設定シートにURLを入力してください。');

  var gc      = genreConfigs[0];
  var genreId = parseGenreIdFromUrl(gc.rakutenUrl);
  var items   = fetchRakutenRanking(config.rakutenAppId, genreId, 30);

  Logger.log('取得件数: ' + items.length);
  if (items.length > 0) Logger.log('サンプル: ' + items[0].itemName);

  var excludeMap = loadExcludeWords();
  var processed  = processItems(items, excludeMap);
  var results    = aggregateKeywords(processed, gc.genreName, 'test');

  Logger.log('キーワード数: ' + results.length);
  var top = results.slice(0, 10);
  for (var i = 0; i < top.length; i++) {
    Logger.log('  [' + top[i].classification + '] ' + top[i].keyword + ' - ' + top[i].count + '回 スコア:' + top[i].finalScore);
  }
  Logger.log('テスト完了');
}

// 初期セットアップ（最初に一度だけ実行）
function runSetup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  createSettingsTemplate(ss);
  initExcludeWordsSheet(ss);
  Logger.log('セットアップ完了。設定シートにジャンルURLを入力してください。');
}

// 自動トリガー登録（一度だけ実行）
function createDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDailyCollection') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('runDailyCollection')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  Logger.log('毎日午前2時のトリガーを設定しました');
}
