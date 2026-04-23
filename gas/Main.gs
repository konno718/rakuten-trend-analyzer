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
  initSynonymSheet(ss);

  var genreConfigs = readGenreConfigs();
  if (genreConfigs.length === 0) {
    throw new Error('有効なジャンル設定がありません。設定シートを確認してください。');
  }

  // モード別除外ワード + 同義語マップ
  var excludeMaps = {};
  excludeMaps[MODES.CHINA]    = loadExcludeWords(MODES.CHINA);
  excludeMaps[MODES.DOMESTIC] = loadExcludeWords(MODES.DOMESTIC);
  var synonymMap = loadSynonymMap();
  Logger.log('除外ワード 中国輸入:' + Object.keys(excludeMaps[MODES.CHINA]).length
             + ' 国内:' + Object.keys(excludeMaps[MODES.DOMESTIC]).length
             + ' / 同義語:' + Object.keys(synonymMap).length);

  // モード別 過去キーワード
  var pastMaps = {};
  pastMaps[MODES.CHINA]    = getPastKeywordMap(30, MODES.CHINA);
  pastMaps[MODES.DOMESTIC] = getPastKeywordMap(30, MODES.DOMESTIC);

  var resultsByMode = {};
  resultsByMode[MODES.CHINA]    = [];
  resultsByMode[MODES.DOMESTIC] = [];

  var genreCountByMode = {};
  genreCountByMode[MODES.CHINA]    = 0;
  genreCountByMode[MODES.DOMESTIC] = 0;

  for (var i = 0; i < genreConfigs.length; i++) {
    var gc = genreConfigs[i];
    genreCountByMode[gc.mode]++;
    try {
      Logger.log('処理中: ' + gc.genreName + ' [' + gc.mode + ']');

      var genreId = parseGenreIdFromUrl(gc.rakutenUrl);
      if (!genreId) {
        Logger.log('ジャンルID取得失敗: ' + gc.rakutenUrl);
        continue;
      }

      var items = fetchRakutenRanking(config.rakutenAppId, genreId, RANKING_TOP_N);
      if (items.length === 0) continue;

      var processed = processItems(items, excludeMaps[gc.mode], synonymMap, gc.mode);
      var rawResults = aggregateKeywords(processed, gc.genreName, dateStr);
      // 同一商品セットのワードを代表+類義ワードにグループ化
      var results = groupByProductFingerprint(rawResults);
      Logger.log('  グループ化: ' + rawResults.length + ' → ' + results.length + ' グループ');

      for (var j = 0; j < results.length; j++) {
        var key = results[j].genre + '::' + results[j].keyword;
        results[j].isNew = !pastMaps[gc.mode][key];
      }
      for (var k = 0; k < results.length; k++) resultsByMode[gc.mode].push(results[k]);

      Utilities.sleep(1000);

    } catch(e) {
      Logger.log(gc.genreName + ' エラー: ' + e);
      if (config.discordWebhook) sendErrorAlert(config.discordWebhook, gc.genreName + ' エラー: ' + e);
    }
  }

  var totalResults = resultsByMode[MODES.CHINA].length + resultsByMode[MODES.DOMESTIC].length;
  if (totalResults === 0) {
    Logger.log('結果が0件です');
    return;
  }

  // モード別に書き込み
  var allResults = [];
  var modeKeys = [MODES.CHINA, MODES.DOMESTIC];
  for (var mi = 0; mi < modeKeys.length; mi++) {
    var mk = modeKeys[mi];
    var modeResults = resultsByMode[mk];
    if (modeResults.length === 0) continue;

    var newEntrantMap = {};
    for (var n = 0; n < modeResults.length; n++) {
      if (modeResults[n].isNew) newEntrantMap[modeResults[n].genre + '::' + modeResults[n].keyword] = true;
    }

    writeData(modeResults, dateStr, newEntrantMap, mk);

    var candidates = detectHighFrequencyCandidates(modeResults, Math.max(1, genreCountByMode[mk]));
    writeExcludeCandidates(candidates, dateStr, mk);

    for (var r = 0; r < modeResults.length; r++) allResults.push(modeResults[r]);
  }

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

  Logger.log('取得件数: ' + items.length + ' モード: ' + gc.mode);
  if (items.length > 0) {
    Logger.log('サンプル(元): ' + items[0].itemName);
    Logger.log('サンプル(整形後): ' + cleanTitlePrefix(items[0].itemName));
  }

  var excludeMap = loadExcludeWords(gc.mode);
  var synonymMap = loadSynonymMap();
  var processed  = processItems(items, excludeMap, synonymMap, gc.mode);
  var rawResults = aggregateKeywords(processed, gc.genreName, 'test');
  var results    = groupByProductFingerprint(rawResults);

  Logger.log('キーワード数(グループ化後): ' + results.length + ' / 元: ' + rawResults.length);
  var top = results.slice(0, 10);
  for (var i = 0; i < top.length; i++) {
    var syn = (top[i].synonyms && top[i].synonyms.length > 0) ? ' [類義: ' + top[i].synonyms.slice(0, 3).join(', ') + (top[i].synonyms.length > 3 ? '...' : '') + ']' : '';
    Logger.log('  [' + top[i].classification + '] ' + top[i].keyword + ' - ' + top[i].count + '回 スコア:' + top[i].finalScore + syn);
  }
  Logger.log('テスト完了');
}

// 初期セットアップ（最初に一度だけ実行）
function runSetup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  createSettingsTemplate(ss);
  initExcludeWordsSheet(ss);
  initSynonymSheet(ss);
  migrateExcludeCandidatesTo5Col();
  Logger.log('セットアップ完了。設定シートにジャンルURL・モードを入力してください。');
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
