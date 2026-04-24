// Main.gs - メインオーケストレーター

/**
 * メイン処理（毎日深夜に自動実行）
 * 設定: createDailyTrigger() を一度だけ実行してトリガーを登録
 */
/**
 * 毎日のランキング収集（Step 3 の実装）
 * - 各ジャンルで楽天ランキング300件取得
 * - 商品名先頭クリーンアップ + isProductExcluded (モード別)
 * - 残った商品を「ランキング_*」シートに商品単位で書き込み
 * - キーワード抽出・集計はしない（HiddenGem側で都度実行）
 */
function runDailyCollection() {
  var startTime = new Date();
  var dateStr   = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy-MM-dd');
  Logger.log('=== ランキング収集開始: ' + dateStr + ' ===');

  var config = getConfig();
  if (!config.rakutenAppId) {
    throw new Error('RAKUTEN_APP_ID が未設定です。setupApiKeys() を実行してください。');
  }

  var genreConfigs = readGenreConfigs();
  if (genreConfigs.length === 0) {
    throw new Error('有効なジャンル設定がありません。設定シートを確認してください。');
  }

  // URL でグループ化（同じURLは1回取得 → 該当する全モードに書き込み）
  var byUrl = {};
  var orderedUrls = [];
  for (var ci = 0; ci < genreConfigs.length; ci++) {
    var gcc = genreConfigs[ci];
    var key = gcc.rakutenUrl;
    if (!byUrl[key]) {
      byUrl[key] = { genreName: gcc.genreName, rakutenUrl: gcc.rakutenUrl, modes: [] };
      orderedUrls.push(key);
    }
    if (byUrl[key].modes.indexOf(gcc.mode) < 0) byUrl[key].modes.push(gcc.mode);
  }

  var totalWritten = 0;
  for (var u = 0; u < orderedUrls.length; u++) {
    var entry = byUrl[orderedUrls[u]];
    try {
      Logger.log('処理中: ' + entry.genreName + ' [' + entry.modes.join(',') + ']');

      var genreId = parseGenreIdFromUrl(entry.rakutenUrl);
      if (!genreId) {
        Logger.log('ジャンルID取得失敗: ' + entry.rakutenUrl);
        continue;
      }

      var items = fetchRakutenRanking(config.rakutenAppId, genreId, RANKING_TOP_N);
      if (items.length === 0) continue;

      // 商品名整形のみ
      var filtered = [];
      for (var k = 0; k < items.length; k++) {
        var cleaned = cleanTitlePrefix(items[k].itemName || '');
        filtered.push({
          rank     : items[k].rank,
          itemName : cleaned,
          itemUrl  : items[k].itemUrl,
          tagIds   : items[k].tagIds || [],
        });
      }

      // 該当する全モードのランキングシートに書き込み
      for (var mi = 0; mi < entry.modes.length; mi++) {
        writeRankingItems(filtered, dateStr, entry.genreName, entry.modes[mi]);
        totalWritten += filtered.length;
      }
      Utilities.sleep(1000);

    } catch(e) {
      Logger.log(entry.genreName + ' エラー: ' + e);
      if (config.discordWebhook) sendErrorAlert(config.discordWebhook, entry.genreName + ' エラー: ' + e);
    }
  }

  var elapsed = (new Date() - startTime) / 1000;
  Logger.log('=== ランキング収集完了: 総' + totalWritten + '件 / ' + elapsed.toFixed(1) + '秒 ===');
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
    var cleaned = cleanTitlePrefix(items[0].itemName);
    Logger.log('サンプル(整形後): ' + cleaned);
    Logger.log('除外判定: ' + isProductExcluded(cleaned, gc.mode));
    Logger.log('tagIds数: ' + (items[0].tagIds || []).length);
  }

  var remaining = items.filter(function(it) {
    return !isProductExcluded(cleanTitlePrefix(it.itemName), gc.mode);
  });
  Logger.log('除外後件数: ' + remaining.length + ' / ' + items.length);
  Logger.log('テスト完了');
}

// 初期セットアップ（最初に一度だけ実行）
function runSetup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  createSettingsTemplate(ss);
  initExcludeWordsSheet(ss);
  initSynonymSheet(ss);
  initSurveyedSheet(ss);
  initWordPoolSheet(ss);
  initTagDictSheet(ss);
  initSuggestSheet(ss, MODES.CHINA);
  initSuggestSheet(ss, MODES.DOMESTIC);
  initWordPoolMonthlySheet(ss);
  migrateExcludeCandidatesTo5Col();
  migrateWordPoolToV2();
  seedDisposableWords();
  Logger.log('セットアップ完了。設定シートにジャンルURL・モードを入力してください。');
}

// === 自動トリガー登録 ===

/**
 * 旧runDailyCollection単体トリガー（後方互換・新フローでは使わない）
 */
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
    .atHour(1)
    .create();
  Logger.log('毎日午前1時のトリガーを設定しました（runDailyCollection）');
}

/**
 * お宝分析システム用トリガーを一括登録（朝8時までに分析完了）
 * - 01:00 runDailyCollection (ランキング収集)
 * - 02:00〜06:00 runWordPoolStep (語彙プール構築・チェックポイント方式で継続)
 * - 07:00 runHiddenGemAnalysis (突合+AI+書き出し)
 *
 * 既存の同名ハンドラトリガーは削除してから再登録するので、重複はしない。
 */
function createAnalysisTriggers() {
  var handlers = ['runDailyCollection', 'runWordPoolStep', 'runHiddenGemAnalysis', 'runSuggestCleanup', 'runMonthlyAggregation'];
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (handlers.indexOf(triggers[i].getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 01:00 ランキング収集
  ScriptApp.newTrigger('runDailyCollection').timeBased().everyDays(1).atHour(1).create();

  // 02:00-06:00 語彙プール構築
  for (var h = 2; h <= 6; h++) {
    ScriptApp.newTrigger('runWordPoolStep').timeBased().everyDays(1).atHour(h).create();
  }

  // 07:00 推奨ワード分析
  ScriptApp.newTrigger('runHiddenGemAnalysis').timeBased().everyDays(1).atHour(7).create();

  // 08:00 推奨ワードクリーンアップ（14日超削除・除外一致削除・調査済み削除・日付desc並び替え）
  ScriptApp.newTrigger('runSuggestCleanup').timeBased().everyDays(1).atHour(8).create();

  // 毎週日曜 23:00 月次集計
  ScriptApp.newTrigger('runMonthlyAggregation')
    .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).create();

  Logger.log('分析トリガー登録完了:');
  Logger.log('  01:00 runDailyCollection');
  Logger.log('  02-06:00 runWordPoolStep');
  Logger.log('  07:00 runHiddenGemAnalysis');
  Logger.log('  08:00 runSuggestCleanup');
  Logger.log('  日曜23:00 runMonthlyAggregation');
}

/**
 * 本システムの全トリガー削除（復旧・停止用）
 */
function removeAnalysisTriggers() {
  var handlers = ['runDailyCollection', 'runWordPoolStep', 'runHiddenGemAnalysis', 'runSuggestCleanup', 'runMonthlyAggregation'];
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (handlers.indexOf(triggers[i].getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log('トリガー ' + removed + '本削除');
}
