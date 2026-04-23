// HiddenGem.gs - お宝分析（ランキング × 語彙プール 突合 + AI分析 + 書き出し）

/**
 * お宝分析メイン処理
 * 1. 当日のランキング（データ_中国輸入/国内メーカー）からワード取得
 * 2. 消耗品ワード除外
 * 3. 語彙プールと突合して区分判定（メインワード / お宝候補）
 * 4. Claude API でメイン+サブワード判定
 * 5. 調査済み（永久非表示）除外 + グレーアウト情報付与
 * 6. 出現回数desc・同数ならお宝優先でジャンル内ソート、上位50件
 * 7. NEW判定（過去14日以内の初出）
 * 8. お宝分析シート（蓄積式）に書き出し
 */
function runHiddenGemAnalysis() {
  var startTime = new Date();
  var dateStr = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy-MM-dd');
  var currentMonth = startTime.getMonth() + 1;
  Logger.log('=== お宝分析開始: ' + dateStr + ' ===');

  var config = getConfig();
  var disposables = loadDisposableWords();
  var wordPool    = loadWordPoolByGenre();
  var surveyed    = loadSurveyedWords();

  Logger.log('消耗品ワード ' + Object.keys(disposables).length
             + ' / 語彙プール ' + Object.keys(wordPool).length + 'ジャンル'
             + ' / 調査済み ' + Object.keys(surveyed).length);

  var genreConfigs = readGenreConfigs();
  var allRows = [];

  for (var gi = 0; gi < genreConfigs.length; gi++) {
    var gc = genreConfigs[gi];
    var rankWords = getRankingWordsForGenre(gc.genreName, gc.mode, dateStr);
    if (rankWords.length === 0) {
      Logger.log('[' + gc.genreName + '] 当日ランキングワードなし');
      continue;
    }

    // 消耗品ワード除外（メインワードが消耗品なら落とす）
    rankWords = rankWords.filter(function(w) { return !disposables[w.word]; });

    // プール判定
    var genrePool = wordPool[gc.genreName] || {};
    var classified = rankWords.map(function(w) {
      return {
        word           : w.word,
        subWords       : w.synonyms || [],
        count          : w.count,
        products       : w.products || [],
        classification : w.classification,
        isMainWord     : !!genrePool[w.word],
      };
    });

    // AI分析（Claude）
    if (config.claudeApiKey) {
      classified = analyzeWithClaude(config.claudeApiKey, gc.genreName, classified);
    } else {
      Logger.log('CLAUDE_API_KEY 未設定。AI分析スキップ（ワードそのまま使用）');
    }

    // 調査済み「永久非表示」除外
    classified = classified.filter(function(x) {
      var entry = surveyed[gc.genreName + '::' + x.word];
      if (entry && entry.status === '永久非表示') return false;
      return true;
    });

    // 順位決め: 出現回数desc / 同数ならお宝優先
    classified.sort(function(a, b) {
      if (b.count !== a.count) return b.count - a.count;
      if (!a.isMainWord && b.isMainWord) return -1;
      if (a.isMainWord && !b.isMainWord) return 1;
      return 0;
    });

    var top = classified.slice(0, HIDDEN_GEM_MAX_PER_GENRE);

    // 各エントリに表示用属性を付与
    for (var t = 0; t < top.length; t++) {
      var x = top[t];
      x.genre = gc.genreName;
      x.evaluation = getClassificationLabel(x.classification);
      x.type = x.isMainWord ? 'メインワード' : 'お宝候補';

      // 背景色（調査済みステータス）
      var surv = surveyed[gc.genreName + '::' + x.word];
      if (surv && surv.status === '毎年X月再表示') {
        x.backgroundColor = (surv.month === currentMonth)
          ? HIDDEN_GEM_COLOR_RECALL
          : HIDDEN_GEM_COLOR_SURVEYED;
      } else {
        x.backgroundColor = null;
      }
      allRows.push(x);
    }
  }

  if (allRows.length === 0) {
    Logger.log('お宝分析: 書き込み対象なし');
    return;
  }

  // NEW判定（過去14日以内に同ジャンル+メインワードが既に存在するか）
  var pastHiddenGem = loadHiddenGemHistory(HIDDEN_GEM_NEW_DAYS);
  for (var r = 0; r < allRows.length; r++) {
    var key = allRows[r].genre + '::' + allRows[r].word;
    var past = pastHiddenGem[key];
    if (!past) {
      allRows[r].newStatus = 'NEW!';
    } else {
      // サブワード集合に1つでも新規があれば 🔄更新
      var newSub = false;
      var pastSubSet = {};
      for (var p = 0; p < (past.subWords || []).length; p++) pastSubSet[past.subWords[p]] = true;
      for (var s = 0; s < (allRows[r].subWords || []).length; s++) {
        if (!pastSubSet[allRows[r].subWords[s]]) { newSub = true; break; }
      }
      allRows[r].newStatus = newSub ? '🔄更新' : '既存';
    }
  }

  writeHiddenGemResults(allRows, dateStr);
  var elapsed = (new Date() - startTime) / 1000;
  Logger.log('=== お宝分析完了: ' + allRows.length + '行 / ' + elapsed.toFixed(1) + '秒 ===');
}

function getClassificationLabel(c) {
  if (c === 'hidden_gem') return SCORE_RULES.HIDDEN_GEM.label;
  if (c === 'trending')   return SCORE_RULES.TRENDING.label;
  if (c === 'saturated')  return SCORE_RULES.SATURATED.label;
  return SCORE_RULES.HIDDEN_GEM.label;  // お宝候補はデフォルトHIDDEN_GEM
}

/**
 * 当日の指定ジャンル・モードのデータシートからワード一覧を取得
 * 同ジャンルの全行をもらい、代表キーワード単位でまとめる
 */
function getRankingWordsForGenre(genreName, mode, dateStr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(getDataSheetName(mode));
  if (!ws || ws.getLastRow() < 2) return [];
  var data = ws.getDataRange().getValues();

  // 列: 日付|ジャンル|代表キーワード|類義ワード|出現回数|平均順位|スコア|分類|新規参入|1位順位|1位URL|...|5位URL
  var words = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var rowDate = (row[0] instanceof Date)
      ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(row[0]).substring(0, 10);
    if (rowDate !== dateStr) continue;
    if (String(row[1]).trim() !== genreName) continue;

    var products = [];
    for (var p = 0; p < PRODUCTS_PER_KEYWORD; p++) {
      var base = 9 + p * 2;  // 1位順位=列10(index9), 1位URL=列11(index10)
      var rank = row[base];
      var url  = row[base + 1];
      if (rank && url) products.push({ rank: rank, itemUrl: url });
    }

    var synonymsStr = String(row[3] || '').trim();
    var synonyms = synonymsStr ? synonymsStr.split(/[,，、]/).map(function(s){return s.trim();}).filter(function(s){return s;}) : [];

    words.push({
      word           : String(row[2] || '').trim(),
      synonyms       : synonyms,
      count          : Number(row[4] || 0),
      classification : mapClassificationLabelToKey(String(row[7] || '')),
      products       : products,
    });
  }
  return words;
}

function mapClassificationLabelToKey(label) {
  if (label.indexOf('隠れた') >= 0) return 'hidden_gem';
  if (label.indexOf('注目')   >= 0) return 'trending';
  if (label.indexOf('飽和')   >= 0) return 'saturated';
  return 'hidden_gem';
}

/**
 * 過去 daysBack 日以内のお宝分析シートから {ジャンル::メインワード → {subWords}} を読む
 * NEW判定・🔄更新判定用
 */
function loadHiddenGemHistory(daysBack) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.HIDDEN_GEM);
  if (!ws || ws.getLastRow() < 2) return {};
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysBack);
  var data = ws.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var d = (row[0] instanceof Date) ? row[0] : new Date(row[0]);
    if (d < cutoff) continue;
    var genre = String(row[3] || '').trim();
    var word  = String(row[4] || '').trim();
    if (!genre || !word) continue;
    var subStr = String(row[5] || '');
    var subs = subStr ? subStr.split(/[,，、]/).map(function(s){return s.trim();}).filter(function(s){return s;}) : [];
    map[genre + '::' + word] = { subWords: subs, date: d };
  }
  return map;
}

/**
 * お宝分析シートに行追加（蓄積式・背景色もセット）
 */
function writeHiddenGemResults(rows, dateStr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.HIDDEN_GEM);
  if (!ws) ws = initHiddenGemSheet(ss);

  var totalCols = 8 + HIDDEN_GEM_URL_COUNT;
  var values = [];
  var bgColors = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var subStr = (r.subWords || []).slice(0, 10).join(', ');
    var row = [
      dateStr,
      r.count || 0,
      r.type || 'お宝候補',
      r.genre || '',
      r.word || '',
      subStr,
      r.evaluation || SCORE_RULES.HIDDEN_GEM.label,
      r.newStatus || '既存',
    ];
    for (var p = 0; p < HIDDEN_GEM_URL_COUNT; p++) {
      row.push(r.products && r.products[p] ? r.products[p].itemUrl : '');
    }
    values.push(row);

    // 背景色行
    var rowBg = [];
    for (var b = 0; b < totalCols; b++) rowBg.push(r.backgroundColor || null);
    bgColors.push(rowBg);
  }

  var startRow = ws.getLastRow() + 1;
  ws.getRange(startRow, 1, values.length, totalCols).setValues(values);
  ws.getRange(startRow, 1, bgColors.length, totalCols).setBackgrounds(bgColors);

  Logger.log('お宝分析シートに ' + values.length + '行書き込み');
}
