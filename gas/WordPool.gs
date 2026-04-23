// WordPool.gs - 楽天APIから語彙プール（ジャンル別）を構築

/**
 * 語彙プール構築 1ステップ実行
 * - ScriptPropertiesでチェックポイント保持（次に処理するジャンルindex）
 * - 5分（WORDPOOL_STEP_TIME_LIMIT_MS）で中断→次回トリガーで続き
 * - 1日分完了したら翌日まで待機
 * トリガーから定時(毎時)呼び出される想定
 */
function runWordPoolStep() {
  var startTime = new Date();
  var dateStr = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy-MM-dd');
  var props = PropertiesService.getScriptProperties();

  var lastDate = props.getProperty(WP_CHECKPOINT_DATE_KEY);
  var index = parseInt(props.getProperty(WP_CHECKPOINT_INDEX_KEY) || '0', 10);

  // 日が変わったら index リセット
  if (lastDate !== dateStr) {
    index = 0;
    props.setProperty(WP_CHECKPOINT_DATE_KEY, dateStr);
    Logger.log('語彙プール: 新しい日(' + dateStr + ')。index=0 からスタート');
  }

  var config = getConfig();
  if (!config.rakutenAppId) {
    throw new Error('RAKUTEN_APP_ID 未設定。setupApiKeys() を実行してください。');
  }

  var genreConfigs = readGenreConfigs();
  // ジャンルID重複を除去（同じジャンルが複数モードで登録されていても語彙プールは1回で十分）
  var seenGenreIds = {};
  var uniqueConfigs = [];
  for (var gi = 0; gi < genreConfigs.length; gi++) {
    var gid = parseGenreIdFromUrl(genreConfigs[gi].rakutenUrl);
    if (!gid || seenGenreIds[gid]) continue;
    seenGenreIds[gid] = true;
    uniqueConfigs.push(genreConfigs[gi]);
  }
  genreConfigs = uniqueConfigs;
  Logger.log('ユニークジャンル数: ' + genreConfigs.length);

  if (index >= genreConfigs.length) {
    Logger.log('語彙プール: 本日分は全ジャンル完了済み（index=' + index + ' / ' + genreConfigs.length + '）');
    return;
  }

  Logger.log('=== 語彙プール更新: index=' + index + ' から ===');

  while (index < genreConfigs.length) {
    var elapsed = new Date() - startTime;
    if (elapsed > WORDPOOL_STEP_TIME_LIMIT_MS) {
      Logger.log('時間切れ (' + (elapsed / 1000).toFixed(1) + '秒)。次回トリガーで index=' + index + ' から継続');
      break;
    }

    var gc = genreConfigs[index];
    try {
      processGenreForWordPool(gc, config.rakutenAppId, dateStr);
    } catch(e) {
      Logger.log(gc.genreName + ' エラー: ' + e);
    }
    index++;
    props.setProperty(WP_CHECKPOINT_INDEX_KEY, String(index));
  }

  if (index >= genreConfigs.length) {
    Logger.log('語彙プール: 全ジャンル(' + genreConfigs.length + ')完了');
  }
}

/**
 * 1ジャンル分の語彙プール更新
 * 4ソース: ItemSearch上位/ItemRanking(daily,weekly,monthly)/タグ辞書/子ジャンル名
 */
function processGenreForWordPool(gc, appId, dateStr) {
  var genreId = parseGenreIdFromUrl(gc.rakutenUrl);
  if (!genreId) {
    Logger.log('ジャンルID取得失敗: ' + gc.rakutenUrl);
    return;
  }

  Logger.log('[' + gc.genreName + '] 語彙プール更新開始');

  var excludeMap = loadExcludeWords(gc.mode);
  var synonymMap = loadSynonymMap();
  var wordRecords = {};  // key = word + '||' + source → {word, source, hits}
  var tagIdSet = {};

  // --- ソース1: ItemSearch 標準ソート上位（レビュー多い実在商品） ---
  for (var p = 1; p <= WORDPOOL_ITEM_PAGES; p++) {
    var items = fetchItemSearchByGenre(appId, genreId, '-reviewCount', 30, p);
    if (items.length === 0) break;
    collectWordsFromItems(items, 'タイトル派生', wordRecords, tagIdSet, excludeMap, synonymMap);
    Utilities.sleep(RAKUTEN_API_DELAY_MS);
  }

  // --- ソース2: ItemRanking 最新ランキング（複数ページで60位まで） ---
  for (var r = 1; r <= WORDPOOL_RANKING_PAGES; r++) {
    var rankItems = fetchItemRankingByGenre(appId, genreId, r);
    if (rankItems.length === 0) break;
    collectWordsFromItems(rankItems, 'ランキング派生(' + r + 'page)', wordRecords, tagIdSet, excludeMap, synonymMap);
    Utilities.sleep(RAKUTEN_API_DELAY_MS);
  }

  // --- ソース3: タグ辞書 ---
  var tagIds = Object.keys(tagIdSet);
  if (tagIds.length > 0) {
    var tagNameMap = resolveTagNames(tagIds, appId);
    for (var t = 0; t < tagIds.length; t++) {
      var name = tagNameMap[tagIds[t]];
      if (!name) continue;
      addWordRecord(wordRecords, name, '検索タグ');
    }
  }

  // --- ソース4: 子ジャンル名 ---
  var subGenres = fetchGenreChildren(appId, genreId);
  for (var s = 0; s < subGenres.length; s++) {
    addWordRecord(wordRecords, subGenres[s], 'サブジャンル');
  }
  Utilities.sleep(RAKUTEN_API_DELAY_MS);

  // --- 語彙プールシートに反映 ---
  updateWordPoolSheet(gc.genreName, wordRecords, dateStr);

  Logger.log('[' + gc.genreName + '] 語彙プール完了: ' + Object.keys(wordRecords).length + 'レコード');
}

/**
 * 商品群からキーワード抽出 + タグID収集
 */
function collectWordsFromItems(items, sourceLabel, wordRecords, tagIdSet, excludeMap, synonymMap) {
  for (var i = 0; i < items.length; i++) {
    var itemName = items[i].itemName || '';
    var cleaned = cleanTitlePrefix(itemName);
    var kws = extractKeywords(cleaned, excludeMap, synonymMap);
    for (var k = 0; k < kws.length; k++) {
      addWordRecord(wordRecords, kws[k], sourceLabel);
    }
    if (items[i].tagIds && items[i].tagIds.length > 0) {
      for (var tg = 0; tg < items[i].tagIds.length; tg++) {
        tagIdSet[items[i].tagIds[tg]] = true;
      }
    }
  }
}

function addWordRecord(wordRecords, word, source) {
  var w = String(word || '').trim();
  if (!w || w.length < 2) return;
  var key = w + '||' + source;
  if (!wordRecords[key]) {
    wordRecords[key] = { word: w, source: source, hits: 0 };
  }
  wordRecords[key].hits++;
}

/**
 * 楽天 商品検索API でジャンル内商品を取得
 */
function fetchItemSearchByGenre(appId, genreId, sort, hits, page) {
  var url = RAKUTEN_ITEM_SEARCH
    + '?format=json&applicationId=' + appId
    + '&genreId=' + genreId
    + '&sort=' + encodeURIComponent(sort)
    + '&hits=' + (hits || 30)
    + '&page=' + (page || 1)
    + '&availability=1';
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('ItemSearch HTTP ' + response.getResponseCode() + ' genreId=' + genreId);
      return [];
    }
    var parsed = JSON.parse(response.getContentText());
    if (!parsed.Items) return [];
    return parsed.Items.map(function(it) {
      var item = it.Item || it;
      return {
        itemName: item.itemName,
        itemUrl : item.itemUrl,
        tagIds  : item.tagIds || [],
      };
    });
  } catch(e) {
    Logger.log('ItemSearch error genreId=' + genreId + ': ' + e);
    return [];
  }
}

/**
 * 楽天 商品ランキングAPI でジャンル別の最新ランキングを取得
 * period パラメータは数値形式(YYYYMMDD)指定のため省略=最新ランキング
 * page: 1-N (1ページ=30件)
 */
function fetchItemRankingByGenre(appId, genreId, page) {
  var url = RAKUTEN_ITEM_RANKING
    + '?format=json&applicationId=' + appId
    + '&genreId=' + genreId
    + '&hits=30'
    + '&page=' + (page || 1);
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('ItemRanking HTTP ' + response.getResponseCode() + ' genreId=' + genreId + ' page=' + page
                 + ' body=' + response.getContentText().substring(0, 200));
      return [];
    }
    var parsed = JSON.parse(response.getContentText());
    if (!parsed.Items) return [];
    return parsed.Items.map(function(it) {
      var item = it.Item || it;
      return {
        itemName: item.itemName,
        itemUrl : item.itemUrl,
        tagIds  : item.tagIds || [],
      };
    });
  } catch(e) {
    Logger.log('ItemRanking error genreId=' + genreId + ' page=' + page + ': ' + e);
    return [];
  }
}

/**
 * 楽天 ジャンル検索API で指定ジャンルの子ジャンル名一覧を取得
 */
function fetchGenreChildren(appId, genreId) {
  var url = RAKUTEN_GENRE_SEARCH
    + '?format=json&applicationId=' + appId
    + '&genreId=' + genreId;
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return [];
    var parsed = JSON.parse(response.getContentText());
    var names = [];
    var children = parsed.children || [];
    for (var i = 0; i < children.length; i++) {
      var c = children[i].child || children[i];
      if (c.genreName) names.push(String(c.genreName).trim());
    }
    return names;
  } catch(e) {
    Logger.log('GenreSearch error genreId=' + genreId + ': ' + e);
    return [];
  }
}

/**
 * タグID配列 → タグ名Map解決（タグ辞書シートをキャッシュとして利用）
 * キャッシュになければTagSearch APIで取得して辞書に追加
 */
function resolveTagNames(tagIds, appId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.TAG_DICT);
  if (!ws) ws = initTagDictSheet(ss);

  // 既存辞書を読む
  var dict = {};
  if (ws.getLastRow() >= 2) {
    var data = ws.getRange(2, 1, ws.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      var id = String(data[i][0] || '').trim();
      if (id) dict[id] = String(data[i][2] || '').trim();
    }
  }

  var unresolved = tagIds.filter(function(id) { return !dict[id]; });
  if (unresolved.length === 0) {
    var resultMap = {};
    for (var r = 0; r < tagIds.length; r++) resultMap[tagIds[r]] = dict[tagIds[r]];
    return resultMap;
  }

  // 未解決タグをTagSearchで問い合わせ（カンマ区切り複数対応）
  // TagSearch は tagId ごとに1回、最大1000件など仕様不明のため1個ずつ
  // GAS時間制限対策で最大50個に限定
  var fetchLimit = Math.min(unresolved.length, 50);
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var newRows = [];
  for (var u = 0; u < fetchLimit; u++) {
    var tagId = unresolved[u];
    var fetched = fetchTagInfo(appId, tagId);
    if (fetched && fetched.tagName) {
      dict[tagId] = fetched.tagName;
      newRows.push([tagId, fetched.groupName || '', fetched.tagName, todayStr]);
    }
    Utilities.sleep(RAKUTEN_API_DELAY_MS);
  }
  if (newRows.length > 0) {
    ws.getRange(ws.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
    Logger.log('タグ辞書 ' + newRows.length + '件追加');
  }

  var finalMap = {};
  for (var f = 0; f < tagIds.length; f++) finalMap[tagIds[f]] = dict[tagIds[f]];
  return finalMap;
}

/**
 * 楽天 タグ検索API でタグID → タグ名 解決
 */
function fetchTagInfo(appId, tagId) {
  var url = RAKUTEN_TAG_SEARCH
    + '?format=json&applicationId=' + appId
    + '&tagId=' + tagId;
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return null;
    var parsed = JSON.parse(response.getContentText());
    var tagGroups = parsed.tagGroups || [];
    for (var i = 0; i < tagGroups.length; i++) {
      var g = tagGroups[i].tagGroup || tagGroups[i];
      var groupName = g.tagGroupName || '';
      var tags = g.tags || [];
      for (var j = 0; j < tags.length; j++) {
        var t = tags[j].tag || tags[j];
        if (String(t.tagId) === String(tagId)) {
          return { tagName: t.tagName, groupName: groupName };
        }
      }
    }
    return null;
  } catch(e) {
    Logger.log('TagSearch error tagId=' + tagId + ': ' + e);
    return null;
  }
}

/**
 * 語彙プールシートに反映
 * 既存行（同ジャンル+ワード+由来）があれば最終更新日とヒット数更新、なければ追加
 */
function updateWordPoolSheet(genreName, wordRecords, dateStr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.WORD_POOL);
  if (!ws) ws = initWordPoolSheet(ss);

  // 既存ジャンル内のレコードを読み込み
  var existing = {};
  if (ws.getLastRow() >= 2) {
    var data = ws.getRange(2, 1, ws.getLastRow() - 1, 6).getValues();
    for (var i = 0; i < data.length; i++) {
      var g = String(data[i][0] || '').trim();
      var w = String(data[i][1] || '').trim();
      var s = String(data[i][2] || '').trim();
      if (g === genreName && w && s) {
        existing[w + '||' + s] = i + 2;  // 絶対行番号
      }
    }
  }

  var newRows = [];
  var updates = [];  // {row, hitsAdd}
  var keys = Object.keys(wordRecords);
  for (var k = 0; k < keys.length; k++) {
    var rec = wordRecords[keys[k]];
    if (existing[keys[k]]) {
      updates.push({ row: existing[keys[k]], hitsAdd: rec.hits });
    } else {
      newRows.push([genreName, rec.word, rec.source, dateStr, dateStr, rec.hits]);
    }
  }

  if (newRows.length > 0) {
    ws.getRange(ws.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
  }
  for (var u = 0; u < updates.length; u++) {
    var row = updates[u].row;
    var curHits = Number(ws.getRange(row, 6).getValue() || 0);
    ws.getRange(row, 5).setValue(dateStr);       // 最終更新日
    ws.getRange(row, 6).setValue(curHits + updates[u].hitsAdd);  // ヒット数累積
  }
}

/**
 * 語彙プール読み込み（ジャンル別ワードset）
 * @return {Object} { genreName: { word: true } }
 */
function loadWordPoolByGenre() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.WORD_POOL);
  if (!ws || ws.getLastRow() < 2) return {};
  var data = ws.getRange(2, 1, ws.getLastRow() - 1, 2).getValues();
  var pool = {};
  for (var i = 0; i < data.length; i++) {
    var genre = String(data[i][0] || '').trim();
    var word  = String(data[i][1] || '').trim();
    if (!genre || !word) continue;
    if (!pool[genre]) pool[genre] = {};
    pool[genre][word] = true;
  }
  return pool;
}

/**
 * 語彙プールのチェックポイントをリセット（手動実行用）
 */
function resetWordPoolCheckpoint() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(WP_CHECKPOINT_DATE_KEY);
  props.deleteProperty(WP_CHECKPOINT_INDEX_KEY);
  Logger.log('語彙プールのチェックポイントをリセットしました');
}
