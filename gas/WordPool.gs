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
  // モード別有効フラグ導入後は genreId 重複除外しない（同ジャンルでも
  // モードが違えば除外ワードが異なるため別々に収集）
  Logger.log('処理ジャンル×モード数: ' + genreConfigs.length);

  if (index >= genreConfigs.length) {
    Logger.log('語彙プール: 本日分は全ジャンル完了済み（index=' + index + ' / ' + genreConfigs.length + '）');
    return;
  }

  Logger.log('=== 語彙プール更新: index=' + index + ' から ===');

  var deadline = new Date(startTime.getTime() + WORDPOOL_STEP_TIME_LIMIT_MS);

  while (index < genreConfigs.length) {
    if (new Date() >= deadline) {
      Logger.log('時間切れ。次回トリガーで index=' + index + ' から継続');
      break;
    }

    var gc = genreConfigs[index];
    try {
      processGenreForWordPool(gc, config.rakutenAppId, dateStr, deadline);
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
 * 4ソース: ItemSearch / ItemRanking / タグ辞書 / 子ジャンル名
 * deadline を渡すと途中で時間切れ判定して早期リターン（部分結果は書き込み済）
 */
function processGenreForWordPool(gc, appId, dateStr, deadline) {
  var genreId = parseGenreIdFromUrl(gc.rakutenUrl);
  if (!genreId) {
    Logger.log('ジャンルID取得失敗: ' + gc.rakutenUrl);
    return;
  }

  Logger.log('[' + gc.genreName + '] 語彙プール更新開始');

  // Step 2: 語彙プール構築時はフィルタせず全トークン登録
  // 除外・装飾語・消耗品・同義語は推奨ワード分析時に適用
  var excludeMap = {};
  var synonymMap = {};
  var wordRecords = {};
  var tagIdSet = {};
  var timeExceeded = function() { return deadline && new Date() >= deadline; };

  // --- ソース1: ItemSearch 上位 (分類=main: レビュー多い定番商品由来) ---
  for (var p = 1; p <= WORDPOOL_ITEM_PAGES; p++) {
    if (timeExceeded()) { Logger.log('  時間切れ(ItemSearch)'); break; }
    var items = fetchItemSearchByGenre(appId, genreId, '-reviewCount', 30, p);
    if (items.length === 0) break;
    collectWordsFromItems(items, 'タイトル派生', 'main', wordRecords, tagIdSet, excludeMap, synonymMap);
    Utilities.sleep(RAKUTEN_API_DELAY_MS);
  }

  // --- ソース2: ItemRanking 最新 (分類=main: 売れ筋商品由来) ---
  for (var r = 1; r <= WORDPOOL_RANKING_PAGES; r++) {
    if (timeExceeded()) { Logger.log('  時間切れ(ItemRanking)'); break; }
    var rankItems = fetchItemRankingByGenre(appId, genreId, r);
    if (rankItems.length === 0) break;
    collectWordsFromItems(rankItems, 'ランキング派生(' + r + 'page)', 'main', wordRecords, tagIdSet, excludeMap, synonymMap);
    Utilities.sleep(RAKUTEN_API_DELAY_MS);
  }

  // --- ソース3: タグ辞書 (分類=sub: 公式検索タグ) ---
  var tagIds = Object.keys(tagIdSet);
  if (tagIds.length > 0 && !timeExceeded()) {
    var tagNameMap = resolveTagNames(tagIds, appId, deadline);
    for (var t = 0; t < tagIds.length; t++) {
      var name = tagNameMap[tagIds[t]];
      if (!name) continue;
      addWordRecord(wordRecords, name, '検索タグ', 'sub');
    }
  }

  // --- ソース4: 子ジャンル名 (分類=sub: カテゴリ階層) ---
  if (!timeExceeded()) {
    var subGenres = fetchGenreChildren(appId, genreId);
    for (var s = 0; s < subGenres.length; s++) {
      addWordRecord(wordRecords, subGenres[s], 'サブジャンル', 'sub');
    }
    Utilities.sleep(RAKUTEN_API_DELAY_MS);
  }

  // --- 語彙プールシートに反映（部分でも書き込み・モード別有効フラグ更新） ---
  updateWordPoolSheet(gc.genreName, wordRecords, dateStr, gc.mode);

  Logger.log('[' + gc.genreName + ':' + gc.mode + '] 語彙プール完了: ' + Object.keys(wordRecords).length + 'レコード');
}

/**
 * 商品群からキーワード抽出 + タグID収集
 */
function collectWordsFromItems(items, sourceLabel, classification, wordRecords, tagIdSet, excludeMap, synonymMap) {
  for (var i = 0; i < items.length; i++) {
    var itemName = items[i].itemName || '';
    var cleaned = cleanTitlePrefix(itemName);
    var kws = extractKeywords(cleaned, excludeMap, synonymMap);
    for (var k = 0; k < kws.length; k++) {
      addWordRecord(wordRecords, kws[k], sourceLabel, classification);
    }
    if (items[i].tagIds && items[i].tagIds.length > 0) {
      for (var tg = 0; tg < items[i].tagIds.length; tg++) {
        tagIdSet[items[i].tagIds[tg]] = true;
      }
    }
  }
}

function addWordRecord(wordRecords, word, source, classification) {
  var w = String(word || '').trim();
  if (!w || w.length < 2) return;
  var key = w + '||' + source;
  if (!wordRecords[key]) {
    wordRecords[key] = { word: w, source: source, classification: classification || 'main', hits: 0 };
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
 * deadlineを渡すと時間切れで早期break（未解決タグは次回以降のrunで補完）
 */
function resolveTagNames(tagIds, appId, deadline) {
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

  // 1run あたり WORDPOOL_TAG_FETCH_LIMIT 件まで（GAS 6分制限対策）
  var fetchLimit = Math.min(unresolved.length, WORDPOOL_TAG_FETCH_LIMIT);
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var newRows = [];
  for (var u = 0; u < fetchLimit; u++) {
    if (deadline && new Date() >= deadline) {
      Logger.log('  TagSearch 時間切れ (' + u + '/' + fetchLimit + ')');
      break;
    }
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
 * 語彙プールシート(v3)に反映
 * 既存行（同ジャンル+ワード+由来）があれば:
 *   - 該当モードの有効列を TRUE に
 *   - 最終更新日・ヒット数を更新
 * なければ新規行追加（該当モード列のみTRUE、他モード列はFALSE）
 * 列: A=ジャンル B=ワード C=由来 D=分類 E=中国輸入有効 F=国内会社有効 G=初出日 H=最終更新日 I=ヒット数
 */
function updateWordPoolSheet(genreName, wordRecords, dateStr, mode) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.WORD_POOL);
  if (!ws) ws = initWordPoolSheet(ss);

  var modeColIdx = (mode === MODES.DOMESTIC) ? 5 : 4;  // 0-index: E=4, F=5

  // 全データを一度だけ読み込み（1 getValues コール）
  var lastRow = ws.getLastRow();
  var allData = (lastRow >= 2) ? ws.getRange(2, 1, lastRow - 1, 9).getValues() : [];

  // 既存行インデックス構築（同ジャンル内）
  var existing = {};
  for (var i = 0; i < allData.length; i++) {
    var g = String(allData[i][0] || '').trim();
    var w = String(allData[i][1] || '').trim();
    var s = String(allData[i][2] || '').trim();
    if (g === genreName && w && s) existing[w + '||' + s] = i;  // 0-based index in allData
  }

  var hasUpdates = false;
  var newRows = [];
  var keys = Object.keys(wordRecords);
  for (var k = 0; k < keys.length; k++) {
    var rec = wordRecords[keys[k]];
    if (existing[keys[k]] !== undefined) {
      var idx = existing[keys[k]];
      // メモリ内で更新
      allData[idx][modeColIdx] = true;
      allData[idx][7] = dateStr;  // 最終更新日 (H列)
      allData[idx][8] = Number(allData[idx][8] || 0) + rec.hits;  // ヒット数 (I列)
      hasUpdates = true;
    } else {
      var isChina = (mode !== MODES.DOMESTIC);
      newRows.push([
        genreName, rec.word, rec.source, rec.classification || 'main',
        isChina, !isChina, dateStr, dateStr, rec.hits
      ]);
    }
  }

  // 既存行の一括書き戻し（変更があれば全行バルク書き込み）
  if (hasUpdates && allData.length > 0) {
    ws.getRange(2, 1, allData.length, 9).setValues(allData);
  }

  // 新規行追加
  if (newRows.length > 0) {
    var startRow = ws.getLastRow() + 1;
    ws.getRange(startRow, 1, newRows.length, 9).setValues(newRows);
    ws.getRange(startRow, 5, newRows.length, 2).insertCheckboxes();
  }
}

/**
 * 語彙プール読み込み（モード別・ジャンル別ワード→分類map）
 * 該当モードの有効フラグ=TRUEの行のみ返す
 * @param {string} mode - 指定モード (省略時は両モード統合)
 * @return {Object} { genreName: { word: 'main'|'sub' } }
 */
function loadWordPoolByGenre(mode) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.WORD_POOL);
  if (!ws || ws.getLastRow() < 2) return {};
  var data = ws.getRange(2, 1, ws.getLastRow() - 1, 6).getValues();
  var modeColIdx = (mode === MODES.DOMESTIC) ? 5 : 4;
  var pool = {};
  for (var i = 0; i < data.length; i++) {
    var genre = String(data[i][0] || '').trim();
    var word  = String(data[i][1] || '').trim();
    var cls   = String(data[i][3] || 'main').trim();
    if (!genre || !word) continue;
    // モード指定があれば該当列TRUEのみ
    if (mode && data[i][modeColIdx] !== true) continue;
    if (!pool[genre]) pool[genre] = {};
    if (pool[genre][word] !== 'main') pool[genre][word] = cls;
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
