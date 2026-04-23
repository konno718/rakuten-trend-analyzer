// Suggest.gs - 楽天サジェスト + 商品検索API 連携（フェーズ2: 共起分析）

/**
 * 楽天サジェストエンドポイントのレスポンス確認（手動実行用）
 * 本実装前にこれを1回実行し、ログでレスポンス形式を確認する
 */
function testSuggest() {
  var seeds = ['ゴミ箱', 'ベッド'];
  for (var i = 0; i < seeds.length; i++) {
    var seed = seeds[i];
    var url = SUGGEST_URL + encodeURIComponent(seed);
    Logger.log('--- seed: ' + seed + ' ---');
    Logger.log('URL: ' + url);
    try {
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
          'Accept': 'application/json, text/plain, */*',
        }
      });
      Logger.log('HTTP: ' + response.getResponseCode());
      var body = response.getContentText();
      Logger.log('Body先頭800文字: ' + body.substring(0, 800));
      try {
        var parsed = JSON.parse(body);
        Logger.log('Parsed top-level keys: ' + Object.keys(parsed).join(', '));
        Logger.log('Stringify先頭1000文字: ' + JSON.stringify(parsed).substring(0, 1000));
      } catch(e) {
        Logger.log('JSONパース失敗: ' + e);
      }
    } catch(e) {
      Logger.log('fetch error: ' + e);
    }
    Utilities.sleep(SUGGEST_DELAY_MS);
  }
}

/**
 * 楽天サジェストからキーワード候補を取得
 * レスポンス構造は testSuggest() で確認し、各パターンに対応
 */
function fetchSuggest(seed) {
  var url = SUGGEST_URL + encodeURIComponent(seed);
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
        'Accept': 'application/json, text/plain, */*',
      }
    });
    if (response.getResponseCode() !== 200) return [];
    var body = response.getContentText();
    var parsed;
    try { parsed = JSON.parse(body); }
    catch(e) { return []; }

    // よくある形式を順に試す
    if (Array.isArray(parsed)) {
      // ["キーワード", ["候補1", "候補2"...]] の OpenSearch 形式
      if (parsed.length >= 2 && Array.isArray(parsed[1])) return parsed[1];
      return parsed.filter(function(x){ return typeof x === 'string'; });
    }
    if (parsed.s && Array.isArray(parsed.s)) return parsed.s;
    if (parsed.suggest && Array.isArray(parsed.suggest)) return parsed.suggest;
    if (parsed.suggestions && Array.isArray(parsed.suggestions)) return parsed.suggestions;
    if (parsed.keywords && Array.isArray(parsed.keywords)) {
      return parsed.keywords.map(function(k){ return typeof k === 'string' ? k : k.keyword || ''; }).filter(function(x){return x;});
    }
    return [];
  } catch(e) {
    Logger.log('fetchSuggest error for "' + seed + '": ' + e);
    return [];
  }
}

/**
 * 楽天市場商品検索APIでヒット総数（市場需要指標）を取得
 */
function fetchKeywordItemCount(appId, keyword) {
  var url = SEARCH_API_URL
    + '?format=json&keyword=' + encodeURIComponent(keyword)
    + '&applicationId=' + appId
    + '&hits=1&availability=1';
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('SearchAPI HTTP ' + response.getResponseCode() + ' for "' + keyword + '"');
      return 0;
    }
    var parsed = JSON.parse(response.getContentText());
    return parsed.count || 0;
  } catch(e) {
    Logger.log('fetchKeywordItemCount error for "' + keyword + '": ' + e);
    return 0;
  }
}

/**
 * サジェスト文字列から種ワード(seed)を除いた修飾語部分を抽出
 * 例: "ゴミ箱 分別 おしゃれ" + seed="ゴミ箱" → "分別 おしゃれ"
 * 修飾語がなければ空文字列を返す
 */
function extractModifier(suggestStr, seed) {
  var s = String(suggestStr || '').trim();
  if (!s || !seed) return '';
  var escaped = seed.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  var rest = s.replace(new RegExp(escaped, 'g'), '').replace(/\s+/g, ' ').trim();
  return rest;
}
