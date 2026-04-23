// HiddenGem.gs - Step 4-7: 商品単位ランキング × 語彙プール突合 → サブワード集計 → お宝分析

/**
 * お宝分析メイン処理
 * Step 4: 商品毎にキーワード抽出 → 語彙プール突合
 * Step 5: メインワード判定（pool main分類に該当する語、複数ならpool hits最大を採用）
 * Step 6: お宝候補（pool未登録）はAIでメインワード推定
 * Step 7: サブワード集計（同一メインワード群での共起頻度上位）
 */
function runHiddenGemAnalysis() {
  var startTime = new Date();
  var dateStr = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy-MM-dd');
  var currentMonth = startTime.getMonth() + 1;
  Logger.log('=== お宝分析開始: ' + dateStr + ' ===');

  var config = getConfig();
  var disposables = loadDisposableWords();
  var wordPool    = loadWordPoolByGenre();   // {genre: {word: 'main'|'sub'}}
  var poolHits    = loadWordPoolHitsByGenre();// {genre: {word: totalHits}}
  var surveyed    = loadSurveyedWords();

  Logger.log('消耗品 ' + Object.keys(disposables).length
             + ' / 語彙プール ' + Object.keys(wordPool).length + 'ジャンル'
             + ' / 調査済み ' + Object.keys(surveyed).length);

  var genreConfigs = readGenreConfigs();
  // ジャンルID重複を除外
  var seen = {};
  var uniqueConfigs = [];
  for (var u = 0; u < genreConfigs.length; u++) {
    var gid = parseGenreIdFromUrl(genreConfigs[u].rakutenUrl);
    if (gid && !seen[gid]) { seen[gid] = true; uniqueConfigs.push(genreConfigs[u]); }
  }
  genreConfigs = uniqueConfigs;

  var deadline = new Date(startTime.getTime() + 5 * 60 * 1000);
  var allRows = [];

  for (var gi = 0; gi < genreConfigs.length; gi++) {
    if (new Date() >= deadline) {
      Logger.log('時間切れ。未処理ジャンル ' + (genreConfigs.length - gi) + ' は次回');
      break;
    }
    var gc = genreConfigs[gi];
    var products = readRankingProducts(gc.mode, gc.genreName, dateStr);
    if (products.length === 0) {
      Logger.log('[' + gc.genreName + '] ランキング生データなし');
      continue;
    }

    var rows = analyzeGenre(gc, products, wordPool, poolHits, disposables, surveyed, config, currentMonth);
    allRows = allRows.concat(rows);
    Logger.log('[' + gc.genreName + '] 出力 ' + rows.length + '行');
  }

  if (allRows.length === 0) {
    Logger.log('書き込み対象なし');
    return;
  }

  // NEW判定（過去14日間のお宝分析履歴と照合）
  var past = loadHiddenGemHistory(HIDDEN_GEM_NEW_DAYS);
  for (var r = 0; r < allRows.length; r++) {
    var key = allRows[r].genre + '::' + allRows[r].word;
    var p = past[key];
    if (!p) {
      allRows[r].newStatus = 'NEW!';
    } else {
      var pSet = {};
      (p.subWords || []).forEach(function(s){ pSet[s] = true; });
      var hasNewSub = false;
      (allRows[r].subWords || []).forEach(function(s){ if (!pSet[s]) hasNewSub = true; });
      allRows[r].newStatus = hasNewSub ? '🔄更新' : '既存';
    }
  }

  writeHiddenGemResults(allRows, dateStr);
  var elapsed = (new Date() - startTime) / 1000;
  Logger.log('=== お宝分析完了: ' + allRows.length + '行 / ' + elapsed.toFixed(1) + '秒 ===');
}

/**
 * 1ジャンルの商品群を分析してお宝分析行を生成
 */
function analyzeGenre(gc, products, wordPool, poolHits, disposables, surveyed, config, currentMonth) {
  var excludeMap = loadExcludeWords(gc.mode);
  var synonymMap = loadSynonymMap();
  var genrePool = wordPool[gc.genreName] || {};
  var genreHits = poolHits[gc.genreName] || {};

  // 各商品を分析
  var analyses = [];  // [{product, primaryMain, subs, unknowns, allWords}]
  for (var i = 0; i < products.length; i++) {
    var p = products[i];
    var kws = extractKeywords(p.itemName, excludeMap, synonymMap);

    var mains = [];
    var subs = [];
    var unknowns = [];
    for (var k = 0; k < kws.length; k++) {
      var w = kws[k];
      if (disposables[w]) continue;  // 消耗品ワードは落とす
      var cls = genrePool[w];
      if (cls === 'main') mains.push(w);
      else if (cls === 'sub') subs.push(w);
      else unknowns.push(w);
    }

    // メインワード採用: pool hits が多い順、無ければnull（お宝候補）
    var primaryMain = null;
    if (mains.length > 0) {
      mains.sort(function(a, b) { return (genreHits[b] || 0) - (genreHits[a] || 0); });
      primaryMain = mains[0];
    }

    analyses.push({
      product     : p,
      primaryMain : primaryMain,
      otherMains  : mains.slice(1),  // サブワード扱いで回収
      subs        : subs,
      unknowns    : unknowns,
    });
  }

  // お宝候補（primaryMain 未決）だけ AI に送ってメインワード推定
  var treasures = [];
  for (var t = 0; t < analyses.length; t++) {
    if (!analyses[t].primaryMain) treasures.push(analyses[t]);
  }
  if (config.geminiApiKey && treasures.length > 0) {
    assignAiMainWords(treasures, config.geminiApiKey, gc.genreName);
  }

  // グループ集計（メインワードごと・サブ候補は整数カウントで蓄積）
  var groups = {};  // mainWord → {products:[], subFreq:{}, isTreasure:bool}
  for (var a = 0; a < analyses.length; a++) {
    var ana = analyses[a];
    var main = ana.primaryMain || ana.aiMain || '';
    if (!main) continue;
    if (!groups[main]) {
      groups[main] = { products: [], subFreq: {}, isTreasure: !ana.primaryMain };
    }
    groups[main].products.push(ana.product);
    // サブ候補: pool sub + 他のmain + unknowns 全部1回として集計
    var all = ana.subs.concat(ana.otherMains, ana.unknowns);
    for (var s = 0; s < all.length; s++) {
      var sw = all[s];
      if (sw === main) continue;
      groups[main].subFreq[sw] = (groups[main].subFreq[sw] || 0) + 1;
    }
  }

  // グループを行オブジェクトに変換
  var rows = [];
  var mainWords = Object.keys(groups);
  for (var m = 0; m < mainWords.length; m++) {
    var mw = mainWords[m];
    var g = groups[mw];
    // 調査済み永久非表示は除外
    var surv = surveyed[gc.genreName + '::' + mw];
    if (surv && surv.status === '永久非表示') continue;

    // サブワードTop10 (4段階優先: pool sub & 複数回 → pool sub & 1回 → その他 & 複数回 → その他 & 1回)
    var subs = pickTopSubWords(g.subFreq, genrePool, 10);

    // 背景色
    var bg = null;
    if (surv && surv.status === '毎年X月再表示') {
      bg = (surv.month === currentMonth) ? HIDDEN_GEM_COLOR_RECALL : HIDDEN_GEM_COLOR_SURVEYED;
    }

    // 上位商品URL（rank昇順）
    var topProducts = g.products.slice().sort(function(a, b) { return a.rank - b.rank; });

    rows.push({
      genre       : gc.genreName,
      word        : mw,
      subWords    : subs,
      count       : g.products.length,
      type        : g.isTreasure ? 'お宝候補' : 'メインワード',
      evaluation  : SCORE_RULES.HIDDEN_GEM.label,
      products    : topProducts.slice(0, HIDDEN_GEM_URL_COUNT),
      backgroundColor: bg,
    });
  }

  // ソート: 出現回数desc、同数ならお宝優先
  rows.sort(function(a, b) {
    if (b.count !== a.count) return b.count - a.count;
    if (a.type === 'お宝候補' && b.type !== 'お宝候補') return -1;
    if (b.type === 'お宝候補' && a.type !== 'お宝候補') return 1;
    return 0;
  });
  return rows.slice(0, HIDDEN_GEM_MAX_PER_GENRE);
}

/**
 * サイズ・数量・重量・寸法・期間などの測定値を表すワードか判定
 * 「3kg」「100ml」「30枚」「2L」「30x40cm」「10週間」等は除外したい
 */
function isMeasurementWord(word) {
  if (!word) return false;
  var w = String(word).trim();
  if (/^\d+(\.\d+)?$/.test(w)) return true;                                     // 数字のみ
  if (/^\d+(\.\d+)?(ml|mL|cc|l|L|ℓ|g|kg|mg|μg|oz|mm|cm|m|寸|枚|個|本|袋|箱|パック|缶|錠|粒|回|日|月|年|週|時間|分|秒|度|℃|畳|人用|台|人|坪)$/i.test(w)) return true;
  if (/^(XXS|XS|S|M|L|LL|XL|XXL|XXXL|SS)$/.test(w)) return true;                // 衣料サイズ
  if (/^\d+L$/.test(w)) return true;                                            // 2L 3L 4L
  if (/^\d+[x×]\d+/.test(w)) return true;                                       // 30x40, 100×200
  if (/^\d+(週間|ヶ月|か月|月齢|歳|才|年生)$/.test(w)) return true;              // 期間・年齢
  if (/^\d+(個入|本入|枚入|個セット|本セット|枚セット|組|入)$/.test(w)) return true; // セット数
  return false;
}

/**
 * サブワードTop N を選定
 *   - 除外: サイズ・数量・重量等の測定ワード
 *   - 採用: freq>=1 全て（新規ワードも積極採用）
 *   - ソート: 頻度desc → 同値なら 語彙プール sub を優先
 */
function pickTopSubWords(subFreq, genrePool, maxN) {
  var candidates = [];
  var words = Object.keys(subFreq);
  for (var i = 0; i < words.length; i++) {
    var w = words[i];
    if (isMeasurementWord(w)) continue;  // サイズ/数量/重量/期間除外
    candidates.push({
      word      : w,
      freq      : subFreq[w],
      isPoolSub : genrePool[w] === 'sub',
    });
  }
  candidates.sort(function(a, b) {
    if (b.freq !== a.freq) return b.freq - a.freq;              // 頻度desc
    return (b.isPoolSub ? 1 : 0) - (a.isPoolSub ? 1 : 0);       // 同頻度なら pool sub 優先
  });
  return candidates.slice(0, maxN).map(function(c) { return c.word; });
}

/**
 * お宝候補（メインワード未決）の商品タイトルをGeminiに渡してメインワード推定
 * 推定結果は treasures[i].aiMain にセット
 */
function assignAiMainWords(treasures, apiKey, genreName) {
  if (treasures.length === 0) return;
  for (var start = 0; start < treasures.length; start += GEMINI_BATCH_SIZE) {
    var end = Math.min(start + GEMINI_BATCH_SIZE, treasures.length);
    var batch = [];
    for (var i = start; i < end; i++) {
      batch.push({ id: i, title: treasures[i].product.itemName });
    }
    aiBatchMainWord(apiKey, genreName, treasures, batch);
  }
}

function aiBatchMainWord(apiKey, genreName, treasures, batch) {
  var prompt = [
    '楽天ランキング「' + genreName + '」の商品タイトル群から、各商品のメインワード（商品カテゴリ名）を1つずつ抽出してください。',
    '',
    '- メインワードは商品カテゴリ・商品名を表す1〜5文字程度の日本語',
    '- ブランド名や型番は避け、汎用的な商品カテゴリ名に',
    '- タイトルから明確な商品カテゴリが読み取れない場合は空文字',
    '',
    '出力はJSON配列のみ（説明文なし）:',
    '[{"id":0,"mainWord":"..."}, ...]',
    '',
    '入力:',
    JSON.stringify(batch)
  ].join('\n');

  var url = GEMINI_API_URL_BASE + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey);
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature      : 0.3,
      maxOutputTokens  : GEMINI_MAX_TOKENS,
      responseMimeType : 'application/json',
      thinkingConfig   : { thinkingBudget: 0 },
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) {
      Logger.log('Gemini error ' + response.getResponseCode() + ': ' + response.getContentText().substring(0, 300));
      return;
    }
    var body = JSON.parse(response.getContentText());
    var cand = body.candidates && body.candidates[0];
    if (!cand || !cand.content || !cand.content.parts) return;
    var text = cand.content.parts.map(function(p){ return p.text || ''; }).join('');
    var parsed = tryParseJsonArray(text);
    if (!parsed) {
      Logger.log('Gemini JSON失敗 [' + genreName + ']: ' + text.substring(0, 200));
      return;
    }
    for (var i = 0; i < parsed.length; i++) {
      var r = parsed[i];
      if (typeof r.id !== 'number' || !treasures[r.id]) continue;
      if (typeof r.mainWord === 'string' && r.mainWord.trim().length > 0) {
        treasures[r.id].aiMain = r.mainWord.trim();
      }
    }
    Logger.log('[' + genreName + '] AI main判定 ' + parsed.length + '件');
  } catch(e) {
    Logger.log('Gemini fetch error: ' + e);
  }
}

/**
 * 語彙プールのジャンル別ヒット数Map
 */
function loadWordPoolHitsByGenre() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.WORD_POOL);
  if (!ws || ws.getLastRow() < 2) return {};
  var data = ws.getRange(2, 1, ws.getLastRow() - 1, 7).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var genre = String(data[i][0] || '').trim();
    var word  = String(data[i][1] || '').trim();
    var hits  = Number(data[i][6] || 0);
    if (!genre || !word) continue;
    if (!map[genre]) map[genre] = {};
    map[genre][word] = (map[genre][word] || 0) + hits;
  }
  return map;
}

/**
 * 過去 daysBack 日以内のお宝分析履歴 (ジャンル::メインワード → {subWords})
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
 * お宝分析シートに行追加（蓄積式・背景色セット）
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
      row.push(r.products && r.products[p] ? stripQueryFromUrl(r.products[p].itemUrl) : '');
    }
    values.push(row);

    var rowBg = [];
    for (var b = 0; b < totalCols; b++) rowBg.push(r.backgroundColor || null);
    bgColors.push(rowBg);
  }

  var startRow = ws.getLastRow() + 1;
  ws.getRange(startRow, 1, values.length, totalCols).setValues(values);
  ws.getRange(startRow, 1, bgColors.length, totalCols).setBackgrounds(bgColors);
  Logger.log('お宝分析シートに ' + values.length + '行書き込み');
}
