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
  Logger.log('=== 推奨ワード分析開始: ' + dateStr + ' ===');

  var config = getConfig();
  var disposables = loadDisposableWords();
  // モード別語彙プール（該当モードの有効フラグTRUEの行のみ）
  var wordPool = {};
  wordPool[MODES.CHINA]    = loadWordPoolByGenre(MODES.CHINA);
  wordPool[MODES.DOMESTIC] = loadWordPoolByGenre(MODES.DOMESTIC);
  var poolHits = {};
  poolHits[MODES.CHINA]    = loadWordPoolHitsByGenre(MODES.CHINA);
  poolHits[MODES.DOMESTIC] = loadWordPoolHitsByGenre(MODES.DOMESTIC);
  var surveyed = loadSurveyedWords();

  Logger.log('消耗品 ' + Object.keys(disposables).length
             + ' / 語彙プール 中国輸入:' + Object.keys(wordPool[MODES.CHINA]).length + 'ジャンル'
             + ' 国内:' + Object.keys(wordPool[MODES.DOMESTIC]).length + 'ジャンル'
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
  var rowsByMode = {};
  rowsByMode[MODES.CHINA]    = [];
  rowsByMode[MODES.DOMESTIC] = [];

  // 除外候補検出用のモード別アキュムレータ
  //   keywordGenres: {キーワード: 出現ジャンル数}
  //   genreCount: そのモードで処理したジャンル数
  var modeAccum = {};
  modeAccum[MODES.CHINA]    = { keywordGenres: {}, genreCount: 0 };
  modeAccum[MODES.DOMESTIC] = { keywordGenres: {}, genreCount: 0 };

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

    // 分析開始順位で絞り込み（例: 51位以降を分析対象にする等）
    if (gc.startRank && gc.startRank > 1) {
      var before = products.length;
      products = products.filter(function(p) { return p.rank >= gc.startRank; });
      Logger.log('[' + gc.genreName + '] 分析開始順位=' + gc.startRank + ' で絞り込み: ' + before + ' → ' + products.length);
    }

    modeAccum[gc.mode].genreCount++;
    var rows = analyzeGenre(gc, products, wordPool[gc.mode] || {}, poolHits[gc.mode] || {}, disposables, surveyed, config, currentMonth, modeAccum[gc.mode].keywordGenres);
    rowsByMode[gc.mode] = rowsByMode[gc.mode].concat(rows);
    Logger.log('[' + gc.genreName + '] 出力 ' + rows.length + '行');
  }

  // 除外候補自動検出（全ジャンル処理後、モード別に 30% 閾値で判定）
  [MODES.CHINA, MODES.DOMESTIC].forEach(function(mode) {
    var accum = modeAccum[mode];
    if (accum.genreCount < 2) return;
    var threshold = Math.max(2, Math.ceil(accum.genreCount * 0.3));
    var candidates = [];
    Object.keys(accum.keywordGenres).forEach(function(kw) {
      if (accum.keywordGenres[kw] >= threshold) candidates.push(kw);
    });
    if (candidates.length > 0) {
      writeExcludeCandidates(candidates, dateStr, mode);
      Logger.log('[' + mode + '] 除外候補 ' + candidates.length
                 + '件検出 (閾値' + threshold + '/' + accum.genreCount + 'ジャンル)');
    }
  });

  var totalWritten = 0;
  [MODES.CHINA, MODES.DOMESTIC].forEach(function(mode) {
    var rows = rowsByMode[mode];
    if (rows.length === 0) return;

    // NEW判定（モード別の推奨ワードシート過去14日履歴）
    var past = loadHiddenGemHistory(HIDDEN_GEM_NEW_DAYS, mode);
    for (var r = 0; r < rows.length; r++) {
      var key = rows[r].genre + '::' + rows[r].word;
      var p = past[key];
      if (!p) {
        rows[r].newStatus = 'NEW!';
        rows[r].isNew = true;
      } else {
        var pSet = {};
        (p.subWords || []).forEach(function(s){ pSet[s] = true; });
        var hasNewSub = false;
        (rows[r].subWords || []).forEach(function(s){ if (!pSet[s]) hasNewSub = true; });
        rows[r].newStatus = hasNewSub ? '🔄更新' : '既存';
        rows[r].isNew = false;
      }
    }

    // 月次シート参照で4区分判定
    //  定番ワード  = 月次シートに登場履歴あり（継続的に記録されてる語）
    //  新着ワード  = 月次未登場 & 14日以内の推奨ワード履歴になし
    //  レアワード  = 月次未登場 & 出現数=1 & ランキング100位以内
    //  トレンドワード = 上記どれにも該当しない（日次のみに出てるメイン）
    var monthly = loadMonthlyWordSet(mode);
    var filtered = [];
    for (var k = 0; k < rows.length; k++) {
      var row = rows[k];
      var type = null;
      if (monthly[row.genre + '::' + row.word]) {
        type = '定番ワード';
      } else if (row.isNew) {
        type = '新着ワード';
      } else if (row.count === 1 && row.topProductRank <= HIDDEN_GEM_RARE_RANK_MAX) {
        type = 'レアワード';
      } else {
        type = 'トレンドワード';
      }
      row.type = type;
      filtered.push(row);
    }

    // 最終ソート: 新着 > トレンド > レア > 定番、各区分内で出現回数desc
    filtered.sort(function(a, b) {
      var order = { '新着ワード': 0, 'トレンドワード': 1, 'レアワード': 2, '定番ワード': 3 };
      var oa = order[a.type] || 9;
      var ob = order[b.type] || 9;
      if (oa !== ob) return oa - ob;
      if (b.count !== a.count) return b.count - a.count;
      return a.topProductRank - b.topProductRank;
    });

    writeHiddenGemResults(filtered, dateStr, mode);
    totalWritten += filtered.length;
  });

  var elapsed = (new Date() - startTime) / 1000;
  Logger.log('=== 推奨ワード分析完了: 総' + totalWritten + '行 / ' + elapsed.toFixed(1) + '秒 ===');
}

/**
 * 1ジャンルの商品群を分析して推奨ワード行を生成
 * keywordGenreAccum を渡すとそのジャンル内の出現キーワード（ユニーク）を +1 して返す
 * （除外候補の横断検出用）
 */
function analyzeGenre(gc, products, wordPool, poolHits, disposables, surveyed, config, currentMonth, keywordGenreAccum) {
  var excludeMap    = loadExcludeWords(gc.mode);
  var decorativeMap = loadDecorativeWords(gc.mode);
  var synonymMap    = {};  // Step 2: 同義語は一旦使わない
  var genrePool = wordPool[gc.genreName] || {};
  var genreHits = poolHits[gc.genreName] || {};

  // Step 2: ランキング生データは全商品保存されているので、ここで商品レベル除外を適用
  var beforeCount = products.length;
  products = products.filter(function(p) { return !isProductExcluded(p.itemName, gc.mode); });
  if (products.length !== beforeCount) {
    Logger.log('[' + gc.genreName + '] 商品除外: ' + beforeCount + ' → ' + products.length);
  }

  var genreUniqueKws = {};  // ジャンル内でユニーク出現した語（除外候補検出用）

  // 各商品を分析
  var analyses = [];
  for (var i = 0; i < products.length; i++) {
    var p = products[i];
    var kws = extractKeywords(p.itemName, excludeMap, synonymMap);

    var mains = [];
    var subs = [];
    var unknowns = [];
    for (var k = 0; k < kws.length; k++) {
      var w = kws[k];
      if (disposables[w]) continue;  // 消耗品ワードは落とす
      genreUniqueKws[w] = true;
      // 装飾語はメイン候補にせず、強制サブ扱い
      if (decorativeMap[w]) {
        subs.push(w);
        continue;
      }
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

    // type はNEW判定後に最終決定するので、ここでは isTreasure + topRank 情報を付けて保持
    var topRank = (topProducts.length > 0) ? topProducts[0].rank : 999;
    rows.push({
      genre          : gc.genreName,
      word           : mw,
      subWords       : subs,
      count          : g.products.length,
      isTreasure     : g.isTreasure,
      topProductRank : topRank,
      evaluation     : SCORE_RULES.HIDDEN_GEM.label,
      products       : topProducts.slice(0, HIDDEN_GEM_URL_COUNT),
      backgroundColor: bg,
    });
  }

  // ソート: 出現回数desc、treasure優先（最終順位は区分判定後に再ソート）
  rows.sort(function(a, b) {
    if (b.count !== a.count) return b.count - a.count;
    if (a.isTreasure && !b.isTreasure) return -1;
    if (!a.isTreasure && b.isTreasure) return 1;
    return 0;
  });

  // 呼び出し元の除外候補アキュムレータにジャンル内ユニーク語を加算
  if (keywordGenreAccum) {
    Object.keys(genreUniqueKws).forEach(function(kw) {
      keywordGenreAccum[kw] = (keywordGenreAccum[kw] || 0) + 1;
    });
  }

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
 * 語彙プール(v3)のモード別・ジャンル別ヒット数Map
 * 該当モードの有効フラグ=TRUE行のみ集計
 * @param {string} mode - 指定モード (省略時は全行)
 */
function loadWordPoolHitsByGenre(mode) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.WORD_POOL);
  if (!ws || ws.getLastRow() < 2) return {};
  var data = ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues();
  var modeColIdx = (mode === MODES.DOMESTIC) ? 5 : 4;
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var genre = String(data[i][0] || '').trim();
    var word  = String(data[i][1] || '').trim();
    var hits  = Number(data[i][8] || 0);
    if (!genre || !word) continue;
    if (mode && data[i][modeColIdx] !== true) continue;
    if (!map[genre]) map[genre] = {};
    map[genre][word] = (map[genre][word] || 0) + hits;
  }
  return map;
}

/**
 * 語彙プール月次シートからモード別の (ジャンル::ワード) 登場セットを返す
 * 定番ワード判定用
 */
function loadMonthlyWordSet(mode) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.WORD_POOL_MONTHLY);
  if (!ws || ws.getLastRow() < 2) return {};
  var data = ws.getRange(2, 1, ws.getLastRow() - 1, 3).getValues();
  var set = {};
  for (var i = 0; i < data.length; i++) {
    var genre = String(data[i][0] || '').trim();
    var m     = String(data[i][1] || '').trim();
    var word  = String(data[i][2] || '').trim();
    if (!genre || !word || m !== mode) continue;
    set[genre + '::' + word] = true;
  }
  return set;
}

/**
 * 過去 daysBack 日以内のモード別推奨ワード履歴 (ジャンル::メインワード → {subWords})
 * NEW/更新判定用
 */
function loadHiddenGemHistory(daysBack, mode) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(getSuggestSheetName(mode));
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
 * 推奨ワードのメインワードを Gemini で分類して、
 * カテゴリ以外（装飾語/ブランド/対象）は除外候補シートに自動追加
 * 既に除外ワード/除外候補登録済みの語は skip
 * 両モードを巡回
 */
function validateMainWordsWithGemini() {
  var config = getConfig();
  if (!config.geminiApiKey) {
    Logger.log('GEMINI_API_KEY 未設定。スキップ');
    return;
  }
  _validateWordsByColumn(config.geminiApiKey, 'main');
}

/**
 * 推奨ワードのサブワードを Gemini で分類
 * メイン処理が終わってから手動実行する想定
 */
function validateSubWordsWithGemini() {
  var config = getConfig();
  if (!config.geminiApiKey) {
    Logger.log('GEMINI_API_KEY 未設定。スキップ');
    return;
  }
  _validateWordsByColumn(config.geminiApiKey, 'sub');
}

/**
 * 内部実装: target='main' なら E列(メインワード)、'sub' なら F列(サブワード) を巡回
 */
function _validateWordsByColumn(apiKey, target) {
  var startTime = new Date();
  var dateStr = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy-MM-dd');
  Logger.log('=== Gemini分類開始: ' + (target === 'main' ? 'メインワード' : 'サブワード') + ' / ' + dateStr + ' ===');

  // モード別の skip リスト（該当モードで有効=TRUE のワード）
  // + AI判定で過去に登録済みワード（両モードで再生成しないように）
  var aiAlreadyGenerated = _loadAiAlreadyGenerated();
  var knownByMode = {};
  knownByMode[MODES.CHINA]    = _mergeMaps(_loadKnownExcludeWordsByMode(MODES.CHINA), aiAlreadyGenerated);
  knownByMode[MODES.DOMESTIC] = _mergeMaps(_loadKnownExcludeWordsByMode(MODES.DOMESTIC), aiAlreadyGenerated);
  Logger.log('既知ワード 中国輸入:' + Object.keys(knownByMode[MODES.CHINA]).length
             + ' 国内:' + Object.keys(knownByMode[MODES.DOMESTIC]).length);

  [MODES.CHINA, MODES.DOMESTIC].forEach(function(mode) {
    _processModeForValidation(apiKey, mode, target, knownByMode, dateStr);
  });
}

function _mergeMaps(a, b) {
  var out = {};
  Object.keys(a).forEach(function(k) { out[k] = true; });
  Object.keys(b).forEach(function(k) { out[k] = true; });
  return out;

  var elapsed = (new Date() - startTime) / 1000;
  Logger.log('=== Gemini分類完了 / ' + elapsed.toFixed(1) + '秒 ===');
}

function _processModeForValidation(apiKey, mode, target, knownByMode, dateStr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(getSuggestSheetName(mode));
  if (!ws || ws.getLastRow() < 2) return;

  var alreadyKnown = knownByMode[mode];

  var data = ws.getRange(2, 1, ws.getLastRow() - 1, 6).getValues();
  var byGenre = {};
  for (var i = 0; i < data.length; i++) {
    var genre = String(data[i][3] || '').trim();
    if (!genre) continue;
    var words = [];
    if (target === 'main') {
      var mw = String(data[i][4] || '').trim();
      if (mw) words.push(mw);
    } else {
      var subStr = String(data[i][5] || '').trim();
      if (subStr) words = subStr.split(/[,，、]/).map(function(s){return s.trim();}).filter(function(s){return s;});
    }
    if (!byGenre[genre]) byGenre[genre] = {};
    for (var w = 0; w < words.length; w++) {
      var word = words[w];
      if (!word || alreadyKnown[word]) continue;
      byGenre[genre][word] = true;
    }
  }

  var totalCategorized = 0;
  var totalRegistered = 0;
  Object.keys(byGenre).forEach(function(genre) {
    var words = Object.keys(byGenre[genre]);
    if (words.length === 0) return;

    var BATCH = 50;
    for (var start = 0; start < words.length; start += BATCH) {
      var batch = words.slice(start, Math.min(start + BATCH, words.length));
      var results = categorizeWordsWithGemini(apiKey, genre, batch, mode);
      totalCategorized += results.length;

      var toRegister = [];
      for (var r = 0; r < results.length; r++) {
        var item = results[r];
        var enrolled = _resolveTypeAndFlags(item.type, mode);
        if (!enrolled) continue;  // カテゴリ → スキップ
        toRegister.push({
          word         : item.word,
          type         : enrolled.type,
          chinaEnabled : enrolled.chinaEnabled,
          domesticEnabled: enrolled.domesticEnabled,
        });
      }
      if (toRegister.length > 0) {
        _appendExcludeCandidates(toRegister, dateStr, mode);
        totalRegistered += toRegister.length;
        for (var t = 0; t < toRegister.length; t++) {
          alreadyKnown[toRegister[t].word] = true;
          // 他モードの skip リストにも追加（重複登録防止）
          knownByMode[mode === MODES.CHINA ? MODES.DOMESTIC : MODES.CHINA][toRegister[t].word] = true;
        }
      }
      Utilities.sleep(500);
    }
  });

  Logger.log('[' + mode + '] Gemini判定 ' + totalCategorized + '件 / 除外候補登録 ' + totalRegistered + '件');
}

/**
 * Gemini分類結果 → 除外候補シートの 種類 + モード別フラグ にマッピング
 * - カテゴリ → null (登録しない)
 * - ブランド/装飾語/対象 → 両モード除外候補（chinaEnabled=true, domesticEnabled=true）
 * - 食べ物/液物/医療機器 → 中国輸入のみ除外（domesticは継続）
 * - 医薬品 → 両モード除外
 */
function _resolveTypeAndFlags(geminiType, mode) {
  switch (geminiType) {
    case 'カテゴリ':
      return null;
    case 'ブランド':
      return { type: 'ブランド', chinaEnabled: true, domesticEnabled: true };
    case '装飾語':
    case '対象':
      return { type: '装飾語', chinaEnabled: true, domesticEnabled: true };
    case '医薬品':
      return { type: '装飾語', chinaEnabled: true, domesticEnabled: true };
    case '食べ物':
    case '液物':
    case '医療機器':
      return { type: '装飾語', chinaEnabled: true, domesticEnabled: false };
    default:
      return null;
  }
}

/**
 * 除外ワード + 除外候補から、該当モードで「有効TRUE登録済み」のワード set を返す
 */
function _loadKnownExcludeWordsByMode(mode) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var modeColIdx = (mode === MODES.DOMESTIC) ? 1 : 0;
  var set = {};
  var ex = ss.getSheetByName(SHEET_NAMES.EXCLUDES);
  if (ex && ex.getLastRow() >= 2) {
    var d = ex.getRange(2, 1, ex.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < d.length; i++) {
      if (d[i][modeColIdx] !== true) continue;
      var w = String(d[i][2] || '').trim();
      if (w) set[w] = true;
    }
  }
  var ca = ss.getSheetByName(SHEET_NAMES.CANDIDATES);
  if (ca && ca.getLastRow() >= 2) {
    var d2 = ca.getRange(2, 1, ca.getLastRow() - 1, 5).getValues();
    for (var j = 0; j < d2.length; j++) {
      // 候補は有効FALSEでもAI判定済みなら skip 対象（重複生成防止）
      // ここでは「該当モードで TRUE」のみ skip
      if (d2[j][modeColIdx] !== true) continue;
      var w2 = String(d2[j][2] || '').trim();
      if (w2) set[w2] = true;
    }
  }
  return set;
}

/**
 * AI判定で生成された候補のワードを、両モード共通の「以前生成済み」セットに記録
 */
function _loadAiAlreadyGenerated() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ca = ss.getSheetByName(SHEET_NAMES.CANDIDATES);
  if (!ca || ca.getLastRow() < 2) return {};
  var d = ca.getRange(2, 1, ca.getLastRow() - 1, 5).getValues();
  var set = {};
  for (var i = 0; i < d.length; i++) {
    var memo = String(d[i][4] || '');
    if (memo.indexOf('AI判定') < 0) continue;
    var w = String(d[i][2] || '').trim();
    if (w) set[w] = true;
  }
  return set;
}

/**
 * 除外候補シートに新規行を追加 (Gemini 判定結果用)
 * items: [{word, type, chinaEnabled, domesticEnabled}]
 */
function _appendExcludeCandidates(items, dateStr, mode) {
  if (items.length === 0) return;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.CANDIDATES);
  if (!ws) return;
  var newRows = [];
  for (var i = 0; i < items.length; i++) {
    newRows.push([
      items[i].chinaEnabled === true,
      items[i].domesticEnabled === true,
      items[i].word,
      items[i].type,
      'AI判定:' + mode + ' ' + dateStr,
    ]);
  }
  var startRow = ws.getLastRow() + 1;
  ws.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
  ws.getRange(startRow, 1, newRows.length, 2).insertCheckboxes();
}

/**
 * モード別推奨ワードシートに書き込み
 * 同日の既存行があれば削除してから新規行を追加（同日再実行で結果が最新だけ残る）
 * 過去日分は残す（NEW判定用履歴として使う）
 */
function writeHiddenGemResults(rows, dateStr, mode) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheetName = getSuggestSheetName(mode);
  var ws = ss.getSheetByName(sheetName);
  if (!ws) ws = initSuggestSheet(ss, mode);

  var totalCols = 8 + HIDDEN_GEM_URL_COUNT;

  // 1. 既存データを全読み込み
  var lastRow = ws.getLastRow();
  var keepData = [];
  var keepBg = [];
  if (lastRow >= 2) {
    var allData = ws.getRange(2, 1, lastRow - 1, totalCols).getValues();
    var allBg   = ws.getRange(2, 1, lastRow - 1, totalCols).getBackgrounds();
    for (var i = 0; i < allData.length; i++) {
      var d = allData[i][0];
      if (!d) continue;
      var rowDateStr = (d instanceof Date)
        ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd')
        : String(d).substring(0, 10);
      if (rowDateStr === dateStr) continue;  // 同日分は除外（上書き対象）
      keepData.push(allData[i]);
      keepBg.push(allBg[i]);
    }
  }

  // 2. 新規書き込み行を構築
  var newValues = [];
  var newBg = [];
  for (var r = 0; r < rows.length; r++) {
    var item = rows[r];
    var subStr = (item.subWords || []).slice(0, 10).join(', ');
    var row = [
      dateStr,
      item.count || 0,
      item.type || 'お宝候補',
      item.genre || '',
      item.word || '',
      subStr,
      item.evaluation || SCORE_RULES.HIDDEN_GEM.label,
      item.newStatus || '既存',
    ];
    for (var p = 0; p < HIDDEN_GEM_URL_COUNT; p++) {
      row.push(item.products && item.products[p] ? stripQueryFromUrl(item.products[p].itemUrl) : '');
    }
    newValues.push(row);

    var rowBg = [];
    for (var b = 0; b < totalCols; b++) rowBg.push(item.backgroundColor || null);
    newBg.push(rowBg);
  }

  // 3. シートをヘッダ以外クリア
  if (lastRow > 1) {
    ws.getRange(2, 1, lastRow - 1, totalCols).clearContent();
    var emptyBg = [];
    for (var eb = 0; eb < lastRow - 1; eb++) {
      var rowEmpty = [];
      for (var ec = 0; ec < totalCols; ec++) rowEmpty.push(null);
      emptyBg.push(rowEmpty);
    }
    if (emptyBg.length > 0) ws.getRange(2, 1, emptyBg.length, totalCols).setBackgrounds(emptyBg);
  }

  // 4. 過去日履歴 + 新規 を連結して書き込み
  var merged = keepData.concat(newValues);
  var mergedBg = keepBg.concat(newBg);
  if (merged.length > 0) {
    ws.getRange(2, 1, merged.length, totalCols).setValues(merged);
    ws.getRange(2, 1, mergedBg.length, totalCols).setBackgrounds(mergedBg);
  }

  Logger.log('[' + mode + '] 推奨ワード更新: 過去日' + keepData.length + '行+当日' + newValues.length + '行');
}
