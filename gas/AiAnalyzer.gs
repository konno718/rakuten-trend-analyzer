// AiAnalyzer.gs - Gemini API でワード群をメイン+サブに分類

/**
 * Gemini API を使ってジャンル内のワード群を「メインワード + サブワード」に分類
 * 入力が多い場合は GEMINI_BATCH_SIZE 毎に分割して複数回呼び出す。
 * 各エントリの .word / .subWords を Gemini の判定で上書き。
 */
function analyzeWithGemini(apiKey, genreName, classified) {
  if (!apiKey || classified.length === 0) return classified;

  var total = classified.length;
  for (var start = 0; start < total; start += GEMINI_BATCH_SIZE) {
    var end = Math.min(start + GEMINI_BATCH_SIZE, total);
    var batch = [];
    for (var i = start; i < end; i++) {
      batch.push({
        id       : i,
        word     : classified[i].word,
        subWords : (classified[i].subWords || []).slice(0, 15),
        inPool   : !!classified[i].isMainWord,
      });
    }
    analyzeGeminiBatch(apiKey, genreName, classified, batch);
  }
  return classified;
}

function analyzeGeminiBatch(apiKey, genreName, classified, batch) {
  var prompt = [
    '楽天ランキング「' + genreName + '」ジャンルのワード群を分析してください。',
    '',
    '各エントリについて、商品カテゴリを表す「メインワード」と、商品特徴を表す「サブワード（付加価値語）」を判定してください。',
    '- メインワードは1個、サブワードは最大10個',
    '- inPoolがtrueのものは楽天が認識している語なので、メインワード候補として優先',
    '- サブワードはそのワードに紐づく付加価値（素材・サイズ・用途・形状等）',
    '- 既存のword/subWordsが適切ならそのまま使ってOK',
    '',
    '出力はJSON配列のみ（説明文なし・厳密なJSON）:',
    '[{"id":0,"mainWord":"...","subWords":["...","..."]}, ...]',
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
      // 思考モードを無効化（出力トークンを温存）
      thinkingConfig   : { thinkingBudget: 0 },
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, {
      method      : 'post',
      contentType : 'application/json',
      payload     : JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) {
      Logger.log('Gemini API error (' + response.getResponseCode() + '): ' + response.getContentText().substring(0, 400));
      return;
    }
    var body = JSON.parse(response.getContentText());
    var candidate = body.candidates && body.candidates[0];
    if (!candidate || !candidate.content || !candidate.content.parts) {
      Logger.log('Gemini応答にcontentなし: ' + JSON.stringify(body).substring(0, 300));
      return;
    }
    var text = candidate.content.parts.map(function(p){ return p.text || ''; }).join('');
    var parsed = tryParseJsonArray(text);
    if (!parsed) {
      Logger.log('Gemini JSONパース失敗 [' + genreName + ']: ' + text.substring(0, 300));
      return;
    }
    for (var i = 0; i < parsed.length; i++) {
      var r = parsed[i];
      if (typeof r.id !== 'number' || !classified[r.id]) continue;
      if (r.mainWord && typeof r.mainWord === 'string' && r.mainWord.trim().length > 0) {
        classified[r.id].word = r.mainWord.trim();
      }
      if (Array.isArray(r.subWords)) {
        // 文字列のみ・2文字以上・空白trim した値を最大10個
        classified[r.id].subWords = r.subWords
          .map(function(x) { return (typeof x === 'string') ? x.trim() : ''; })
          .filter(function(x) { return x && x.length >= 2; })
          .slice(0, 10);
      }
    }
    Logger.log('[' + genreName + '] Gemini batch ' + parsed.length + '/' + batch.length);
  } catch(e) {
    Logger.log('Gemini API fetch error: ' + e);
  }
}

/**
 * ワード群を Gemini で分類
 * 中国輸入モードは「液物/食べ物/医薬品/医療機器」も追加分類対象、国内は「医薬品」のみ
 * 入力: genreName, words, mode
 * 出力: [{word, type}] type: 'カテゴリ'|'ブランド'|'装飾語'|'対象'|'食べ物'|'液物'|'医薬品'|'医療機器'
 */
function categorizeWordsWithGemini(apiKey, genreName, words, mode) {
  if (!apiKey || words.length === 0) return [];

  var modeLabel = (mode === MODES.DOMESTIC) ? '国内メーカーモード' : '中国輸入モード';
  var classes = [
    '- カテゴリ: 具体的な商品種別 (例: ゴミ箱, 充電器, ベビーカー, ベッド)',
    '- ブランド: メーカー名・商品ブランド名 (例: 山崎実業, ロイヤルカナン, ニトリ, パンパース)',
    '- 装飾語: 形容詞・属性・特性 (例: おしゃれ, 軽量, 防水, 北欧, シンプル)',
    '- 対象: 利用者・場面・用途・状況 (例: 新生児, 出産祝い, 子供, 女性, ペット用)',
    '- 医薬品: 第1類/第2類医薬品・サプリ・栄養補助食品系 (例: サプリ, ビタミン, プロテイン, 錠剤)',
  ];
  if (mode !== MODES.DOMESTIC) {
    // 中国輸入モードは食べ物・液物・医療機器も別分類
    classes.push('- 食べ物: 食品全般 (例: フード, おやつ, 餌, ペットフード, お菓子)');
    classes.push('- 液物: 液体製品 (例: シャンプー, 洗剤, 化粧水, ジュース)');
    classes.push('- 医療機器: 医療用具 (例: コンタクトレンズ, カラコン, 補聴器)');
  }

  var prompt = [
    '楽天ランキング「' + genreName + '」ジャンル ' + modeLabel + ' から抽出されたワード群を分析してください。',
    '',
    '各ワードを以下の分類のどれか1つに判定してください:',
  ].concat(classes).concat([
    '',
    '判定が難しい場合は「カテゴリ」として返してください。',
    '',
    '出力はJSON配列のみ（説明文なし）:',
    '[{"word":"...","type":"..."}, ...]',
    '',
    '入力:',
    JSON.stringify(words)
  ]).join('\n');

  var url = GEMINI_API_URL_BASE + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey);
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature      : 0.2,
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
      Logger.log('Gemini categorize error ' + response.getResponseCode() + ': ' + response.getContentText().substring(0, 300));
      return [];
    }
    var body = JSON.parse(response.getContentText());
    var cand = body.candidates && body.candidates[0];
    if (!cand || !cand.content || !cand.content.parts) return [];
    var text = cand.content.parts.map(function(p){ return p.text || ''; }).join('');
    var parsed = tryParseJsonArray(text);
    if (!parsed) return [];

    var results = [];
    for (var i = 0; i < parsed.length; i++) {
      var r = parsed[i];
      if (typeof r.word !== 'string' || typeof r.type !== 'string') continue;
      results.push({ word: r.word.trim(), type: r.type.trim() });
    }
    return results;
  } catch(e) {
    Logger.log('Gemini categorize fetch error: ' + e);
    return [];
  }
}

/**
 * 想定どおりのJSON配列ならそのまま、途中で切れている場合は有効な要素だけ抽出
 */
function tryParseJsonArray(text) {
  var match = text.match(/\[[\s\S]*\]/);
  if (match) {
    try { return JSON.parse(match[0]); } catch(e) { /* fallthrough */ }
  }
  // 切れた配列から完全なオブジェクト要素だけを復旧
  var openIdx = text.indexOf('[');
  if (openIdx < 0) return null;
  var items = [];
  var depth = 0;
  var start = -1;
  for (var i = openIdx; i < text.length; i++) {
    var ch = text.charAt(i);
    if (ch === '{') {
      if (depth === 0) start = i;
      depth++;
    } else if (ch === '}') {
      depth--;
      if (depth === 0 && start >= 0) {
        var chunk = text.substring(start, i + 1);
        try {
          items.push(JSON.parse(chunk));
        } catch(e) { /* skip */ }
        start = -1;
      }
    }
  }
  return items.length > 0 ? items : null;
}
