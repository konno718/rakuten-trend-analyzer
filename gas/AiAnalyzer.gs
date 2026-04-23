// AiAnalyzer.gs - Gemini API でワード群をメイン+サブに分類

/**
 * Gemini API を使ってジャンル内のワード群を「メインワード + サブワード」に分類
 * 入力: ジャンル内のワード配列（.word, .subWords, .isMainWord, .count 等）
 * 出力: 同じ配列構造だが .word と .subWords が Gemini の判定で上書きされる
 * APIキー未設定なら元の配列をそのまま返す
 */
function analyzeWithGemini(apiKey, genreName, classified) {
  if (!apiKey || classified.length === 0) return classified;

  var items = classified.map(function(x, i) {
    return {
      id       : i,
      word     : x.word,
      subWords : (x.subWords || []).slice(0, 15),
      inPool   : !!x.isMainWord,
      count    : x.count || 0,
    };
  });

  var prompt = [
    '楽天ランキング「' + genreName + '」ジャンルのワード群を分析してください。',
    '',
    '各エントリについて、商品カテゴリを表す「メインワード」と、商品特徴を表す「サブワード（付加価値語）」を判定してください。',
    '- メインワードは1個、サブワードは最大10個',
    '- inPoolがtrueのものは楽天が認識している語なので、メインワード候補として優先',
    '- サブワードはそのワードに紐づく付加価値（素材・サイズ・用途・形状等）',
    '- 既存のword/subWordsが適切ならそのまま使ってOK',
    '- 判定不能なら元のwordをそのままmainWordに',
    '',
    '出力はJSON配列のみ（説明文なし・厳密なJSON）:',
    '[{"id":0,"mainWord":"...","subWords":["...","..."]}, ...]',
    '',
    '入力:',
    JSON.stringify(items)
  ].join('\n');

  var url = GEMINI_API_URL_BASE + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey);
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature       : 0.3,
      maxOutputTokens   : GEMINI_MAX_TOKENS,
      responseMimeType  : 'application/json',
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
      Logger.log('Gemini API error (' + response.getResponseCode() + '): ' + response.getContentText().substring(0, 500));
      return classified;
    }
    var body = JSON.parse(response.getContentText());
    var candidate = body.candidates && body.candidates[0];
    if (!candidate || !candidate.content || !candidate.content.parts) {
      Logger.log('Gemini応答にcontentが無い: ' + JSON.stringify(body).substring(0, 300));
      return classified;
    }
    var text = candidate.content.parts.map(function(p){ return p.text || ''; }).join('');
    var jsonMatch = text.match(/\[[\s\S]*\]/);
    if (!jsonMatch) {
      Logger.log('Gemini応答にJSONが見つからず: ' + text.substring(0, 300));
      return classified;
    }
    var parsed;
    try { parsed = JSON.parse(jsonMatch[0]); }
    catch(e) {
      Logger.log('Gemini JSONパース失敗: ' + e + ' / ' + jsonMatch[0].substring(0, 300));
      return classified;
    }
    for (var i = 0; i < parsed.length; i++) {
      var r = parsed[i];
      if (typeof r.id !== 'number' || !classified[r.id]) continue;
      if (r.mainWord) classified[r.id].word = r.mainWord;
      if (Array.isArray(r.subWords)) classified[r.id].subWords = r.subWords.slice(0, 10);
    }
    Logger.log('[' + genreName + '] Gemini分析完了 ' + parsed.length + '件');
  } catch(e) {
    Logger.log('Gemini API fetch error: ' + e);
  }

  return classified;
}
