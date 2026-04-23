// AiAnalyzer.gs - Claude API 呼び出しでワード群をメイン+サブに分類

/**
 * Claude API を使ってジャンル内のワード群を「メインワード + サブワード」に分類
 * 入力: ジャンル内のワード配列（.word, .subWords, .isMainWord, .count 等）
 * 出力: 同じ配列構造だが .word と .subWords が Claude の判定で上書きされる
 * APIキー未設定なら元の配列をそのまま返す
 */
function analyzeWithClaude(apiKey, genreName, classified) {
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

  var payload = {
    model     : CLAUDE_MODEL,
    max_tokens: CLAUDE_MAX_TOKENS,
    messages  : [{ role: 'user', content: prompt }],
  };

  try {
    var response = UrlFetchApp.fetch(CLAUDE_API_URL, {
      method      : 'post',
      contentType : 'application/json',
      headers     : {
        'x-api-key'         : apiKey,
        'anthropic-version' : CLAUDE_API_VERSION,
      },
      payload     : JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) {
      Logger.log('Claude API error (' + response.getResponseCode() + '): ' + response.getContentText().substring(0, 500));
      return classified;
    }
    var body = JSON.parse(response.getContentText());
    var content = (body.content && body.content[0] && body.content[0].text) || '';
    var jsonMatch = content.match(/\[[\s\S]*\]/);
    if (!jsonMatch) {
      Logger.log('Claude応答にJSONが見つからず: ' + content.substring(0, 300));
      return classified;
    }
    var parsed;
    try { parsed = JSON.parse(jsonMatch[0]); }
    catch(e) {
      Logger.log('Claude JSONパース失敗: ' + e + ' / ' + jsonMatch[0].substring(0, 300));
      return classified;
    }
    // 結果を反映
    for (var i = 0; i < parsed.length; i++) {
      var r = parsed[i];
      if (typeof r.id !== 'number' || !classified[r.id]) continue;
      if (r.mainWord) classified[r.id].word = r.mainWord;
      if (Array.isArray(r.subWords)) classified[r.id].subWords = r.subWords.slice(0, 10);
    }
    Logger.log('[' + genreName + '] Claude分析完了 ' + parsed.length + '件');
  } catch(e) {
    Logger.log('Claude API fetch error: ' + e);
  }

  return classified;
}
