// Extractor.gs - 商品名からキーワードを抽出する

var SEPARATORS_RE = /[\s\u3000\u30FB\/\|【】「」『』（）()\[\]\u3014\u3015＜＞<>★☆◆◇■□●○※◎\-_＿ー～〜,，、。！!？?#＃@＠&＆\+＋]/g;

var CATEGORY_EXCLUDES = [
  '食品','飲料','お菓子','スナック','チョコ','ケーキ','パン','コーヒー',
  'お茶','ジュース','ビール','ワイン','日本酒','ラーメン','パスタ','調味料',
  '医薬品','サプリ','サプリメント','ビタミン','プロテイン','錠剤','カプセル','目薬',
  'コンタクト','コンタクトレンズ','カラコン',
  'ティッシュ','トイレットペーパー','キッチンペーパー','ラップ','アルミホイル',
  'シャンプー','トリートメント','コンディショナー','洗剤','柔軟剤',
];

function loadExcludeWords() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName(SHEET_NAMES.EXCLUDES);
  if (!ws) return {};
  const data = ws.getDataRange().getValues();
  const excludeMap = {};
  for (let i = 1; i < data.length; i++) {
    const word    = String(data[i][0] || '').trim();
    const enabled = String(data[i][1] || '').trim();
    if (word && enabled !== '×') excludeMap[word] = true;
  }
  return excludeMap;
}

function isProductExcluded(itemName) {
  for (let i = 0; i < CATEGORY_EXCLUDES.length; i++) {
    if (itemName.indexOf(CATEGORY_EXCLUDES[i]) >= 0) return true;
  }
  return false;
}

function extractKeywords(itemName, excludeMap) {
  var name = itemName.substring(0, 100);
  var parts = name.split(SEPARATORS_RE);
  var tokens = [];
  for (var i = 0; i < parts.length; i++) {
    var t = parts[i].trim();
    if (t.length >= 2) tokens.push(t);
  }

  var keywords = [];
  var seen = {};

  for (var j = 0; j < tokens.length; j++) {
    var tok = tokens[j];
    if (/^\d+$/.test(tok)) continue;
    if (/^[a-zA-Z0-9\-]+$/.test(tok)) continue;
    if (excludeMap && excludeMap[tok]) continue;
    if (!seen[tok]) {
      seen[tok] = true;
      keywords.push(tok);
    }
  }

  // 複合語（隣接する2トークンを結合）
  for (var k = 0; k < tokens.length - 1; k++) {
    var compound = tokens[k] + tokens[k + 1];
    if (compound.length >= 4 && compound.length <= 15) {
      if (!seen[compound] && !(excludeMap && excludeMap[compound])) {
        seen[compound] = true;
        keywords.push(compound);
      }
    }
  }

  return keywords;
}

function processItems(items, excludeMap) {
  return items.map(function(item) {
    var excluded = isProductExcluded(item.itemName);
    return {
      rank      : item.rank,
      itemCode  : item.itemCode,
      itemName  : item.itemName,
      itemUrl   : item.itemUrl,
      itemPrice : item.itemPrice,
      genreId   : item.genreId,
      excluded  : excluded,
      keywords  : excluded ? [] : extractKeywords(item.itemName, excludeMap),
    };
  });
}
