// Extractor.gs - 商品名からキーワードを抽出する

// 長音符 ー は区切り扱いしない（カタカナ語が分断されるのを防ぐ）
var SEPARATORS_RE = /[\s　・\/\|【】「」『』（）()\[\]［］〔〕＜＞<>★☆◆◇■□●○※◎\-_＿～〜,，、。！!？?#＃@＠&＆\+＋]/g;

/**
 * 商品名先頭の販促プレフィックス（括弧ブロック・プロモトークン）を剥がす
 * 例: 【LINE500円OFFクーポン】[期間限定]10%OFFクーポン ＼1位獲得／ 本物の商品名
 *   → 本物の商品名
 * 半角[] / 全角［］ / 【】 / 〔〕 / （） / () / 「」 / 『』 / ＼...／ に対応
 */
function cleanTitlePrefix(name) {
  var s = String(name);
  var bracketRe    = /^[\s　]*(?:【[^】]*】|\[[^\]]*\]|［[^］]*］|＼[^／]*／|〔[^〕]*〕|\([^)]*\)|（[^）]*）|「[^」]*」|『[^』]*』)/;
  var promoTokenRe = /^[\s　]*[^\s　【\[［＼〔(（「『]*(?:クーポン|OFF|ポイント倍|獲得|レビュー特典|保証[!！]*|SALE|半額|当店最安)[^\s　【\[［＼〔(（「『]*/i;
  var prev = null;
  while (s !== prev) {
    prev = s;
    s = s.replace(bracketRe, '');
    s = s.replace(promoTokenRe, '');
  }
  return s.replace(/^[\s　]+/, '');
}

// 両モード共通除外（第1類・第2類医薬品）
var CATEGORY_EXCLUDES_COMMON = [
  '第1類医薬品', '第2類医薬品', '指定第2類医薬品',
  '第１類医薬品', '第２類医薬品', '指定第２類医薬品',
  '第1類', '第2類', '指定第2類',
  '第１類', '第２類', '指定第２類',
];

// 中国輸入モードで追加除外（食べ物・液物・サプリ類・医療機器）
var CATEGORY_EXCLUDES_CHINA = [
  // 食べ物
  '食品','飲料','お菓子','スナック','チョコ','ケーキ','パン','コーヒー',
  'お茶','ジュース','ビール','ワイン','日本酒','ラーメン','パスタ','調味料',
  'フード','おやつ','餌','エサ',
  // 液物
  'シャンプー','トリートメント','コンディショナー','洗剤','柔軟剤',
  '化粧水','乳液','香水','美容液',
  // サプリ・プロテイン・目薬（中国輸入は飲用NG）
  'サプリ','サプリメント','ビタミン','プロテイン','錠剤','カプセル','目薬',
  // 医療機器
  'コンタクト','コンタクトレンズ','カラコン',
];

// 国内メーカーモードで追加除外（なし。COMMON のみ）
var CATEGORY_EXCLUDES_DOMESTIC = [];

/**
 * 同義語シートから { 同義語 → 正規ワード } のマップを読む
 * シート構成: A=正規ワード, B=同義語（カンマ/読点区切り）
 */
function loadSynonymMap() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.SYNONYMS);
  if (!ws || ws.getLastRow() < 2) return {};
  var data = ws.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var canonical = String(data[i][0] || '').trim();
    var synStr    = String(data[i][1] || '').trim();
    if (!canonical || !synStr) continue;
    var syns = synStr.split(/[,，、]/);
    for (var j = 0; j < syns.length; j++) {
      var s = syns[j].trim();
      if (s) map[s] = canonical;
    }
  }
  return map;
}

/**
 * 除外ワードシートからモードに応じた有効ワードを読む
 * シート構成: A=中国輸入有効, B=国内会社有効, C=ワード, D=種類, E=メモ
 */
function loadExcludeWords(mode) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.EXCLUDES);
  if (!ws) return {};
  var data = ws.getDataRange().getValues();
  var modeColIdx = (mode === MODES.DOMESTIC) ? 1 : 0;  // 0-indexed
  var wordColIdx = 2;
  var excludeMap = {};
  for (var i = 1; i < data.length; i++) {
    var enabled = data[i][modeColIdx];
    var word    = String(data[i][wordColIdx] || '').trim();
    if (word && enabled === true) excludeMap[word] = true;
  }
  return excludeMap;
}

function isProductExcluded(itemName, mode) {
  for (var i = 0; i < CATEGORY_EXCLUDES_COMMON.length; i++) {
    if (itemName.indexOf(CATEGORY_EXCLUDES_COMMON[i]) >= 0) return true;
  }
  var modeList = (mode === MODES.DOMESTIC) ? CATEGORY_EXCLUDES_DOMESTIC : CATEGORY_EXCLUDES_CHINA;
  for (var j = 0; j < modeList.length; j++) {
    if (itemName.indexOf(modeList[j]) >= 0) return true;
  }
  return false;
}

function extractKeywords(itemName, excludeMap, synonymMap) {
  var name = itemName.substring(0, 100);
  var parts = name.split(SEPARATORS_RE);
  var tokens = [];
  for (var i = 0; i < parts.length; i++) {
    var t = parts[i].trim();
    if (t.length >= 2) {
      // 同義語は抽出段階で正規ワードに置換
      if (synonymMap && synonymMap[t]) t = synonymMap[t];
      tokens.push(t);
    }
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

  // 複合語（隣接する2トークンを結合・正規化済み）
  for (var k = 0; k < tokens.length - 1; k++) {
    var compound = tokens[k] + tokens[k + 1];
    if (compound.length >= 4 && compound.length <= 15) {
      if (synonymMap && synonymMap[compound]) compound = synonymMap[compound];
      if (!seen[compound] && !(excludeMap && excludeMap[compound])) {
        seen[compound] = true;
        keywords.push(compound);
      }
    }
  }

  return keywords;
}

function processItems(items, excludeMap, synonymMap, mode) {
  return items.map(function(item) {
    var cleanedName = cleanTitlePrefix(item.itemName);
    var excluded = isProductExcluded(cleanedName, mode);
    return {
      rank      : item.rank,
      itemCode  : item.itemCode,
      itemName  : cleanedName,
      itemNameOriginal : item.itemName,
      itemUrl   : item.itemUrl,
      itemPrice : item.itemPrice,
      genreId   : item.genreId,
      excluded  : excluded,
      keywords  : excluded ? [] : extractKeywords(cleanedName, excludeMap, synonymMap),
    };
  });
}
