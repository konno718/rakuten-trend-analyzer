// SheetsManager.gs - Google Sheets読み書き

function getOrCreateSheet(ss, name, headers) {
  var ws = ss.getSheetByName(name);
  if (!ws) {
    ws = ss.insertSheet(name);
    if (headers) {
      ws.appendRow(headers);
      ws.getRange(1, 1, 1, headers.length)
        .setBackground('#4A86E8')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      ws.setFrozenRows(1);
    }
  }
  return ws;
}

function readGenreConfigs() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!ws) {
    createSettingsTemplate(ss);
    throw new Error('設定シートを作成しました。楽天ランキングURLを入力してから再実行してください。');
  }
  var data = ws.getDataRange().getValues();
  var configs = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || !row[1]) continue;
    var genreName  = String(row[0]).trim();
    var rakutenUrl = String(row[1]).trim();
    var enabledRaw = row[2];
    var enabled = enabledRaw === true || (String(enabledRaw).trim() !== '×' && String(enabledRaw).trim() !== 'false' && String(enabledRaw).trim() !== 'FALSE');
    var keepaCat   = String(row[3] || '').trim();
    if (enabled && genreName && rakutenUrl) {
      configs.push({ genreName: genreName, rakutenUrl: rakutenUrl, keepaCat: keepaCat });
    }
  }
  Logger.log('有効ジャンル数: ' + configs.length);
  return configs;
}

function createSettingsTemplate(ss) {
  var ws = getOrCreateSheet(ss, SHEET_NAMES.SETTINGS, [
    'ジャンル名', '楽天ランキングURL', '有効(○/×)', 'KeepaカテゴリID（任意）', 'メモ'
  ]);
  var samples = [
    ['インテリア・雑貨', 'https://ranking.rakuten.co.jp/daily/100533/', '○', '', 'URLは実際のカテゴリに変更'],
    ['ペット用品',       'https://ranking.rakuten.co.jp/daily/101213/', '○', '', ''],
    ['アウトドア',       'https://ranking.rakuten.co.jp/daily/101070/', '○', '', ''],
    ['美容・健康',       'https://ranking.rakuten.co.jp/daily/100227/', '○', '', ''],
    ['キッチン用品',     'https://ranking.rakuten.co.jp/daily/558885/', '×', '', '不要なら×に'],
  ];
  for (var i = 0; i < samples.length; i++) ws.appendRow(samples[i]);
}

function initExcludeWordsSheet(ss) {
  var ws = getOrCreateSheet(ss, SHEET_NAMES.EXCLUDES, ['ワード', '有効(○/×)', '種類', 'メモ']);
  if (ws.getLastRow() > 1) return;
  var defaults = [
    ['母の日','○','検索タグ',''], ['父の日','○','検索タグ',''], ['クリスマス','○','検索タグ',''],
    ['バレンタイン','○','検索タグ',''], ['誕生日','○','検索タグ',''],
    ['プレゼント','○','検索タグ',''], ['ギフト','○','検索タグ',''],
    ['敬老の日','○','検索タグ',''], ['ハロウィン','○','検索タグ',''],
    ['卒業','○','検索タグ',''], ['入学','○','検索タグ',''],
    ['高品質','○','汎用語',''], ['プレミアム','○','汎用語',''],
    ['おしゃれ','○','汎用語',''], ['かわいい','○','汎用語',''],
    ['シンプル','○','汎用語',''], ['北欧','○','汎用語',''],
    ['人気','○','汎用語',''], ['新品','○','汎用語',''],
    ['送料無料','○','汎用語',''], ['お得','○','汎用語',''],
    ['セール','○','汎用語',''], ['最新','○','汎用語',''],
    ['セット','○','補助語',''], ['タイプ','○','補助語',''],
    ['サイズ','○','補助語',''], ['専用','○','補助語',''],
  ];
  for (var i = 0; i < defaults.length; i++) ws.appendRow(defaults[i]);
}

function writeWordStats(allResults, dateStr, newEntrantMap) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = getOrCreateSheet(ss, SHEET_NAMES.WORDS, [
    '日付', 'ジャンル', 'キーワード', '出現回数', '平均順位', 'スコア', '分類', '新規参入'
  ]);
  var rows = [];
  for (var i = 0; i < allResults.length; i++) {
    var r = allResults[i];
    var classLabel = {
      hidden_gem: SCORE_RULES.HIDDEN_GEM.label,
      trending  : SCORE_RULES.TRENDING.label,
      saturated : SCORE_RULES.SATURATED.label,
    }[r.classification] || r.classification;
    var key = r.genre + '::' + r.keyword;
    var isNew = newEntrantMap && newEntrantMap[key];
    rows.push([dateStr, r.genre, r.keyword, r.count, r.avgRank, r.finalScore, classLabel, isNew ? '🆕新規' : '']);
  }
  if (rows.length > 0) {
    ws.getRange(ws.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    Logger.log('ワード集計 ' + rows.length + '行書き込み');
  }
}

function writeProducts(allResults, dateStr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = getOrCreateSheet(ss, SHEET_NAMES.PRODUCTS, [
    '日付', 'ジャンル', 'キーワード', '分類', '順位', '商品名', '商品URL'
  ]);
  var rows = [];
  for (var i = 0; i < allResults.length; i++) {
    var r = allResults[i];
    if (r.classification === 'saturated') continue;
    var classLabel = { hidden_gem: SCORE_RULES.HIDDEN_GEM.label, trending: SCORE_RULES.TRENDING.label }[r.classification] || '';
    var products = r.products || [];
    for (var j = 0; j < products.length; j++) {
      var p = products[j];
      rows.push([dateStr, r.genre, r.keyword, classLabel, p.rank, p.itemName, p.itemUrl]);
    }
  }
  if (rows.length > 0) {
    ws.getRange(ws.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    Logger.log('商品一覧 ' + rows.length + '行書き込み');
  }
}

function writeExcludeCandidates(candidates, dateStr) {
  if (!candidates || candidates.length === 0) return;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  // 除外候補シートに書き込む（除外ワードシートとは別）
  var ws = getOrCreateSheet(ss, '除外候補', ['有効', 'ワード', '種類', 'メモ']);

  var lastRow = ws.getLastRow();
  var existingRow = {};
  var existingMemo = {};
  if (lastRow >= 2) {
    // B列=ワード, C列=種類, D列=メモ を一括取得
    var data = ws.getRange(2, 2, lastRow - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      var w = String(data[i][0] || '').trim();
      if (!w) continue;
      existingRow[w]  = i + 2;                    // 絶対行番号
      existingMemo[w] = String(data[i][2] || ''); // D列メモ
    }
  }

  var newRows = [];
  var updates = [];
  for (var j = 0; j < candidates.length; j++) {
    var word = candidates[j];
    if (existingRow[word]) {
      updates.push({ row: existingRow[word], memo: buildExcludeMemo(existingMemo[word], dateStr) });
    } else {
      // A列=FALSE(チェックボックス), B列=ワード, C列=種類, D列=メモ
      newRows.push([false, word, '自動検出', '初出:' + dateStr + ' / 最新:' + dateStr + ' / 検出:1回']);
    }
  }

  if (newRows.length > 0) {
    var startRow = ws.getLastRow() + 1;
    ws.getRange(startRow, 1, newRows.length, 4).setValues(newRows);
    ws.getRange(startRow, 1, newRows.length, 1).insertCheckboxes();
    Logger.log('除外候補 ' + newRows.length + '件追加');
  }
  for (var k = 0; k < updates.length; k++) {
    ws.getRange(updates[k].row, 4).setValue(updates[k].memo);
  }
  if (updates.length > 0) {
    Logger.log('除外候補 ' + updates.length + '件メモ更新');
  }
}

/**
 * 既存メモをパースして最新日・検出回数を更新したメモ文字列を返す。
 * 形式: "初出:YYYY-MM-DD / 最新:YYYY-MM-DD / 検出:N回"
 * 旧形式 "YYYY-MM-DD に頻出" からの移行にも対応。
 */
function buildExcludeMemo(existingMemo, todayStr) {
  var firstMatch = existingMemo.match(/初出:(\d{4}-\d{2}-\d{2})/);
  var countMatch = existingMemo.match(/検出:(\d+)回/);
  var oldFormat  = existingMemo.match(/(\d{4}-\d{2}-\d{2})\s*に頻出/);

  var first, count;
  if (firstMatch) {
    first = firstMatch[1];
    count = countMatch ? parseInt(countMatch[1], 10) + 1 : 2;
  } else if (oldFormat) {
    first = oldFormat[1];
    count = 2;
  } else {
    first = todayStr;
    count = 1;
  }
  return '初出:' + first + ' / 最新:' + todayStr + ' / 検出:' + count + '回';
}

function writeSummary(text, dateStr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = getOrCreateSheet(ss, SHEET_NAMES.SUMMARY, ['日付', '分析結果']);
  ws.appendRow([dateStr, text]);
}

function getPastKeywordMap(daysBack) {
  daysBack = daysBack || 30;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.WORDS);
  if (!ws || ws.getLastRow() < 2) return {};
  var data = ws.getDataRange().getValues();
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysBack);
  var pastMap = {};
  for (var i = 1; i < data.length; i++) {
    var rowDate = new Date(data[i][0]);
    if (rowDate >= cutoff) {
      var key = String(data[i][1]) + '::' + String(data[i][2]);
      pastMap[key] = true;
    }
  }
  return pastMap;
}
