// SheetsManager.gs - Google Sheets読み書き

function getOrCreateSheet(ss, name, headers) {
  var ws = ss.getSheetByName(name);
  if (!ws) {
    ws = ss.insertSheet(name);
    if (headers) {
      ws.getRange(1, 1, 1, headers.length).setValues([headers])
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
    var modeRaw    = String(row[4] || '').trim();
    var mode = (modeRaw === MODES.DOMESTIC) ? MODES.DOMESTIC : MODES.CHINA;  // デフォルト 中国輸入
    if (enabled && genreName && rakutenUrl) {
      configs.push({
        genreName  : genreName,
        rakutenUrl : rakutenUrl,
        keepaCat   : keepaCat,
        mode       : mode,
      });
    }
  }
  Logger.log('有効ジャンル数: ' + configs.length);
  return configs;
}

function createSettingsTemplate(ss) {
  var ws = getOrCreateSheet(ss, SHEET_NAMES.SETTINGS, [
    'ジャンル名', '楽天ランキングURL', '有効', 'KeepaカテゴリID（任意）', 'モード', 'メモ'
  ]);
  var samples = [
    ['インテリア・寝具・収納',     'https://ranking.rakuten.co.jp/daily/100804/', true, '', '中国輸入',   ''],
    ['ペット・ペットグッズ',       'https://ranking.rakuten.co.jp/daily/101213/', true, '', '中国輸入',   ''],
    ['スポーツ・アウトドア',       'https://ranking.rakuten.co.jp/daily/101070/', true, '', '中国輸入',   ''],
    ['日用品雑貨・文房具・手芸',   'https://ranking.rakuten.co.jp/daily/215783/', true, '', '中国輸入',   ''],
    ['キッチン用品・食器・調理器具', 'https://ranking.rakuten.co.jp/daily/558944/', true, '', '中国輸入',   ''],
    ['キッズ・ベビー・マタニティ', 'https://ranking.rakuten.co.jp/daily/100533/', true, '', '国内メーカー', ''],
  ];
  if (samples.length > 0) {
    ws.getRange(2, 1, samples.length, samples[0].length).setValues(samples);
    ws.getRange(2, 3, samples.length, 1).insertCheckboxes();
  }
  // モード列のデータ入力規則
  var modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['中国輸入', '国内メーカー'], true)
    .setAllowInvalid(false).build();
  ws.getRange(2, 5, 1000, 1).setDataValidation(modeRule);
}

/**
 * 除外ワードシート初期化（5列構造）
 * A=中国輸入有効, B=国内会社有効, C=ワード, D=種類, E=メモ
 */
function initExcludeWordsSheet(ss) {
  var ws = getOrCreateSheet(ss, SHEET_NAMES.EXCLUDES, ['中国輸入\n有効', '国内会社\n有効', 'ワード', '種類', 'メモ']);
  if (ws.getLastRow() > 1) return;
  var defaults = [
    [true, true, '母の日',       '季節タグ', ''],
    [true, true, '父の日',       '季節タグ', ''],
    [true, true, 'クリスマス',   '季節タグ', ''],
    [true, true, 'バレンタイン', '季節タグ', ''],
    [true, true, '誕生日',       '季節タグ', ''],
    [true, true, 'プレゼント',   '!マーケ', ''],
    [true, true, 'ギフト',       '!マーケ', ''],
    [true, true, '敬老の日',     '季節タグ', ''],
    [true, true, 'ハロウィン',   '季節タグ', ''],
    [true, true, '卒業',         '季節タグ', ''],
    [true, true, '入学',         '季節タグ', ''],
    [true, true, '高品質',       '!品質',   ''],
    [true, true, 'プレミアム',   '!品質',   ''],
    [true, true, 'おしゃれ',     '汎用語',  ''],
    [true, true, 'かわいい',     '汎用語',  ''],
    [true, true, 'シンプル',     '汎用語',  ''],
    [true, true, '北欧',         '汎用語',  ''],
    [true, true, '人気',         '汎用語',  ''],
    [true, true, '新品',         '汎用語',  ''],
    [true, true, '送料無料',     '!マーケ', ''],
    [true, true, 'お得',         '!マーケ', ''],
    [true, true, 'セール',       '!マーケ', ''],
    [true, true, '最新',         '汎用語',  ''],
    [true, true, 'セット',       '補助語',  ''],
    [true, true, 'タイプ',       '補助語',  ''],
    [true, true, 'サイズ',       '補助語',  ''],
    [true, true, '専用',         '補助語',  ''],
  ];
  ws.getRange(2, 1, defaults.length, 5).setValues(defaults);
  ws.getRange(2, 1, defaults.length, 2).insertCheckboxes();
}

/**
 * 除外候補シートが旧4列構造なら5列構造に移行
 * 旧: A(有効), B(ワード), C(種類), D(メモ)
 * 新: A(中国輸入有効), B(国内会社有効), C(ワード), D(種類), E(メモ)
 * 旧A列の値を新A列・新B列の両方にコピー（後で手動で調整する運用）
 */
function migrateExcludeCandidatesTo5Col() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.CANDIDATES);
  if (!ws) {
    Logger.log('除外候補シートが存在しません。init後に実行してください。');
    return;
  }
  var lastCol = ws.getLastColumn();
  var header  = ws.getRange(1, 1, 1, Math.max(1, lastCol)).getValues()[0];
  if (String(header[0]).indexOf('中国輸入') >= 0) {
    Logger.log('除外候補シートは既に新構造です。移行スキップ');
    return;
  }
  var lastRow = ws.getLastRow();
  var oldData = (lastRow >= 2) ? ws.getRange(2, 1, lastRow - 1, 4).getValues() : [];

  ws.clear();
  // 旧C列(種類)に残っているデータ入力規則を全解除
  ws.getRange(1, 1, ws.getMaxRows(), ws.getMaxColumns()).clearDataValidations();

  ws.getRange(1, 1, 1, 5).setValues([['中国輸入\n有効', '国内会社\n有効', 'ワード', '種類', 'メモ']])
    .setBackground('#4A86E8').setFontColor('#FFFFFF').setFontWeight('bold');
  ws.setFrozenRows(1);

  if (oldData.length > 0) {
    var newData = [];
    for (var i = 0; i < oldData.length; i++) {
      var enabled = oldData[i][0] === true;
      newData.push([enabled, enabled, oldData[i][1], oldData[i][2], oldData[i][3]]);
    }
    ws.getRange(2, 1, newData.length, 5).setValues(newData);
    ws.getRange(2, 1, newData.length, 2).insertCheckboxes();
    Logger.log('除外候補 ' + newData.length + '行を5列構造に移行しました');
  } else {
    Logger.log('除外候補シートを5列構造に初期化しました（データなし）');
  }

  // 新 D列(種類) に再度データ入力規則を設定（ブランド・自動検出も含む）
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['!検索タグ', '!マーケ', '!品質', 'デザイン', '機能', '季節タグ', '状態', '汎用語', '補助語', '自動検出', 'ブランド'], true)
    .setAllowInvalid(true).build();
  ws.getRange(2, 4, Math.max(1, ws.getMaxRows() - 1), 1).setDataValidation(typeRule);
}

/**
 * URLからクエリ文字列を除去（?rafcid=... 等）
 */
function stripQueryFromUrl(url) {
  if (!url) return '';
  var s = String(url);
  var idx = s.indexOf('?');
  return idx >= 0 ? s.substring(0, idx) : s;
}

/**
 * ランキングシート（商品単位・フィルタ後）の取得・初期化
 * 列: 日付 | ジャンル | 順位 | 商品名 | URL | タグID
 */
function getOrCreateRankingSheet(ss, sheetName) {
  return getOrCreateSheet(ss, sheetName, ['日付', 'ジャンル', '順位', '商品名', 'URL', 'タグID']);
}

/**
 * モード別ランキングシートに書き込む（商品単位）
 * items: [{rank, itemName, itemUrl, tagIds}]
 */
function writeRankingItems(items, dateStr, genreName, mode) {
  if (!items || items.length === 0) return;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = getOrCreateRankingSheet(ss, getRankingSheetName(mode));
  var rows = items.map(function(item) {
    return [
      dateStr,
      genreName,
      item.rank,
      String(item.itemName || ''),
      stripQueryFromUrl(item.itemUrl),
      (item.tagIds || []).join('|'),
    ];
  });
  ws.getRange(ws.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  Logger.log('[' + mode + ':' + genreName + '] ランキング ' + rows.length + '件書き込み');
}

/**
 * 指定モード+ジャンル+日付 のランキング商品リストを読み込む
 */
function readRankingProducts(mode, genreName, dateStr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(getRankingSheetName(mode));
  if (!ws || ws.getLastRow() < 2) return [];
  var data = ws.getDataRange().getValues();
  var products = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var rowDate = (row[0] instanceof Date)
      ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(row[0]).substring(0, 10);
    if (rowDate !== dateStr) continue;
    if (String(row[1]).trim() !== genreName) continue;
    var tagIdStr = String(row[5] || '');
    var tagIds = tagIdStr ? tagIdStr.split('|').filter(function(x){return x;}) : [];
    products.push({
      rank     : Number(row[2] || 0),
      itemName : String(row[3] || ''),
      itemUrl  : String(row[4] || ''),
      tagIds   : tagIds,
    });
  }
  return products;
}

/**
 * 同義語シートを初期化（空枠のみ）
 * A=正規ワード, B=同義語（カンマ/読点区切り）, C=メモ
 */
function initSynonymSheet(ss) {
  var ws = getOrCreateSheet(ss, SHEET_NAMES.SYNONYMS, ['正規ワード', '同義語（カンマ区切り）', 'メモ']);
  if (ws.getLastRow() > 1) return;
  var samples = [
    ['ゴミ箱', 'ダストワゴン,ダストボックス,ごみ箱', ''],
    ['ベッド', 'ベット', ''],
  ];
  ws.getRange(2, 1, samples.length, 3).setValues(samples);
}

/**
 * 除外候補シートに書き込む（5列構造・モード別フラグ）
 * 新規行: 該当モードのチェックは false で書き、ユーザーがレビューして有効化する運用
 * 既存行: D列メモの初出/最新/検出回数を更新
 */
function writeExcludeCandidates(candidates, dateStr, mode) {
  if (!candidates || candidates.length === 0) return;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = getOrCreateSheet(ss, SHEET_NAMES.CANDIDATES, ['中国輸入\n有効', '国内会社\n有効', 'ワード', '種類', 'メモ']);

  var lastRow = ws.getLastRow();
  var existingRow = {};
  var existingMemo = {};
  if (lastRow >= 2) {
    // C=ワード, D=種類, E=メモ
    var data = ws.getRange(2, 3, lastRow - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      var w = String(data[i][0] || '').trim();
      if (!w) continue;
      existingRow[w]  = i + 2;
      existingMemo[w] = String(data[i][2] || '');
    }
  }

  var newRows = [];
  var updates = [];
  for (var j = 0; j < candidates.length; j++) {
    var word = candidates[j];
    if (existingRow[word]) {
      updates.push({ row: existingRow[word], memo: buildExcludeMemo(existingMemo[word], dateStr, mode) });
    } else {
      // 新規 → 両モードFALSE（ユーザーが手動でチェック）、メモに検出モードを記録
      newRows.push([false, false, word, '自動検出', '初出:' + dateStr + ' / 最新:' + dateStr + ' / 検出:1回 [' + mode + ']']);
    }
  }

  if (newRows.length > 0) {
    var startRow = ws.getLastRow() + 1;
    ws.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
    ws.getRange(startRow, 1, newRows.length, 2).insertCheckboxes();
    Logger.log('[' + mode + '] 除外候補 ' + newRows.length + '件追加');
  }
  for (var k = 0; k < updates.length; k++) {
    ws.getRange(updates[k].row, 5).setValue(updates[k].memo);
  }
  if (updates.length > 0) {
    Logger.log('[' + mode + '] 除外候補 ' + updates.length + '件メモ更新');
  }
}

/**
 * 既存メモをパースして最新日・検出回数を更新したメモ文字列を返す。
 * 形式: "初出:YYYY-MM-DD / 最新:YYYY-MM-DD / 検出:N回 [モード,モード]"
 * 旧形式 "YYYY-MM-DD に頻出" からの移行にも対応。
 */
function buildExcludeMemo(existingMemo, todayStr, mode) {
  var firstMatch = existingMemo.match(/初出:(\d{4}-\d{2}-\d{2})/);
  var countMatch = existingMemo.match(/検出:(\d+)回/);
  var modesMatch = existingMemo.match(/\[([^\]]+)\]/);
  var oldFormat  = existingMemo.match(/(\d{4}-\d{2}-\d{2})\s*に頻出/);

  var first, count, modes;
  if (firstMatch) {
    first = firstMatch[1];
    count = countMatch ? parseInt(countMatch[1], 10) + 1 : 2;
    modes = modesMatch ? modesMatch[1].split(',').map(function(x){return x.trim();}) : [];
  } else if (oldFormat) {
    first = oldFormat[1];
    count = 2;
    modes = [];
  } else {
    first = todayStr;
    count = 1;
    modes = [];
  }
  if (mode && modes.indexOf(mode) < 0) modes.push(mode);
  var modeStr = modes.length > 0 ? ' [' + modes.join(',') + ']' : '';
  return '初出:' + first + ' / 最新:' + todayStr + ' / 検出:' + count + '回' + modeStr;
}

function writeSummary(text, dateStr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = getOrCreateSheet(ss, SHEET_NAMES.SUMMARY, ['日付', '分析結果']);
  ws.appendRow([dateStr, text]);
}

// === お宝分析システム関連シート ===

/**
 * 調査済みワードシート初期化
 * 列: メインワード, ジャンル, 最終調査日, ステータス, 再表示月(1-12), メモ
 */
function initSurveyedSheet(ss) {
  var ws = getOrCreateSheet(ss, SHEET_NAMES.SURVEYED, [
    'メインワード', 'ジャンル', '最終調査日', 'ステータス', '再表示月(1-12)', 'メモ'
  ]);
  // ステータス列のデータ入力規則
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['永久非表示', '毎年X月再表示'], true)
    .setAllowInvalid(false).build();
  ws.getRange(2, 4, Math.max(1, ws.getMaxRows() - 1), 1).setDataValidation(statusRule);
}

/**
 * 語彙プールシート初期化
 * 列: ジャンル, ワード, 由来, 分類, 初出日, 最終更新日, ヒット数
 *   分類: main (ItemSearch/ItemRanking由来=メインワード候補) / sub (TagSearch/GenreSearch由来=サブワード候補)
 */
function initWordPoolSheet(ss) {
  return getOrCreateSheet(ss, SHEET_NAMES.WORD_POOL, [
    'ジャンル', 'ワード', '由来', '分類', '初出日', '最終更新日', 'ヒット数'
  ]);
}

/**
 * 既存の語彙プール (6列構造) を 7列構造 (分類列付き) にマイグレーション
 * 由来に基づいて分類を自動推定
 */
function migrateWordPoolToV2() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.WORD_POOL);
  if (!ws) { initWordPoolSheet(ss); return; }
  var headers = ws.getRange(1, 1, 1, Math.max(1, ws.getLastColumn())).getValues()[0];
  if (String(headers[3] || '').trim() === '分類') {
    Logger.log('語彙プールは既に新構造');
    return;
  }
  var lastRow = ws.getLastRow();
  if (lastRow < 1) return;
  var oldData = (lastRow >= 2) ? ws.getRange(2, 1, lastRow - 1, 6).getValues() : [];
  ws.clear();
  ws.getRange(1, 1, 1, 7).setValues([['ジャンル', 'ワード', '由来', '分類', '初出日', '最終更新日', 'ヒット数']])
    .setBackground('#4A86E8').setFontColor('#FFFFFF').setFontWeight('bold');
  ws.setFrozenRows(1);
  if (oldData.length === 0) return;
  var newData = oldData.map(function(r) {
    var src = String(r[2] || '');
    var cls = (src.indexOf('タグ') >= 0 || src.indexOf('サブジャンル') >= 0) ? 'sub' : 'main';
    return [r[0], r[1], r[2], cls, r[3], r[4], r[5]];
  });
  ws.getRange(2, 1, newData.length, 7).setValues(newData);
  Logger.log('語彙プール ' + newData.length + '行を分類列付き構造に移行');
}

/**
 * タグ辞書シート初期化（タグID → タググループ名 / タグ名）
 */
function initTagDictSheet(ss) {
  return getOrCreateSheet(ss, SHEET_NAMES.TAG_DICT, [
    'タグID', 'タググループ名', 'タグ名', '取得日'
  ]);
}

/**
 * 推奨ワードシート初期化（モード別2枚: 推奨ワード_中国輸入 / 推奨ワード_国内メーカー）
 * 列: 日付, 出現回数, 区分, ジャンル, メインワード, サブワード, 評価, 新規参入, URL1..URL6
 */
function initSuggestSheet(ss, mode) {
  var headers = ['日付', '出現回数', '区分', 'ジャンル', 'メインワード', 'サブワード', '評価', '新規参入'];
  for (var u = 1; u <= HIDDEN_GEM_URL_COUNT; u++) headers.push('URL' + u);
  return getOrCreateSheet(ss, getSuggestSheetName(mode), headers);
}

/**
 * 消耗品ワード初期登録（除外ワードシートに 種類=消耗品 で追加）
 * 既に登録済みのワードはスキップ
 */
function seedDisposableWords() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  initExcludeWordsSheet(ss);
  var ws = ss.getSheetByName(SHEET_NAMES.EXCLUDES);

  // D列(種類)の入力規則を最新版に更新
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['!検索タグ', '!マーケ', '!品質', 'デザイン', '機能', '季節タグ', '状態', '汎用語', '補助語', '自動検出', 'ブランド', '消耗品', '装飾語'], true)
    .setAllowInvalid(true).build();
  ws.getRange(2, 4, Math.max(1, ws.getMaxRows() - 1), 1).setDataValidation(typeRule);

  var disposables = [
    '詰め替え', '詰替', 'リフィル', 'refill',
    '替え', '替刃', '替芯', '替えブラシ', '替えパッド', '替えヘッド',
    '交換用', '補充用', '使い捨て', '消耗',
    'カートリッジ', 'スペア', '予備',
    'ストック', 'パック',
    '使い切り'
  ];
  var data = ws.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < data.length; i++) {
    var w = String(data[i][2] || '').trim();
    if (w) existing[w] = true;
  }
  var newRows = [];
  for (var j = 0; j < disposables.length; j++) {
    if (!existing[disposables[j]]) {
      newRows.push([true, true, disposables[j], '消耗品', '']);
    }
  }
  if (newRows.length > 0) {
    var startRow = ws.getLastRow() + 1;
    ws.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
    ws.getRange(startRow, 1, newRows.length, 2).insertCheckboxes();
    Logger.log('消耗品ワード ' + newRows.length + '件追加');
  } else {
    Logger.log('消耗品ワードは既に登録済み');
  }
}

/**
 * 除外ワードシートから 種類=消耗品 のワードだけを返す
 */
function loadDisposableWords() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.EXCLUDES);
  if (!ws) return {};
  var data = ws.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var word = String(data[i][2] || '').trim();
    var type = String(data[i][3] || '').trim();
    if (word && type === '消耗品') map[word] = true;
  }
  return map;
}

/**
 * 調査済みワードシートを読み、ジャンル+メインワード をキーにステータスMap返す
 * @return {Object} { "ジャンル::メインワード": {lastDate, status, month, memo} }
 */
function loadSurveyedWords() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAMES.SURVEYED);
  if (!ws || ws.getLastRow() < 2) return {};
  var data = ws.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var word  = String(data[i][0] || '').trim();
    var genre = String(data[i][1] || '').trim();
    if (!word || !genre) continue;
    map[genre + '::' + word] = {
      lastDate : data[i][2],
      status   : String(data[i][3] || '').trim(),
      month    : Number(data[i][4] || 0),
      memo     : String(data[i][5] || ''),
    };
  }
  return map;
}

/**
 * モード別ランキングシートから過去daysBack日以内の商品Map（rank参照用）
 * 新規参入判定など過去履歴参照に使う
 */
function getPastKeywordMap(daysBack, mode) {
  daysBack = daysBack || 30;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(getRankingSheetName(mode));
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
