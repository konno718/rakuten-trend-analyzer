// Discord.gs - Discord Webhook通知

var MAX_MSG_LEN = 1900;

function sendDiscordMessage(webhookUrl, content) {
  if (!webhookUrl) return;
  try {
    UrlFetchApp.fetch(webhookUrl, {
      method             : 'post',
      contentType        : 'application/json',
      payload            : JSON.stringify({ content: content }),
      muteHttpExceptions : true,
    });
  } catch(e) {
    Logger.log('Discord送信エラー: ' + e);
  }
}

function sendDailyDigest(webhookUrl, allResults, dateStr) {
  if (!webhookUrl) return;

  sendDiscordMessage(webhookUrl, '## 📊 楽天トレンドワード日報 - ' + dateStr);

  var byGenre = {};
  for (var i = 0; i < allResults.length; i++) {
    var r = allResults[i];
    if (!byGenre[r.genre]) {
      byGenre[r.genre] = { hidden_gem: [], trending: [], saturated: [] };
    }
    byGenre[r.genre][r.classification].push(r);
  }

  var genres = Object.keys(byGenre);
  for (var gi = 0; gi < genres.length; gi++) {
    var genre = genres[gi];
    var cats  = byGenre[genre];
    var msg   = '### 【' + genre + '】\n';

    var hidden = cats.hidden_gem.sort(function(a,b){return b.finalScore-a.finalScore;}).slice(0,5);
    if (hidden.length > 0) {
      msg += '**🎯 隠れた狙い目**\n';
      for (var h = 0; h < hidden.length; h++) {
        var r = hidden[h];
        msg += '　`' + r.keyword + '`　' + r.count + '回 / 平均' + Math.round(r.avgRank) + '位' + (r.isNew ? ' 🆕' : '') + '\n';
      }
    }

    var trending = cats.trending.sort(function(a,b){return b.finalScore-a.finalScore;}).slice(0,5);
    if (trending.length > 0) {
      msg += '**📈 注目ワード**\n';
      for (var t = 0; t < trending.length; t++) {
        var r = trending[t];
        msg += '　`' + r.keyword + '`　' + r.count + '回 / 平均' + Math.round(r.avgRank) + '位' + (r.isNew ? ' 🆕' : '') + '\n';
      }
    }

    if (cats.saturated.length > 0) {
      msg += '**⚠️ 飽和**: ' + cats.saturated.length + 'ワード（スプシ参照）\n';
    }

    if (msg.length > MAX_MSG_LEN) msg = msg.substring(0, MAX_MSG_LEN) + '...\n';
    sendDiscordMessage(webhookUrl, msg);
    Utilities.sleep(500);
  }

  sendDiscordMessage(webhookUrl,
    '📋 詳細は https://docs.google.com/spreadsheets/d/' + SHEET_ID + ' を確認してください。'
  );
}

function sendErrorAlert(webhookUrl, errorMsg) {
  if (!webhookUrl) return;
  var now = new Date().toLocaleString('ja-JP');
  var msg = '⛔ **エラー発生** (' + now + ')\n' + String(errorMsg).substring(0, 500);
  sendDiscordMessage(webhookUrl, msg);
}
