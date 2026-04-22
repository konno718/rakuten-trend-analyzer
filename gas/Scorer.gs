// Scorer.gs - キーワードのスコアリングと分類

function rankToScore(rank, maxRank) {
  maxRank = maxRank || RANKING_TOP_N;
  return Math.max(1, maxRank - rank + 1);
}

function classifyKeyword(count) {
  if (count <= SCORE_RULES.HIDDEN_GEM.max) return 'hidden_gem';
  if (count <= SCORE_RULES.TRENDING.max)   return 'trending';
  return 'saturated';
}

function aggregateKeywords(processedItems, genreName, dateStr) {
  var kwMap = {};

  for (var i = 0; i < processedItems.length; i++) {
    var item = processedItems[i];
    if (item.excluded) continue;
    var rankScore = rankToScore(item.rank);

    for (var j = 0; j < item.keywords.length; j++) {
      var kw = item.keywords[j];
      if (!kwMap[kw]) {
        kwMap[kw] = { count: 0, rankScores: [], ranks: [], products: [] };
      }
      kwMap[kw].count++;
      kwMap[kw].rankScores.push(rankScore);
      kwMap[kw].ranks.push(item.rank);
      kwMap[kw].products.push({
        rank     : item.rank,
        itemCode : item.itemCode,
        itemUrl  : item.itemUrl,
        itemName : item.itemName.substring(0, 50),
      });
    }
  }

  var results = [];
  var keys = Object.keys(kwMap);

  for (var m = 0; m < keys.length; m++) {
    var kw = keys[m];
    var data = kwMap[kw];
    var count = data.count;
    var sumRanks = 0;
    var sumScores = 0;
    for (var n = 0; n < data.ranks.length; n++) {
      sumRanks  += data.ranks[n];
      sumScores += data.rankScores[n];
    }
    var avgRank      = sumRanks / count;
    var avgRankScore = sumScores / count;
    var classification = classifyKeyword(count);
    var boost = (classification === 'hidden_gem') ? 1.5 : 1.0;
    var finalScore = Math.round(avgRankScore * boost * 10) / 10;

    results.push({
      date           : dateStr,
      genre          : genreName,
      keyword        : kw,
      count          : count,
      avgRank        : Math.round(avgRank * 10) / 10,
      finalScore     : finalScore,
      classification : classification,
      products       : data.products.sort(function(a, b) { return a.rank - b.rank; }),
      isNew          : false,
    });
  }

  results.sort(function(a, b) { return b.finalScore - a.finalScore; });
  Logger.log('[' + genreName + '] ' + results.length + 'キーワード集計完了');
  return results;
}

function detectHighFrequencyCandidates(allResults, totalGenres) {
  var kwGenreCount = {};
  for (var i = 0; i < allResults.length; i++) {
    var kw = allResults[i].keyword;
    kwGenreCount[kw] = (kwGenreCount[kw] || 0) + 1;
  }
  var threshold = totalGenres * 0.3;
  var candidates = [];
  var keys = Object.keys(kwGenreCount);
  for (var j = 0; j < keys.length; j++) {
    if (kwGenreCount[keys[j]] >= threshold) candidates.push(keys[j]);
  }
  return candidates;
}
