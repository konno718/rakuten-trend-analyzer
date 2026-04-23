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

/**
 * 同一商品セット（= 同一URLセット）を指すキーワードをグループ化し、
 * 代表1ワード + 類義ワード配列 に集約する。
 * 代表選定: count desc → keyword.length asc
 */
function groupByProductFingerprint(results) {
  var groups = {};
  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    var urls = (r.products || []).map(function(p){ return p.itemUrl; });
    urls.sort();
    var fp = urls.join('|');
    if (!groups[fp]) groups[fp] = [];
    groups[fp].push(r);
  }
  var merged = [];
  var fpKeys = Object.keys(groups);
  for (var k = 0; k < fpKeys.length; k++) {
    var bucket = groups[fpKeys[k]];
    bucket.sort(function(a, b) {
      if (b.count !== a.count) return b.count - a.count;
      return a.keyword.length - b.keyword.length;
    });
    var rep = bucket[0];
    var synonyms = [];
    for (var j = 1; j < bucket.length; j++) synonyms.push(bucket[j].keyword);
    merged.push({
      date           : rep.date,
      genre          : rep.genre,
      keyword        : rep.keyword,
      synonyms       : synonyms,
      count          : rep.count,
      avgRank        : rep.avgRank,
      finalScore     : rep.finalScore,
      classification : rep.classification,
      products       : rep.products,
      isNew          : rep.isNew,
    });
  }
  merged.sort(function(a, b) { return b.finalScore - a.finalScore; });
  return merged;
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
