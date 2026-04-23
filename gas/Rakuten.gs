// Rakuten.gs - 楽天市場ランキングAPI

function parseGenreIdFromUrl(url) {
  const match = url.match(/\/daily\/(\d+)\/?/);
  return match ? match[1] : null;
}

function fetchRakutenRanking(appId, genreId, maxItems) {
  maxItems = maxItems || RANKING_TOP_N;
  const baseUrl = 'https://app.rakuten.co.jp/services/api/IchibaItem/Ranking/20220601';
  const itemsPerPage = 30;
  const maxPages = Math.min(10, Math.ceil(maxItems / itemsPerPage));
  const allItems = [];
  let currentRank = 1;

  for (let page = 1; page <= maxPages; page++) {
    const params = {
      applicationId : appId,
      genreId       : genreId,
      page          : page,
      hits          : itemsPerPage,
      format        : 'json',
    };
    const qs = Object.entries(params)
      .map(function(kv) { return kv[0] + '=' + encodeURIComponent(kv[1]); })
      .join('&');

    try {
      const resp = UrlFetchApp.fetch(baseUrl + '?' + qs, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) { break; }

      const data = JSON.parse(resp.getContentText());
      const items = data.Items || [];
      if (items.length === 0) break;

      for (let i = 0; i < items.length; i++) {
        const item = items[i].Item || items[i];
        allItems.push({
          rank      : currentRank,
          itemCode  : item.itemCode  || '',
          itemName  : item.itemName  || '',
          itemUrl   : item.itemUrl   || '',
          itemPrice : item.itemPrice || 0,
          tagIds    : item.tagIds    || [],
          genreId   : genreId,
        });
        currentRank++;
        if (allItems.length >= maxItems) break;
      }
      if (allItems.length >= maxItems) break;
      Utilities.sleep(1200);

    } catch(e) {
      Logger.log('Rakuten API error page ' + page + ': ' + e);
      break;
    }
  }

  Logger.log('Fetched ' + allItems.length + ' items for genre ' + genreId);
  return allItems;
}
