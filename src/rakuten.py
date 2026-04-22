"""
rakuten.py - 楽天市場ランキングAPI取得モジュール
楽天ランキングURLからジャンルIDを抽出し、上位300件を取得する
"""

import re
import time
import requests
import logging
from typing import Optional

logger = logging.getLogger(__name__)


def parse_genre_id_from_url(url: str) -> Optional[str]:
    """
    楽天ランキングURLからジャンルIDを抽出する
    例: https://ranking.rakuten.co.jp/daily/100533/ → "100533"
    """
    match = re.search(r'/daily/(\d+)/?', url)
    if match:
        return match.group(1)
    match = re.search(r'/(\d+)/?$', url.rstrip('/'))
    if match:
        return match.group(1)
    return None


def fetch_ranking(app_id: str, genre_id: str, max_items: int = 300) -> list[dict]:
    """
    楽天ランキングAPIから上位max_items件を取得する
    1ページ30件、最大10ページ = 最大300件

    Returns:
        list of dict: 商品情報のリスト
        [{
            'rank': 順位,
            'itemCode': 商品コード,
            'itemName': 商品名,
            'itemUrl': 商品URL,
            'itemPrice': 価格,
            'genreId': ジャンルID,
        }]
    """
    base_url = "https://app.rakuten.co.jp/services/api/IchibaItem/Ranking/20220601"
    items_per_page = 30
    max_pages = min(10, (max_items + items_per_page - 1) // items_per_page)

    all_items = []
    current_rank = 1

    for page in range(1, max_pages + 1):
        params = {
            "applicationId": app_id,
            "genreId": genre_id,
            "page": page,
            "hits": items_per_page,
            "format": "json",
        }

        try:
            resp = requests.get(base_url, params=params, timeout=15)
            resp.raise_for_status()
            data = resp.json()
        except requests.RequestException as e:
            logger.error(f"Rakuten API error (page {page}): {e}")
            break
        except Exception as e:
            logger.error(f"Unexpected error (page {page}): {e}")
            break

        items = data.get("Items", [])
        if not items:
            break

        for item_wrapper in items:
            item = item_wrapper.get("Item", item_wrapper)
            all_items.append({
                "rank": current_rank,
                "itemCode": item.get("itemCode", ""),
                "itemName": item.get("itemName", ""),
                "itemUrl": item.get("itemUrl", ""),
                "itemPrice": item.get("itemPrice", 0),
                "genreId": genre_id,
            })
            current_rank += 1

            if len(all_items) >= max_items:
                break

        if len(all_items) >= max_items:
            break

        # API負荷軽減のため少し待機
        time.sleep(0.5)

    logger.info(f"Fetched {len(all_items)} items for genre {genre_id}")
    return all_items
