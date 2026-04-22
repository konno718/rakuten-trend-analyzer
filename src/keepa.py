"""
keepa.py - Keepa APIとの連携モジュール
1日1〜2ジャンルのローテーションで楽天キーワードをAmazonで補完検索する
"""

import requests
import logging
import time
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)

KEEPA_BASE_URL = "https://api.keepa.com"
DOMAIN_JP = 5  # Amazon.co.jp


class KeepaClient:
    def __init__(self, api_key: str):
        self.api_key = api_key

    def get_remaining_tokens(self) -> int:
        """残トークン数を確認"""
        try:
            resp = requests.get(
                f"{KEEPA_BASE_URL}/token",
                params={"key": self.api_key},
                timeout=10
            )
            resp.raise_for_status()
            data = resp.json()
            return data.get("tokensLeft", 0)
        except Exception as e:
            logger.error(f"Keepa token check error: {e}")
            return 0

    def get_bestsellers(self, category_id: int, limit: int = 100) -> list[str]:
        """
        指定カテゴリのAmazonベストセラーASINリストを取得
        Keepaトークン消費: 約50トークン

        Returns: list of ASIN strings
        """
        try:
            resp = requests.get(
                f"{KEEPA_BASE_URL}/bestsellers",
                params={
                    "key": self.api_key,
                    "domain": DOMAIN_JP,
                    "category": category_id,
                },
                timeout=30
            )
            resp.raise_for_status()
            data = resp.json()
            asins = data.get("asinList", [])
            return asins[:limit]
        except Exception as e:
            logger.error(f"Keepa bestsellers error: {e}")
            return []

    def get_products(self, asins: list[str]) -> list[dict]:
        """
        ASIN一覧から商品詳細を取得
        Keepaトークン消費: ASINごとに約10トークン

        Returns: list of product dicts
        """
        if not asins:
            return []

        # Keepa APIは一度に最大100 ASINまで
        chunks = [asins[i:i+100] for i in range(0, len(asins), 100)]
        all_products = []

        for chunk in chunks:
            try:
                resp = requests.get(
                    f"{KEEPA_BASE_URL}/product",
                    params={
                        "key": self.api_key,
                        "domain": DOMAIN_JP,
                        "asin": ",".join(chunk),
                        "stats": 90,  # 過去90日の統計
                        "offers": 20,
                    },
                    timeout=30
                )
                resp.raise_for_status()
                data = resp.json()

                for product in data.get("products", []):
                    all_products.append(self._parse_product(product))

                time.sleep(1)  # API負荷軽減

            except Exception as e:
                logger.error(f"Keepa product fetch error: {e}")

        return all_products

    def _parse_product(self, raw: dict) -> dict:
        """商品データを整形する"""
        asin = raw.get("asin", "")
        title = raw.get("title", "")

        # BSR（ベストセラーランク）を取得
        sales_ranks = raw.get("salesRanks", {})
        bsr = None
        for cat_id, ranks in sales_ranks.items():
            if ranks:
                # 最新のランクを取得（rankは[timestamp, value]のペアの配列）
                bsr = ranks[-1] if ranks else None
                break

        # 出品者数
        stats = raw.get("stats", {})
        seller_count = stats.get("current", {}).get("offerCountNew", None) if stats else None

        # 価格（現在の最安値）
        current_price = stats.get("current", {}).get("buyBoxPrice", None) if stats else None
        if current_price:
            current_price = current_price / 100  # Keepaは円×100で格納

        # レビュー数
        review_count = raw.get("reviewCount", None)

        return {
            "asin": asin,
            "title": title[:80] if title else "",
            "url": f"https://www.amazon.co.jp/dp/{asin}",
            "bsr": bsr,
            "seller_count": seller_count,
            "current_price": current_price,
            "review_count": review_count,
        }

    def search_by_keyword(self, keyword: str, limit: int = 20) -> list[dict]:
        """
        キーワードでAmazon商品を検索（Keepa経由）
        ※ KeepaはAmazon検索をAPIでは直接提供していないため、
        　 ベストセラー一覧をタイトルでフィルタするアプローチを使用
        """
        # このメソッドは将来の拡張用プレースホルダー
        # 実際には search API があれば使う
        logger.info(f"Keepa keyword search for: {keyword} (placeholder)")
        return []


def select_genres_for_today(genres: list[dict], rotation_data: dict,
                             date_str: str, genres_per_day: int = 2) -> list[dict]:
    """
    1日に処理するジャンルをローテーションで選択する

    Args:
        genres: 全ジャンル設定リスト
        rotation_data: DBから取得したローテーション情報
        date_str: 今日の日付
        genres_per_day: 1日に処理するジャンル数（デフォルト2）

    Returns:
        今日処理すべきジャンルのリスト
    """
    if not genres:
        return []

    # Keepaカテゴリが設定されているジャンルのみ対象
    keepa_genres = [g for g in genres if g.get("keepa_category")]
    if not keepa_genres:
        return []

    # 最後に実行した日付でソート（古い順）
    def last_run_key(g):
        info = rotation_data.get(g["genre_name"])
        if not info or not info.get("last_run_date"):
            return "0000-00-00"  # 一度も実行していないものを優先
        return info["last_run_date"]

    sorted_genres = sorted(keepa_genres, key=last_run_key)
    selected = sorted_genres[:genres_per_day]

    logger.info(f"Keepa today's genres: {[g['genre_name'] for g in selected]}")
    return selected
