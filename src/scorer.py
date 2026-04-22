"""
scorer.py - キーワードのスコアリングと分類モジュール

スコア設計:
  - 順位スコア: 1位=300点, 300位=1点（線形）
  - 出現回数による分類:
      1〜3回: 隠れた狙い目（hidden_gem）
      4〜9回: 注目ワード（trending）
      10回以上: 飽和状態（saturated）
  - 最終スコア = 合計順位スコア / 出現回数（平均化）× 出現回数ボーナス
"""

import logging
from collections import defaultdict
from typing import Optional

logger = logging.getLogger(__name__)

# 出現回数→分類のマッピング
CLASSIFICATION_RULES = {
    "hidden_gem":  (1, 3),    # 隠れた狙い目
    "trending":    (4, 9),    # 注目ワード
    "saturated":   (10, 9999) # 飽和状態
}

def classify_keyword(count: int) -> str:
    for label, (lo, hi) in CLASSIFICATION_RULES.items():
        if lo <= count <= hi:
            return label
    return "unknown"


def rank_to_score(rank: int, max_rank: int = 300) -> float:
    """順位をスコアに変換（1位=max_rank点、max_rank位=1点）"""
    return max(1, max_rank - rank + 1)


def aggregate_keywords(processed_items: list[dict], genre_name: str, date_str: str) -> list[dict]:
    """
    処理済み商品リストからキーワードを集計してスコアを計算する

    Returns:
        list of dict: [{
            'date': 日付,
            'genre': ジャンル名,
            'keyword': キーワード,
            'count': 出現回数,
            'total_rank_score': 合計順位スコア,
            'avg_rank': 平均順位,
            'final_score': 最終スコア,
            'classification': hidden_gem/trending/saturated,
            'product_ranks': [(rank, itemCode, itemUrl), ...],  # 該当商品の詳細
        }]
    """
    # キーワードごとに集計
    kw_data = defaultdict(lambda: {
        "count": 0,
        "rank_scores": [],
        "ranks": [],
        "products": [],
    })

    for item in processed_items:
        if item.get("excluded"):
            continue

        rank = item.get("rank", 999)
        rank_score = rank_to_score(rank)
        item_code = item.get("itemCode", "")
        item_url = item.get("itemUrl", "")
        item_name = item.get("itemName", "")

        for kw in item.get("keywords", []):
            kw_data[kw]["count"] += 1
            kw_data[kw]["rank_scores"].append(rank_score)
            kw_data[kw]["ranks"].append(rank)
            kw_data[kw]["products"].append({
                "rank": rank,
                "itemCode": item_code,
                "itemUrl": item_url,
                "itemName": item_name[:60],  # 表示用に短縮
            })

    # スコア計算と結果生成
    results = []
    for kw, data in kw_data.items():
        count = data["count"]
        avg_rank = sum(data["ranks"]) / count
        total_rank_score = sum(data["rank_scores"])

        # 最終スコア: 平均順位スコア（高順位=高スコアが基本）
        # hidden_gemにボーナス: 少ない出現回数でも上位にいれば高スコア
        avg_rank_score = total_rank_score / count
        classification = classify_keyword(count)

        # hidden_gemはスコアを1.5倍ブースト（発見しやすくする）
        boost = 1.5 if classification == "hidden_gem" else 1.0
        final_score = avg_rank_score * boost

        results.append({
            "date": date_str,
            "genre": genre_name,
            "keyword": kw,
            "count": count,
            "total_rank_score": round(total_rank_score, 1),
            "avg_rank": round(avg_rank, 1),
            "final_score": round(final_score, 1),
            "classification": classification,
            "products": sorted(data["products"], key=lambda x: x["rank"]),
        })

    # 最終スコアで降順ソート
    results.sort(key=lambda x: x["final_score"], reverse=True)

    logger.info(f"Aggregated {len(results)} keywords for genre '{genre_name}' on {date_str}")
    return results


def detect_high_frequency_candidates(all_results: list[dict], threshold_pct: float = 0.3) -> list[str]:
    """
    全ジャンルで一定割合以上に登場するワードを「汎用ワード候補」として検出する
    → 除外候補リストに追加してレビューを促す

    threshold_pct: ジャンル数に対する登場割合（デフォルト30%）
    """
    from collections import Counter
    genre_set = set(r["genre"] for r in all_results)
    total_genres = len(genre_set)
    if total_genres == 0:
        return []

    kw_genre_count = Counter()
    for r in all_results:
        kw_genre_count[r["keyword"]] += 1  # ジャンルをまたいだ場合もカウント

    threshold = total_genres * threshold_pct
    candidates = [kw for kw, cnt in kw_genre_count.items() if cnt >= threshold]
    return candidates
