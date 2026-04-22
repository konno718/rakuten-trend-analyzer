"""
extractor.py - 商品名からキーワードを抽出するモジュール
- 商品名の前半100文字に限定
- Janomeで形態素解析（日本語）
- 除外ワードリストでフィルタリング
- 複合語（2語まで）も抽出
"""

import json
import re
import logging
from pathlib import Path
from janome.tokenizer import Tokenizer

logger = logging.getLogger(__name__)

# Janomeトークナイザー（起動コストが高いのでモジュールレベルで初期化）
_tokenizer = None

def get_tokenizer():
    global _tokenizer
    if _tokenizer is None:
        _tokenizer = Tokenizer()
    return _tokenizer


def load_exclude_words(json_path: str) -> dict:
    """除外ワードJSONを読み込む"""
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    exclude_set = set()

    # 静的除外ワード
    static = data.get("static_excludes", {})
    for category, words in static.items():
        if category != "_comment" and isinstance(words, list):
            exclude_set.update(words)

    # 動的除外ワード（承認済みのみ）
    dynamic = data.get("dynamic_excludes", {}).get("words", [])
    for entry in dynamic:
        if isinstance(entry, dict) and entry.get("status") == "active":
            exclude_set.add(entry["word"])
        elif isinstance(entry, str):
            exclude_set.add(entry)

    return {
        "exclude_set": exclude_set,
        "category_excludes": data.get("category_excludes", {}),
        "raw": data,
    }


def is_product_excluded(item_name: str, category_excludes: dict) -> bool:
    """
    商品名に除外カテゴリのキーワードが含まれる場合、その商品自体を除外する
    （コンタクトレンズ、食品、医薬品など）
    """
    name_lower = item_name.lower()
    for cat, words in category_excludes.items():
        if cat == "_comment":
            continue
        for word in words:
            if word in name_lower or word in item_name:
                return True
    return False


def extract_keywords(item_name: str, exclude_set: set) -> list[str]:
    """
    商品名から有効なキーワードを抽出する

    処理:
    1. 前半100文字に限定
    2. 形態素解析で名詞を抽出
    3. 除外ワードをフィルタ
    4. 1文字のみのトークンは除外
    5. 複合語（連続する名詞2語）も生成
    """
    # 前半100文字に限定
    name = item_name[:100]

    # 英数字の商品型番・記号を除去（例: ABC-123）
    name = re.sub(r'[A-Za-z0-9]+[-/][A-Za-z0-9]+', '', name)

    t = get_tokenizer()
    tokens = list(t.tokenize(name))

    nouns = []
    for token in tokens:
        pos = token.part_of_speech.split(',')[0]
        pos2 = token.part_of_speech.split(',')[1] if ',' in token.part_of_speech else ''

        # 名詞のみ（固有名詞、一般名詞）
        if pos != '名詞':
            continue
        # 数詞、非自立語、接尾語は除外
        if pos2 in ['数', '非自立', '接尾', '代名詞']:
            continue

        surface = token.surface.strip()

        # 1文字は除外（ノイズが多い）
        if len(surface) <= 1:
            continue

        # 除外ワードリストに含まれるものをスキップ
        if surface in exclude_set:
            continue

        # 数字のみは除外
        if surface.isdigit():
            continue

        nouns.append(surface)

    # 複合語を生成（連続する名詞2つを結合）
    # 例: ["ソープ", "ディスペンサー"] → "ソープディスペンサー"
    compound_nouns = []
    for i in range(len(nouns) - 1):
        compound = nouns[i] + nouns[i + 1]
        if compound not in exclude_set:
            compound_nouns.append(compound)

    # 単語 + 複合語を合わせてユニーク化（順序維持）
    all_keywords = nouns + compound_nouns
    seen = set()
    result = []
    for kw in all_keywords:
        if kw not in seen:
            seen.add(kw)
            result.append(kw)

    return result


def process_items(items: list[dict], exclude_config: dict) -> list[dict]:
    """
    商品リストを処理してキーワードを抽出する

    Returns:
        list of dict: {
            'rank': 順位,
            'itemCode': 商品コード,
            'itemName': 商品名,
            'itemUrl': URL,
            'itemPrice': 価格,
            'keywords': [抽出キーワードリスト],
            'excluded': True/False（商品自体が除外カテゴリかどうか）
        }
    """
    exclude_set = exclude_config["exclude_set"]
    category_excludes = exclude_config["category_excludes"]
    results = []

    for item in items:
        item_name = item.get("itemName", "")
        excluded = is_product_excluded(item_name, category_excludes)

        if excluded:
            item["keywords"] = []
            item["excluded"] = True
        else:
            keywords = extract_keywords(item_name, exclude_set)
            item["keywords"] = keywords
            item["excluded"] = False

        results.append(item)

    return results
