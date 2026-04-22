"""
analyzer.py - Claude AIを使った週次・任意タイミングの分析モジュール
毎日は使わず、週1回か手動実行時のみ使用
"""

import json
import logging
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)


def build_analysis_prompt(genre_summaries: dict, date_str: str) -> str:
    """
    Claude APIに送るプロンプトを構築する

    genre_summaries: {
        "ジャンル名": {
            "hidden_gem": [{"keyword": str, "count": int, "avg_rank": float, "is_new": bool}, ...],
            "trending": [...],
            "saturated": [...],
        }
    }
    """
    prompt = f"""あなたは中国輸入ビジネスの商品リサーチ専門家です。
以下は楽天市場のランキングデータから抽出したキーワード分析結果（{date_str}）です。

## 分析目的
- 中国から仕入れて日本のEC市場で販売できる商品のキーワードを見つける
- 半年後を見据えた仕入れ判断の参考にする
- 需要はあるが競合が少ない「穴場」を発見する

## データ
"""

    for genre, data in genre_summaries.items():
        prompt += f"\n### ジャンル: {genre}\n"

        if data.get("hidden_gem"):
            prompt += "**隠れた狙い目（出現1〜3回）:**\n"
            for kw in data["hidden_gem"][:10]:
                new_flag = " [新規]" if kw.get("is_new") else ""
                prompt += f"- {kw['keyword']} (出現{kw['count']}回, 平均{kw['avg_rank']:.0f}位{new_flag})\n"

        if data.get("trending"):
            prompt += "**注目ワード（出現4〜9回）:**\n"
            for kw in data["trending"][:10]:
                new_flag = " [新規]" if kw.get("is_new") else ""
                prompt += f"- {kw['keyword']} (出現{kw['count']}回, 平均{kw['avg_rank']:.0f}位{new_flag})\n"

    prompt += """
## 出力形式
以下の3つのカテゴリで分析結果をまとめてください:

1. **今週のトレンドワード**: 複数ジャンルにまたがって注目されているワードや、市場全体で盛り上がっているキーワード

2. **仕入れ候補ワード（狙い目）**: 中国輸入ビジネスに適していると思われる具体的な商品キーワード。理由も添えて。

3. **要注意ワード**: 避けるべき理由があるキーワード（飽和・季節限定・大手独占など）

各カテゴリ最大5件まで。箇条書きで簡潔に。"""

    return prompt


def analyze_with_claude(api_key: str, genre_summaries: dict, date_str: str) -> str:
    """
    Claude APIを呼び出して分析結果を取得する

    Returns:
        str: 分析テキスト
    """
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)

        prompt = build_analysis_prompt(genre_summaries, date_str)

        message = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=1500,
            messages=[{"role": "user", "content": prompt}]
        )
        result = message.content[0].text
        logger.info("Claude analysis completed")
        return result

    except ImportError:
        logger.warning("anthropic package not installed. Skipping AI analysis.")
        return "AI分析スキップ（anthropicパッケージ未インストール）"
    except Exception as e:
        logger.error(f"Claude API error: {e}")
        return f"AI分析エラー: {str(e)}"


def prepare_genre_summaries(all_results: list[dict], db=None, date_str: str = None) -> dict:
    """
    全ジャンルの結果をAI分析用に整形する
    """
    summaries = {}
    today = date_str or datetime.now().strftime("%Y-%m-%d")

    for r in all_results:
        genre = r["genre"]
        if genre not in summaries:
            summaries[genre] = {"hidden_gem": [], "trending": [], "saturated": []}

        is_new = False
        if db and date_str:
            is_new = db.is_new_entrant(r["keyword"], genre, today)

        entry = {
            "keyword": r["keyword"],
            "count": r["count"],
            "avg_rank": r["avg_rank"],
            "final_score": r["final_score"],
            "is_new": is_new,
        }
        summaries[genre][r["classification"]].append(entry)

    return summaries
