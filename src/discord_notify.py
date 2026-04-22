"""
discord_notify.py - Discord Webhook通知モジュール
毎朝、ジャンル別の注目ワードサマリーを送信する
"""

import requests
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

MAX_MESSAGE_LENGTH = 1900  # Discordの上限2000文字より少し短め


def send_daily_digest(webhook_url: str, all_results: list[dict], date_str: str):
    """
    全ジャンルの日次ダイジェストをDiscordに送信する
    ジャンルごとに分割して複数メッセージで送る
    """
    # ヘッダーメッセージ
    header = f"## 📊 楽天トレンドワード日報 - {date_str}\n"
    _send(webhook_url, header)

    # ジャンルごとにまとめる
    genre_results = {}
    for r in all_results:
        genre = r["genre"]
        if genre not in genre_results:
            genre_results[genre] = {"hidden_gem": [], "trending": [], "saturated": []}
        genre_results[genre][r["classification"]].append(r)

    for genre, categories in genre_results.items():
        msg = f"### 【{genre}】\n"

        # 🎯 隠れた狙い目（top 5）
        hidden = sorted(categories["hidden_gem"], key=lambda x: x["final_score"], reverse=True)[:5]
        if hidden:
            msg += "**🎯 隠れた狙い目**\n"
            for r in hidden:
                new_label = " 🆕" if r.get("is_new") else ""
                msg += f"　`{r['keyword']}`　出現{r['count']}回 / 平均{r['avg_rank']:.0f}位{new_label}\n"

        # 📈 注目ワード（top 5）
        trending = sorted(categories["trending"], key=lambda x: x["final_score"], reverse=True)[:5]
        if trending:
            msg += "**📈 注目ワード**\n"
            for r in trending:
                new_label = " 🆕" if r.get("is_new") else ""
                msg += f"　`{r['keyword']}`　出現{r['count']}回 / 平均{r['avg_rank']:.0f}位{new_label}\n"

        # ⚠️ 飽和（件数だけ）
        sat_count = len(categories["saturated"])
        if sat_count > 0:
            msg += f"**⚠️ 飽和状態**: {sat_count}ワード（詳細はスプシ参照）\n"

        msg += "\n"

        if len(msg) > MAX_MESSAGE_LENGTH:
            msg = msg[:MAX_MESSAGE_LENGTH] + "...\n"

        _send(webhook_url, msg)

    # フッター
    footer = "📋 詳細は Google Spreadsheet を確認してください。"
    _send(webhook_url, footer)


def send_error_alert(webhook_url: str, error_msg: str):
    """エラー発生時にアラートを送信"""
    msg = f"⛔ **エラー発生** ({datetime.now().strftime('%Y-%m-%d %H:%M')})\n```{error_msg[:500]}```"
    _send(webhook_url, msg)


def send_keepa_report(webhook_url: str, genre: str, keepa_data: list[dict], date_str: str):
    """Keepa分析結果をDiscordに送信"""
    msg = f"### 🔍 Keepa補完データ【{genre}】- {date_str}\n"
    for item in keepa_data[:10]:  # 上位10件
        msg += (
            f"　**{item.get('title', '')[:40]}**\n"
            f"　　BSR: {item.get('bsr', 'N/A')} / 出品者数: {item.get('seller_count', 'N/A')}\n"
            f"　　[Amazon]({item.get('url', '')})\n"
        )
    _send(webhook_url, msg)


def _send(webhook_url: str, content: str):
    """Discord Webhookにメッセージを送信"""
    try:
        resp = requests.post(webhook_url, json={"content": content}, timeout=10)
        resp.raise_for_status()
    except Exception as e:
        logger.error(f"Discord send error: {e}")
