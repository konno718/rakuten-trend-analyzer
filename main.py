"""
main.py - 楽天トレンドワード収集・分析システム
メインオーケストレーター

使い方:
  python main.py          # 通常実行（ランキング取得 + Sheets書き込み + Discord通知）
  python main.py --ai     # AI分析も実行（週次推奨）
  python main.py --test   # テスト実行（最初の1ジャンルのみ、書き込みなし）
"""

import os
import sys
import logging
import argparse
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from dotenv import load_dotenv
from pathlib import Path

# パス設定
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

from src.rakuten import fetch_ranking, parse_genre_id_from_url
from src.extractor import load_exclude_words, process_items
from src.scorer import aggregate_keywords, detect_high_frequency_candidates
from src.database import TrendDatabase
from src.sheets import SheetsManager
from src.keepa import KeepaClient, select_genres_for_today
from src.discord_notify import send_daily_digest, send_error_alert, send_keepa_report
from src.analyzer import analyze_with_claude, prepare_genre_summaries

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(BASE_DIR / "data" / "run.log", encoding="utf-8"),
    ]
)
logger = logging.getLogger(__name__)


def process_genre(genre_config: dict, app_id: str, exclude_config: dict,
                  date_str: str, max_items: int = 300) -> list[dict]:
    """
    1ジャンルのランキングを取得してキーワード集計を行う（並列処理用）

    Returns:
        list of dict: キーワード集計結果
    """
    genre_name = genre_config["genre_name"]
    rakuten_url = genre_config["rakuten_url"]

    logger.info(f"Processing genre: {genre_name}")

    genre_id = parse_genre_id_from_url(rakuten_url)
    if not genre_id:
        logger.error(f"Failed to parse genre ID from URL: {rakuten_url}")
        return []

    # ランキング取得
    items = fetch_ranking(app_id, genre_id, max_items=max_items)
    if not items:
        logger.warning(f"No items for genre {genre_name}")
        return []

    # ジャンル名をitemsに追加
    for item in items:
        item["genreName"] = genre_name

    # キーワード抽出
    processed = process_items(items, exclude_config)

    # スコアリング
    results = aggregate_keywords(processed, genre_name, date_str)

    logger.info(f"Genre '{genre_name}': {len(results)} keywords found")
    return results


def main():
    parser = argparse.ArgumentParser(description="楽天トレンドワード収集システム")
    parser.add_argument("--ai", action="store_true", help="AI分析を実行する")
    parser.add_argument("--test", action="store_true", help="テストモード（書き込みなし）")
    parser.add_argument("--date", type=str, help="処理日付 YYYY-MM-DD（デフォルト: 今日）")
    args = parser.parse_args()

    # 環境変数読み込み
    load_dotenv(BASE_DIR / ".env")

    date_str = args.date or datetime.now().strftime("%Y-%m-%d")
    logger.info(f"=== 楽天トレンドワード収集開始: {date_str} ===")

    # 設定読み込み
    rakuten_app_id     = os.getenv("RAKUTEN_APP_ID")
    keepa_api_key      = os.getenv("KEEPA_API_KEY")
    google_sheet_id    = os.getenv("GOOGLE_SHEET_ID")
    google_creds_path  = os.getenv("GOOGLE_CREDENTIALS_PATH", str(BASE_DIR / "credentials.json"))
    discord_webhook    = os.getenv("DISCORD_WEBHOOK_URL")
    claude_api_key     = os.getenv("CLAUDE_API_KEY")
    max_items          = int(os.getenv("RANKING_TOP_N", "300"))

    if not rakuten_app_id:
        logger.error("RAKUTEN_APP_ID が設定されていません。.env ファイルを確認してください。")
        sys.exit(1)

    # 除外ワード読み込み
    exclude_config = load_exclude_words(str(BASE_DIR / "data" / "exclude_words.json"))

    # データベース初期化
    db = TrendDatabase(str(BASE_DIR / "data" / "trend_analyzer.db"))

    # Sheets読み込み
    try:
        sheets = SheetsManager(google_creds_path, google_sheet_id)
        genre_configs = sheets.read_genre_configs()
    except Exception as e:
        logger.error(f"Google Sheets接続エラー: {e}")
        if discord_webhook:
            send_error_alert(discord_webhook, f"Sheets接続エラー: {e}")
        sys.exit(1)

    if not genre_configs:
        logger.error("有効なジャンル設定がありません。設定シートを確認してください。")
        sys.exit(1)

    if args.test:
        genre_configs = genre_configs[:1]
        logger.info("テストモード: 最初の1ジャンルのみ処理")

    # ============================================================
    # メイン処理: ジャンルを並列処理
    # ============================================================
    all_results = []

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {
            executor.submit(
                process_genre, gc, rakuten_app_id, exclude_config, date_str, max_items
            ): gc["genre_name"]
            for gc in genre_configs
        }

        for future in as_completed(futures):
            genre_name = futures[future]
            try:
                results = future.result()
                all_results.extend(results)
            except Exception as e:
                logger.error(f"Genre '{genre_name}' processing error: {e}")
                if discord_webhook:
                    send_error_alert(discord_webhook, f"ジャンル '{genre_name}' エラー: {e}")

    if not all_results:
        logger.warning("結果が0件です。処理を中断します。")
        sys.exit(0)

    # ============================================================
    # 新規参入フラグを追加
    # ============================================================
    for r in all_results:
        r["is_new"] = db.is_new_entrant(r["keyword"], r["genre"], date_str)

    # ============================================================
    # DBに保存
    # ============================================================
    if not args.test:
        try:
            db.save_word_stats(all_results)
            logger.info("DB保存完了")
        except Exception as e:
            logger.error(f"DB保存エラー: {e}")

    # ============================================================
    # Google Sheets書き込み
    # ============================================================
    if not args.test:
        try:
            # 除外候補ワードを検出
            candidates = detect_high_frequency_candidates(all_results)

            sheets.write_word_stats(all_results, date_str, db=db)
            sheets.write_products(all_results, date_str)
            sheets.write_exclude_candidates(candidates, date_str)
            logger.info("Google Sheets書き込み完了")
        except Exception as e:
            logger.error(f"Sheets書き込みエラー: {e}")
            if discord_webhook:
                send_error_alert(discord_webhook, f"Sheets書き込みエラー: {e}")

    # ============================================================
    # Keepa補完（ローテーション）
    # ============================================================
    if keepa_api_key and not args.test:
        try:
            keepa = KeepaClient(keepa_api_key)
            remaining_tokens = keepa.get_remaining_tokens()
            logger.info(f"Keepa残トークン: {remaining_tokens}")

            if remaining_tokens > 500:
                rotation_data = db.get_keepa_rotation()
                today_genres = select_genres_for_today(genre_configs, rotation_data, date_str)

                for gc in today_genres:
                    cat_id = gc.get("keepa_category")
                    if not cat_id:
                        continue
                    try:
                        asins = keepa.get_bestsellers(int(cat_id), limit=50)
                        products = keepa.get_products(asins[:30])
                        db.update_keepa_rotation(gc["genre_name"], date_str)

                        if discord_webhook and products:
                            send_keepa_report(discord_webhook, gc["genre_name"], products, date_str)
                        logger.info(f"Keepa completed for {gc['genre_name']}: {len(products)} products")
                    except Exception as e:
                        logger.error(f"Keepa error for {gc['genre_name']}: {e}")
            else:
                logger.warning(f"Keepaトークン不足でスキップ: {remaining_tokens}")
        except Exception as e:
            logger.error(f"Keepa処理エラー: {e}")

    # ============================================================
    # AI分析（--ai オプション時のみ）
    # ============================================================
    if args.ai and claude_api_key:
        try:
            logger.info("AI分析を開始...")
            summaries = prepare_genre_summaries(all_results, db=db, date_str=date_str)
            analysis = analyze_with_claude(claude_api_key, summaries, date_str)

            if not args.test:
                sheets.write_summary(analysis, date_str)

            # Discordにも送信
            if discord_webhook:
                from src.discord_notify import _send
                _send(discord_webhook, f"### 🤖 AI分析結果 - {date_str}\n{analysis[:1500]}")

            logger.info("AI分析完了")
        except Exception as e:
            logger.error(f"AI分析エラー: {e}")

    # ============================================================
    # Discord日次通知
    # ============================================================
    if discord_webhook and not args.test:
        try:
            send_daily_digest(discord_webhook, all_results, date_str)
            logger.info("Discord通知送信完了")
        except Exception as e:
            logger.error(f"Discord通知エラー: {e}")

    logger.info(f"=== 完了: {len(all_results)}件のキーワードを処理 ===")


if __name__ == "__main__":
    main()
