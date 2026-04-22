"""
sheets.py - Google Sheetsとの連携モジュール
- 設定シートからジャンルURLを読み込む
- 日次結果をスプレッドシートに書き込む
"""

import logging
from datetime import datetime
from typing import Optional
import gspread
from google.oauth2.service_account import Credentials

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# スプレッドシートのシート名
SHEET_SETTINGS  = "設定"           # ジャンルURLを入力するシート
SHEET_WORDS     = "ワード集計"     # キーワード集計（日次追記）
SHEET_PRODUCTS  = "商品一覧"       # 商品URL×順位一覧
SHEET_EXCLUDE   = "除外候補"       # 自動検出した除外候補（人間レビュー用）
SHEET_SUMMARY   = "サマリー"       # AI分析結果やサマリー


class SheetsManager:
    def __init__(self, credentials_path: str, sheet_id: str):
        creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
        self.gc = gspread.authorize(creds)
        self.sheet_id = sheet_id
        self._wb = None

    @property
    def wb(self):
        if self._wb is None:
            self._wb = self.gc.open_by_key(self.sheet_id)
        return self._wb

    def _get_or_create_sheet(self, name: str, headers: list[str] = None):
        """シートを取得、なければ作成してヘッダーを設定"""
        try:
            ws = self.wb.worksheet(name)
        except gspread.WorksheetNotFound:
            ws = self.wb.add_worksheet(title=name, rows=10000, cols=20)
            if headers:
                ws.append_row(headers)
            logger.info(f"Created sheet: {name}")
        return ws

    # --------------------------------------------------------
    # 設定シート: ジャンルURL読み込み
    # --------------------------------------------------------
    def read_genre_configs(self) -> list[dict]:
        """
        設定シートからジャンル設定を読み込む

        シート形式（設定シート）:
        | A: ジャンル名 | B: 楽天ランキングURL | C: 有効(○/×) | D: Keepaカテゴリ（任意）|

        Returns:
            list of dict: [{
                'genre_name': str,
                'rakuten_url': str,
                'enabled': bool,
                'keepa_category': str (optional)
            }]
        """
        try:
            ws = self.wb.worksheet(SHEET_SETTINGS)
        except gspread.WorksheetNotFound:
            # 設定シートがなければサンプル付きで作成
            self._create_settings_template()
            logger.warning("設定シートを作成しました。ジャンルURLを入力してください。")
            return []

        rows = ws.get_all_values()
        configs = []

        for i, row in enumerate(rows):
            if i == 0:  # ヘッダー行をスキップ
                continue
            if len(row) < 2 or not row[0].strip() or not row[1].strip():
                continue

            genre_name = row[0].strip()
            rakuten_url = row[1].strip()
            enabled = len(row) < 3 or row[2].strip() not in ["×", "x", "X", "false", "0", "無効"]
            keepa_category = row[3].strip() if len(row) > 3 else ""

            configs.append({
                "genre_name": genre_name,
                "rakuten_url": rakuten_url,
                "enabled": enabled,
                "keepa_category": keepa_category,
            })

        active = [c for c in configs if c["enabled"]]
        logger.info(f"Loaded {len(active)} active genre configs from Sheets")
        return active

    def _create_settings_template(self):
        """設定シートのテンプレートを作成"""
        ws = self._get_or_create_sheet(SHEET_SETTINGS, headers=[
            "ジャンル名", "楽天ランキングURL", "有効(○/×)", "Keepaカテゴリ（任意）", "メモ"
        ])
        # サンプル行
        samples = [
            ["インテリア・雑貨", "https://ranking.rakuten.co.jp/daily/100533/", "○", "", ""],
            ["ペット用品",       "https://ranking.rakuten.co.jp/daily/101213/", "○", "", ""],
            ["アウトドア",       "https://ranking.rakuten.co.jp/daily/101070/", "○", "", ""],
            ["美容・健康",       "https://ranking.rakuten.co.jp/daily/100227/", "○", "", ""],
            ["キッチン用品",     "https://ranking.rakuten.co.jp/daily/100227/", "×", "", "URLは適宜変更"],
        ]
        for row in samples:
            ws.append_row(row)

    # --------------------------------------------------------
    # ワード集計シートへの書き込み
    # --------------------------------------------------------
    def write_word_stats(self, results: list[dict], date_str: str, db=None):
        """
        キーワード集計結果をスプレッドシートに書き込む
        （既存データに日次追記）
        """
        ws = self._get_or_create_sheet(SHEET_WORDS, headers=[
            "日付", "ジャンル", "キーワード", "出現回数", "平均順位",
            "スコア", "分類", "新規参入", "連続日数"
        ])

        rows_to_add = []
        for r in results:
            is_new = False
            consecutive = 0
            if db:
                is_new = db.is_new_entrant(r["keyword"], r["genre"], date_str)
                consecutive = db.get_consecutive_days(r["keyword"], r["genre"], date_str)

            classification_label = {
                "hidden_gem": "🎯隠れた狙い目",
                "trending": "📈注目ワード",
                "saturated": "⚠️飽和状態",
            }.get(r["classification"], r["classification"])

            rows_to_add.append([
                date_str,
                r["genre"],
                r["keyword"],
                r["count"],
                r["avg_rank"],
                r["final_score"],
                classification_label,
                "🆕新規" if is_new else "",
                consecutive,
            ])

        if rows_to_add:
            ws.append_rows(rows_to_add)
            logger.info(f"Wrote {len(rows_to_add)} word stats to Sheets")

    # --------------------------------------------------------
    # 商品一覧シートへの書き込み
    # --------------------------------------------------------
    def write_products(self, results: list[dict], date_str: str):
        """
        各キーワードの該当商品URL一覧を書き込む
        """
        ws = self._get_or_create_sheet(SHEET_PRODUCTS, headers=[
            "日付", "ジャンル", "キーワード", "分類", "順位", "商品名（短縮）", "商品URL"
        ])

        rows_to_add = []
        for r in results:
            # hidden_gem と trending のみ商品URLを記録（saturatedは省略）
            if r["classification"] == "saturated":
                continue

            label = {
                "hidden_gem": "🎯隠れた狙い目",
                "trending": "📈注目ワード",
            }.get(r["classification"], "")

            for p in r.get("products", []):
                rows_to_add.append([
                    date_str,
                    r["genre"],
                    r["keyword"],
                    label,
                    p["rank"],
                    p["itemName"],
                    p["itemUrl"],
                ])

        if rows_to_add:
            ws.append_rows(rows_to_add)
            logger.info(f"Wrote {len(rows_to_add)} product rows to Sheets")

    # --------------------------------------------------------
    # 除外候補シートへの書き込み
    # --------------------------------------------------------
    def write_exclude_candidates(self, candidates: list[str], date_str: str):
        """自動検出した除外候補ワードをシートに追記"""
        if not candidates:
            return

        ws = self._get_or_create_sheet(SHEET_EXCLUDE, headers=[
            "検出日", "候補ワード", "レビュー（除外する場合は○）", "メモ"
        ])

        existing = set(ws.col_values(2))  # すでにある候補は追加しない
        new_candidates = [c for c in candidates if c not in existing]

        if new_candidates:
            rows = [[date_str, c, "", ""] for c in new_candidates]
            ws.append_rows(rows)
            logger.info(f"Added {len(new_candidates)} exclude candidates to Sheets")

    # --------------------------------------------------------
    # サマリーシートへの書き込み（AI分析結果）
    # --------------------------------------------------------
    def write_summary(self, summary_text: str, date_str: str):
        """AI分析結果をサマリーシートに書き込む"""
        ws = self._get_or_create_sheet(SHEET_SUMMARY, headers=["日付", "分析結果"])
        ws.append_row([date_str, summary_text])
        logger.info(f"Wrote summary to Sheets for {date_str}")
