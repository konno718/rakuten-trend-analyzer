"""
database.py - SQLiteによるデータ永続化モジュール
1年以上の蓄積に対応。季節分析・新規参入判定に使用。
"""

import sqlite3
import json
import logging
from datetime import datetime, timedelta
from pathlib import Path
from contextlib import contextmanager

logger = logging.getLogger(__name__)


class TrendDatabase:
    def __init__(self, db_path: str = "./data/trend_analyzer.db"):
        self.db_path = db_path
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextmanager
    def _conn(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _init_db(self):
        """テーブルを初期化する（存在しない場合のみ作成）"""
        with self._conn() as conn:
            conn.executescript("""
                -- 商品データ（日次）
                CREATE TABLE IF NOT EXISTS products (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    date        TEXT NOT NULL,
                    genre       TEXT NOT NULL,
                    rank        INTEGER NOT NULL,
                    item_code   TEXT NOT NULL,
                    item_name   TEXT,
                    item_url    TEXT,
                    item_price  INTEGER,
                    excluded    INTEGER DEFAULT 0,
                    created_at  TEXT DEFAULT (datetime('now', 'localtime'))
                );
                CREATE INDEX IF NOT EXISTS idx_products_date_genre ON products(date, genre);
                CREATE INDEX IF NOT EXISTS idx_products_item_code ON products(item_code);

                -- キーワード集計（日次）
                CREATE TABLE IF NOT EXISTS word_stats (
                    id              INTEGER PRIMARY KEY AUTOINCREMENT,
                    date            TEXT NOT NULL,
                    genre           TEXT NOT NULL,
                    keyword         TEXT NOT NULL,
                    count           INTEGER,
                    avg_rank        REAL,
                    final_score     REAL,
                    classification  TEXT,
                    products_json   TEXT,
                    created_at      TEXT DEFAULT (datetime('now', 'localtime'))
                );
                CREATE INDEX IF NOT EXISTS idx_word_stats_date_genre ON word_stats(date, genre);
                CREATE INDEX IF NOT EXISTS idx_word_stats_keyword ON word_stats(keyword);

                -- 除外候補ワード（人間レビュー用）
                CREATE TABLE IF NOT EXISTS exclude_candidates (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    word        TEXT UNIQUE NOT NULL,
                    reason      TEXT,
                    detected_at TEXT DEFAULT (datetime('now', 'localtime')),
                    status      TEXT DEFAULT 'pending',  -- pending / approved / rejected
                    reviewed_at TEXT
                );

                -- ジャンル設定（スプレッドシートから読み込んだ内容を記録）
                CREATE TABLE IF NOT EXISTS genre_config (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    genre_name  TEXT NOT NULL,
                    rakuten_url TEXT,
                    genre_id    TEXT,
                    keepa_category_id TEXT,
                    enabled     INTEGER DEFAULT 1,
                    last_keepa_run TEXT,
                    updated_at  TEXT DEFAULT (datetime('now', 'localtime'))
                );

                -- Keepaローテーション管理
                CREATE TABLE IF NOT EXISTS keepa_rotation (
                    genre_name      TEXT PRIMARY KEY,
                    last_run_date   TEXT,
                    last_run_order  INTEGER DEFAULT 0
                );
            """)
        logger.info(f"Database initialized: {self.db_path}")

    # --------------------------------------------------------
    # 商品データ保存
    # --------------------------------------------------------
    def save_products(self, items: list[dict], date_str: str):
        """商品データを一括保存"""
        with self._conn() as conn:
            conn.executemany("""
                INSERT INTO products (date, genre, rank, item_code, item_name, item_url, item_price, excluded)
                VALUES (:date, :genre, :rank, :item_code, :item_name, :item_url, :item_price, :excluded)
            """, [{
                "date": date_str,
                "genre": item.get("genreName", item.get("genreId", "")),
                "rank": item.get("rank", 0),
                "item_code": item.get("itemCode", ""),
                "item_name": item.get("itemName", ""),
                "item_url": item.get("itemUrl", ""),
                "item_price": item.get("itemPrice", 0),
                "excluded": 1 if item.get("excluded") else 0,
            } for item in items])
        logger.info(f"Saved {len(items)} products for {date_str}")

    # --------------------------------------------------------
    # キーワード集計保存
    # --------------------------------------------------------
    def save_word_stats(self, results: list[dict]):
        """キーワード集計を保存"""
        with self._conn() as conn:
            conn.executemany("""
                INSERT INTO word_stats
                    (date, genre, keyword, count, avg_rank, final_score, classification, products_json)
                VALUES
                    (:date, :genre, :keyword, :count, :avg_rank, :final_score, :classification, :products_json)
            """, [{
                "date": r["date"],
                "genre": r["genre"],
                "keyword": r["keyword"],
                "count": r["count"],
                "avg_rank": r["avg_rank"],
                "final_score": r["final_score"],
                "classification": r["classification"],
                "products_json": json.dumps(r.get("products", []), ensure_ascii=False),
            } for r in results])
        logger.info(f"Saved {len(results)} word stats")

    # --------------------------------------------------------
    # 新規参入判定
    # --------------------------------------------------------
    def get_first_appearance(self, keyword: str, genre: str) -> str | None:
        """キーワードが初めてランクインした日付を返す"""
        with self._conn() as conn:
            row = conn.execute("""
                SELECT MIN(date) as first_date FROM word_stats
                WHERE keyword = ? AND genre = ?
            """, (keyword, genre)).fetchone()
            return row["first_date"] if row else None

    def is_new_entrant(self, keyword: str, genre: str, date_str: str, days: int = 7) -> bool:
        """直近N日以内に初登場したキーワードかどうか"""
        first_date = self.get_first_appearance(keyword, genre)
        if not first_date:
            return True
        threshold = (datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=days)).strftime("%Y-%m-%d")
        return first_date >= threshold

    # --------------------------------------------------------
    # 定番商品判定（30日以上連続ランクイン）
    # --------------------------------------------------------
    def get_consecutive_days(self, keyword: str, genre: str, date_str: str) -> int:
        """直近から何日連続でランクインしているか"""
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT DISTINCT date FROM word_stats
                WHERE keyword = ? AND genre = ?
                ORDER BY date DESC
                LIMIT 60
            """, (keyword, genre)).fetchall()

        if not rows:
            return 0

        dates = [r["date"] for r in rows]
        count = 0
        current = datetime.strptime(date_str, "%Y-%m-%d")

        for d_str in dates:
            d = datetime.strptime(d_str, "%Y-%m-%d")
            diff = (current - d).days
            if diff == count:
                count += 1
            else:
                break
        return count

    # --------------------------------------------------------
    # 過去データ取得（季節分析用）
    # --------------------------------------------------------
    def get_word_history(self, keyword: str, genre: str = None, limit: int = 365) -> list[dict]:
        """キーワードの過去データを取得"""
        with self._conn() as conn:
            if genre:
                rows = conn.execute("""
                    SELECT date, genre, count, avg_rank, final_score, classification
                    FROM word_stats WHERE keyword = ? AND genre = ?
                    ORDER BY date DESC LIMIT ?
                """, (keyword, genre, limit)).fetchall()
            else:
                rows = conn.execute("""
                    SELECT date, genre, count, avg_rank, final_score, classification
                    FROM word_stats WHERE keyword = ?
                    ORDER BY date DESC LIMIT ?
                """, (keyword, limit)).fetchall()
        return [dict(r) for r in rows]

    def get_top_words_for_period(self, genre: str, start_date: str, end_date: str,
                                  classification: str = None, limit: int = 50) -> list[dict]:
        """指定期間の上位キーワードを取得（季節分析用）"""
        with self._conn() as conn:
            if classification:
                rows = conn.execute("""
                    SELECT keyword, AVG(final_score) as avg_score, SUM(count) as total_count,
                           COUNT(DISTINCT date) as days_appeared
                    FROM word_stats
                    WHERE genre = ? AND date BETWEEN ? AND ? AND classification = ?
                    GROUP BY keyword
                    ORDER BY avg_score DESC
                    LIMIT ?
                """, (genre, start_date, end_date, classification, limit)).fetchall()
            else:
                rows = conn.execute("""
                    SELECT keyword, AVG(final_score) as avg_score, SUM(count) as total_count,
                           COUNT(DISTINCT date) as days_appeared
                    FROM word_stats
                    WHERE genre = ? AND date BETWEEN ? AND ?
                    GROUP BY keyword
                    ORDER BY avg_score DESC
                    LIMIT ?
                """, (genre, start_date, end_date, limit)).fetchall()
        return [dict(r) for r in rows]

    # --------------------------------------------------------
    # Keepaローテーション管理
    # --------------------------------------------------------
    def get_keepa_rotation(self) -> dict:
        """各ジャンルの最終Keepa実行日を取得"""
        with self._conn() as conn:
            rows = conn.execute("SELECT * FROM keepa_rotation").fetchall()
        return {r["genre_name"]: dict(r) for r in rows}

    def update_keepa_rotation(self, genre_name: str, run_date: str):
        """Keepa実行日を更新"""
        with self._conn() as conn:
            conn.execute("""
                INSERT INTO keepa_rotation (genre_name, last_run_date)
                VALUES (?, ?)
                ON CONFLICT(genre_name) DO UPDATE SET last_run_date = excluded.last_run_date
            """, (genre_name, run_date))
