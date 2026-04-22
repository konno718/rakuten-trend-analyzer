# 楽天トレンドワード収集システム セットアップガイド

## 必要なもの

| 必要なもの | 入手先 | 費用 |
|---|---|---|
| Python 3.10以上 | python.org | 無料 |
| 楽天APIキー | webservice.rakuten.co.jp | 無料 |
| Google Sheetsサービスアカウント | console.cloud.google.com | 無料 |
| Keepa APIキー | keepa.com/api | 月約2,000円 |
| Discord Webhook URL | Discord サーバー設定 | 無料 |
| Claude APIキー（任意） | console.anthropic.com | 従量課金 |

---

## STEP 1: Pythonパッケージのインストール

```bash
cd trend-analyzer
pip3 install -r requirements.txt
```

---

## STEP 2: 楽天APIキーの確認

1. https://webservice.rakuten.co.jp/ にログイン
2. 「アプリ一覧」からアプリIDを確認
3. `.env` ファイルに記載

---

## STEP 3: Google Sheetsサービスアカウントの設定

### 3-1. Google Cloud Consoleでプロジェクト作成
1. https://console.cloud.google.com/ にアクセス
2. 「新しいプロジェクト」を作成（名前は任意）
3. 「APIとサービス」→「ライブラリ」で以下を有効化:
   - `Google Sheets API`
   - `Google Drive API`

### 3-2. サービスアカウントの作成
1. 「APIとサービス」→「認証情報」→「認証情報を作成」→「サービスアカウント」
2. サービスアカウント名は任意（例: `trend-analyzer`）
3. 作成後、「キー」タブ→「鍵を追加」→「JSONキー」をダウンロード
4. ダウンロードしたJSONファイルを `trend-analyzer/credentials.json` として保存

### 3-3. Googleスプレッドシートの準備
1. Google Sheetsで新しいスプレッドシートを作成
2. URLの `/d/` の後ろの文字列がシートID（例: `1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms`）
3. スプレッドシートを右クリック → 「共有」→ サービスアカウントのメールアドレスを「編集者」として追加
   ※ サービスアカウントのメールは `credentials.json` 内の `client_email` を参照

---

## STEP 4: .envファイルの設定

```bash
cp .env.example .env
```

`.env` を開いて以下を設定:

```
RAKUTEN_APP_ID=（楽天のアプリID）
KEEPA_API_KEY=（KeepaのAPIキー）
GOOGLE_SHEET_ID=（スプレッドシートのID）
GOOGLE_CREDENTIALS_PATH=./credentials.json
DISCORD_WEBHOOK_URL=（Discord Webhook URL）
CLAUDE_API_KEY=（Claude APIキー、任意）
```

---

## STEP 5: ジャンル設定（スプレッドシート）

スクリプトを一度実行すると「設定」シートが自動作成されます。

```bash
python3 main.py --test
```

作成された「設定」シートに楽天ランキングURLを貼り付けてください:

| A: ジャンル名 | B: 楽天ランキングURL | C: 有効 | D: KeepaカテゴリID |
|---|---|---|---|
| インテリア雑貨 | https://ranking.rakuten.co.jp/daily/100533/ | ○ | |
| ペット用品 | https://ranking.rakuten.co.jp/daily/101213/ | ○ | 2619525011 |

### 楽天ランキングURLの調べ方
1. https://ranking.rakuten.co.jp/ にアクセス
2. 見たいカテゴリを選択
3. そのページのURLをコピー（例: `https://ranking.rakuten.co.jp/daily/100533/`）

### KeepaカテゴリIDの調べ方
1. https://keepa.com/ にアクセスしてAmazonのカテゴリを確認
2. またはKeepa APIの `/category` エンドポイントで検索

---

## STEP 6: テスト実行

```bash
# テストモード（書き込みなし、1ジャンルのみ）
python3 main.py --test

# 通常実行
python3 main.py

# AI分析付き実行（週1回推奨）
python3 main.py --ai
```

---

## STEP 7: 毎日自動実行（cron設定）

Macで毎日深夜2時に自動実行:

```bash
crontab -e
```

以下を追加（パスは環境に合わせて変更）:

```cron
# 毎日2:00 通常実行
0 2 * * * cd /path/to/trend-analyzer && /usr/bin/python3 main.py >> data/cron.log 2>&1

# 毎週月曜2:30 AI分析付き実行
30 2 * * 1 cd /path/to/trend-analyzer && /usr/bin/python3 main.py --ai >> data/cron.log 2>&1
```

※ `/path/to/trend-analyzer` は実際のパスに変更してください
※ MacがスリープするとCronが動かないため、システム環境設定でスリープを無効化か、
　 `pmset` コマンドで自動起動を設定してください

---

## スプレッドシートの構成

自動作成されるシート一覧:

| シート名 | 内容 |
|---|---|
| **設定** | ジャンルURLを入力するシート |
| **ワード集計** | 日次のキーワード × スコア一覧 |
| **商品一覧** | キーワードに紐づく商品URL × 順位 |
| **除外候補** | 自動検出した汎用ワード（人間レビュー用）|
| **サマリー** | AI分析結果 |

---

## よくある質問

**Q: Keepa APIキーはどこで確認する？**
A: keepa.com にログイン → API キーの管理ページで確認

**Q: 楽天APIのレート制限は？**
A: 1秒1リクエスト程度が安全です。このスクリプトは0.5秒待機しています

**Q: スプレッドシートが重くなったら？**
A: 「ワード集計」シートは月1回程度、古いデータを別シートに移動してください

**Q: 除外候補シートのレビュー方法は？**
A: C列に「○」を入れると次回 `exclude_words.json` を手動更新する際の目安になります
   （現在は自動でJSONに反映されないため、手動更新が必要です）
