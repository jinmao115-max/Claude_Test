# Hotel Price Checker

毎朝7時に Booking.com と Agoda のホテル最安値をスクレイピングし、Gmail で通知するシステムです。

## 機能

- Booking.com / Agoda からホテル価格を自動取得（Playwright）
- 価格比較レポートを生成
- Gmail API でHTML形式のメール送信
- GitHub Actions で毎朝7時（JST）に自動実行
- GitHub Pages で管理画面を公開

## 管理画面

**[https://jinmao115-max.github.io/Claude_Test/](https://jinmao115-max.github.io/Claude_Test/)**

## セットアップ

### 1. 依存パッケージのインストール

```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. Gmail API 認証

1. [Google Cloud Console](https://console.cloud.google.com/) で OAuth2 クライアント ID を作成
2. `credentials.json` をプロジェクトルートに配置
3. 初回認証トークンを生成:

```bash
python generate_token.py
```

### 3. GitHub Secrets の設定

| Secret 名 | 内容 |
|-----------|------|
| `GMAIL_TOKEN_JSON` | `token.json` の内容をそのまま貼り付け |
| `GMAIL_SENDER` | 送信元Gmailアドレス |
| `NOTIFY_EMAIL` | 通知先メールアドレス |
| `BOOKING_URL` | （任意）検索URL |
| `AGODA_URL` | （任意）検索URL |

### 4. GitHub Pages の有効化

リポジトリの **Settings → Pages → Source** で `main` ブランチの `/docs` フォルダを選択。

## ディレクトリ構成

```
.
├── main.py                  # エントリーポイント
├── src/
│   ├── __init__.py
│   ├── scraper_booking.py   # Booking.com スクレイパー
│   ├── scraper_agoda.py     # Agoda スクレイパー
│   ├── compare_prices.py    # 価格比較・レポート生成
│   └── send_gmail.py        # Gmail 送信
├── docs/
│   └── index.html           # GitHub Pages 管理画面
├── .github/
│   └── workflows/
│       └── hotel_checker.yml
├── requirements.txt
└── README.md
```

## ローカル実行

```bash
python main.py
```

## ライセンス

MIT
