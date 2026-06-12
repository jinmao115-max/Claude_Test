# Hotel Price Checker 作成手順書

毎朝 Gmail にホテル最安値レポートを自動送信するシステムの構築・運用手順書です。

---

## 1. システム概要

```
毎朝 7:00（JST）
    ↓
GitHub Actions が自動起動
    ↓
Booking.com / Agoda をスクレイピング
    ↓
最安値を比較
    ↓
Gmail でレポートをメール送信
```

### 構成ファイル

| ファイル | 役割 |
|---|---|
| `main.py` | エントリーポイント。設定読み込み・比較・送信を統括 |
| `src/compare_prices.py` | 各サイトの価格を比較し最安値を返す |
| `src/scraper_booking.py` | Booking.com スクレイパー（Playwright） |
| `src/scraper_agoda.py` | Agoda スクレイパー（Playwright） |
| `src/send_gmail.py` | Gmail API でメール送信 |
| `config/search_conditions.json` | 検索するホテル・日程・条件の設定ファイル |
| `generate_token.py` | Gmail 認証トークン発行ユーティリティ（ローカル実行用） |
| `.github/workflows/daily_check.yml` | GitHub Actions ワークフロー定義 |
| `requirements.txt` | 依存ライブラリ一覧 |
| `.gitignore` | 認証情報ファイルの除外設定 |

---

## 2. 使用技術・ライブラリ

| ライブラリ | 用途 |
|---|---|
| `playwright` | ブラウザ自動操作（スクレイピング） |
| `google-auth` | Google OAuth2 認証 |
| `google-auth-oauthlib` | OAuth フロー処理 |
| `google-api-python-client` | Gmail API クライアント |

---

## 3. 初期セットアップ手順

### 3.1 Google Cloud Console でOAuth認証情報を作成

1. [https://console.cloud.google.com/](https://console.cloud.google.com/) にアクセス
2. プロジェクト作成（例: `hotel-price-checker-498923`）
3. 左メニュー **「APIとサービス」→「ライブラリ」** で **Gmail API** を有効化
4. **「APIとサービス」→「認証情報」→「＋認証情報を作成」→「OAuth クライアント ID」**
5. アプリケーションの種類: **デスクトップアプリ**
6. 作成後 **「JSONをダウンロード」** → `credentials.json` にリネーム

> ⚠️ `credentials.json` はローカルにのみ保管。Git にコミットしない（`.gitignore` で除外済み）

---

### 3.2 Gmail 送信用トークンの発行

`credentials.json` を `Claude_Test` フォルダに置いてから実行：

```powershell
cd C:\Users\ka94654\AppData\Local\Temp\Claude_Test
python generate_token.py
```

ブラウザが自動で開く：

1. Gmail アカウント（`jinmao115@gmail.com`）を選択
2. 「このアプリは確認されていません」→ **「詳細」→「安全でないページに移動」**
3. Gmail 送信権限 → **「許可」**
4. ブラウザに **「The authentication flow has completed.」** と表示されたら成功
5. 同フォルダに `token.json` が生成される

---

### 3.3 GitHub Secrets への登録

[https://github.com/jinmao115-max/Claude_Test/settings/secrets/actions](https://github.com/jinmao115-max/Claude_Test/settings/secrets/actions) を開く。

以下の2つの Secret を登録：

| Secret 名 | 設定する値 |
|---|---|
| `GMAIL_TOKEN` | `token.json` の中身（テキスト全体）をペースト |
| `GMAIL_CREDENTIALS` | `credentials.json` の中身（テキスト全体）をペースト |

```powershell
# 中身の確認コマンド（コピー用）
Get-Content "C:\Users\ka94654\AppData\Local\Temp\Claude_Test\token.json"
Get-Content "C:\Users\ka94654\AppData\Local\Temp\Claude_Test\credentials.json"
```

---

### 3.4 検索条件の設定

`config/search_conditions.json` を編集してホテル・日程を設定する。

```json
{
  "settings": {
    "notify_email": "jinmao115@gmail.com"
  },
  "searches": [
    {
      "hotel_name": "検索するホテル名",
      "location": "大阪",
      "checkin": "2026-08-01",
      "checkout": "2026-08-02",
      "guests": 2,
      "rooms": 1,
      "bed_type": "any",
      "breakfast": "any",
      "free_cancellation": false
    }
  ]
}
```

---

## 4. GitHub Actions ワークフロー

### スケジュール

```yaml
# 毎日 22:00 UTC = 日本時間 翌朝 7:00
cron: '0 22 * * *'
```

### 認証情報の受け渡し方法

```yaml
env:
  GMAIL_CREDENTIALS: ${{ secrets.GMAIL_CREDENTIALS }}
  GMAIL_TOKEN: ${{ secrets.GMAIL_TOKEN }}
```

ファイルではなく **環境変数経由** で認証情報を渡す設計。  
`send_gmail.py` が `os.environ.get("GMAIL_TOKEN")` で読み取る。

### 手動実行（テスト）

[https://github.com/jinmao115-max/Claude_Test/actions/workflows/daily_check.yml](https://github.com/jinmao115-max/Claude_Test/actions/workflows/daily_check.yml)

→ **「Run workflow」→「Run workflow」** をクリック

---

## 5. セキュリティ対応記録（2026年6月）

### 発生した問題

`credentials.json` と `token.json`（Google OAuth2 の実認証情報）を  
誤って **公開 GitHub リポジトリにコミット**してしまった。

含まれていたリスク情報：
- `client_secret`（Google OAuth2 クライアントシークレット）
- `refresh_token`（Gmail 送信権限・有効期限なし）

### 実施した対応

| 手順 | 内容 | 方法 |
|---|---|---|
| ① | Google Cloud Console でクライアントシークレットを**無効化・再発行** | Google Cloud Console |
| ② | `git filter-repo` で全コミット履歴から2ファイルを**完全削除** | `python -m git_filter_repo --path credentials.json --path token.json --invert-paths --force` |
| ③ | GitHub へ **force push** で書き換えた履歴を反映 | `git push origin main --force` |
| ④ | **`.gitignore`** を追加して再発防止 | `credentials.json`, `token.json` 等を除外 |
| ⑤ | 新しい `credentials.json` で `generate_token.py` を実行し `token.json` を再発行 | ローカル実行 |
| ⑥ | GitHub Secrets（`GMAIL_TOKEN` / `GMAIL_CREDENTIALS`）を新しい認証情報で**更新** | GitHub Settings |

### 教訓・再発防止策

- **認証情報は最初から `.gitignore` に追加**してからプロジェクトを開始する
- GitHub Actions では認証情報を **Secrets 経由** で渡し、ファイルをリポジトリに置かない
- `.gitignore` に必ず含めるパターン:

```gitignore
credentials.json
token.json
*.key
*.pem
.env
.env.*
secrets.json
service_account.json
```

---

## 6. トークン再発行手順（定期メンテナンス）

Gmail トークンが無効になった場合（エラーメール・Actions 失敗時）：

```powershell
# 1. 作業フォルダへ移動
cd C:\Users\ka94654\AppData\Local\Temp\Claude_Test

# 2. credentials.json をここに置く（Google Cloud Console からダウンロード）

# 3. トークン再発行
python generate_token.py
# ブラウザで認証 → token.json が生成される

# 4. GitHub Secrets を更新
#    GMAIL_TOKEN に token.json の中身をペースト
#    → https://github.com/jinmao115-max/Claude_Test/settings/secrets/actions

# 5. 手動実行でテスト
#    → https://github.com/jinmao115-max/Claude_Test/actions
```

---

## 7. トラブルシューティング

| 症状 | 原因 | 対処 |
|---|---|---|
| Actions が `FileNotFoundError: credentials.json` で失敗 | ファイルをリポジトリに置こうとしている | Secrets 経由に変更。`send_gmail.py` は `GMAIL_TOKEN` 環境変数を参照 |
| Actions が `invalid_grant` で失敗 | refresh_token が無効化されている | §6 の手順でトークン再発行 |
| `generate_token.py` が `No such file or directory` | 実行フォルダが違う | `cd C:\Users\ka94654\AppData\Local\Temp\Claude_Test` してから実行 |
| ブラウザに「確認されていません」と表示 | OAuth 同意画面が未設定 | 「詳細」→「安全でないページに移動」→「許可」 |
| メールが届かない | Secrets の値が古い | `GMAIL_TOKEN` / `GMAIL_CREDENTIALS` を最新値で Update |

---

## 8. リポジトリ構成

```
Claude_Test/
├── .github/
│   └── workflows/
│       └── daily_check.yml       # GitHub Actions（毎朝実行）
├── 27卒就職比較表/                # 就職比較表（別プロジェクト）
├── config/
│   └── search_conditions.json    # 検索条件（編集して使う）
├── docs/
│   └── index.html                # 管理ダッシュボード（GitHub Pages）
├── src/
│   ├── compare_prices.py
│   ├── scraper_booking.py
│   ├── scraper_agoda.py
│   └── send_gmail.py
├── .gitignore                    # credentials.json 等を除外
├── generate_token.py             # トークン発行ユーティリティ
├── main.py                       # エントリーポイント
├── requirements.txt
└── Hotel_Price_Checker_作成手順書.md  # 本ファイル
```

---

*最終更新: 2026年6月（セキュリティ対応・GitHub Secrets 移行を反映）*
