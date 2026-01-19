# 在庫管理スクレイピングシステム

ローカルPC上のPythonとGoogleスプレッドシートを連携させた在庫管理ツールです。
Google Cloud Platform (GCP) やGASのウェブアプリ公開機能を利用せず、ローカルブラウザの自動操作（Selenium）のみでデータの読み書きを完結させる「ハイブリッド・ローカル実行モデル」を採用しています。

## 機能

1. **データ読み込み**: スプレッドシートの「在庫管理」シートをCSVとして取得
2. **スクレイピング**: 各ECサイト（Amazon, Mercari, Yahoo!等）をスクレイピングし、価格・在庫情報を取得
3. **データ書き込み**: 結果CSVをGoogleフォーム経由で送信
4. **自動更新**: GASの「フォーム送信時トリガー」が発火し、CSV内容をスプレッドシートに反映

## セットアップ

### 1. 必要な環境

- Python 3.10以上
- Google Chrome（ローカルにインストール済み）
- Googleアカウント（スプレッドシートとフォームへのアクセス権限）

### 2. インストール

```bash
cd inventory_scraper
pip install -r requirements.txt
```

### 3. 環境変数の設定

`.env.example`を`.env`にコピーして、実際の値を設定してください。

```bash
cp .env.example .env
```

`.env`ファイルを編集：

```ini
# スプレッドシート情報
SPREADSHEET_ID=あなたのスプレッドシートID
SHEET_GID=0

# Googleフォーム情報
GOOGLE_FORM_URL=https://docs.google.com/forms/d/e/xxxxx/viewform

# ローカルChrome設定
CHROME_PROFILE_PATH=C:\Users\Username\AppData\Local\Google\Chrome\User Data
CHROME_PROFILE_NAME=Default

# ダウンロードフォルダ（Windowsの場合）
DOWNLOAD_FOLDER=C:\Users\Username\Downloads

# ログ設定
LOG_LEVEL=INFO
```

### 4. Googleフォームの作成

1. Googleフォームを新規作成
2. 「ファイルアップロード」タイプの質問を追加
3. フォームURLを`.env`の`GOOGLE_FORM_URL`に設定

### 5. GAS側の設定

1. スプレッドシートの「拡張機能」→「Apps Script」を開く
2. `WebScrapingFormHandler.gs`を追加
3. `setupWebScrapingFormTrigger()`関数を実行してトリガーを設定
   - または手動でトリガーを設定：
     - 関数: `onFormSubmit`
     - イベント: フォームから、送信時

## 使用方法

### 実行

```bash
python main.py
```

### 実行フロー

1. スプレッドシートの「在庫管理」シートからCSVをダウンロード
2. 「仕入れ元URL」列が空でない行を抽出
3. 各URLに対してスクレイピングを実行
4. 結果をCSVファイルに保存
5. GoogleフォームにCSVファイルをアップロード
6. GASのトリガーが自動実行され、スプレッドシートを更新

## ディレクトリ構成

```
inventory_scraper/
├── data/                  # CSVファイル一時保存用
├── logs/                  # 実行ログ
├── src/
│   ├── __init__.py
│   ├── config.py          # 定数・設定読み込み
│   ├── browser.py         # Seleniumドライバー初期化・設定
│   ├── downloader.py      # スプレッドシートDL処理
│   ├── scraper.py         # スクレイピングロジック（Strategyパターン）
│   └── uploader.py        # Googleフォームへのアップロード処理
├── .env                   # 環境変数（URL, パス等）
├── main.py                # エントリーポイント
└── requirements.txt
```

## 対応ECサイト

- Amazon（amazon.co.jp, amazon.com）
- メルカリ（mercari.com, mercari.jp）
- Yahoo!ショッピング（shopping.yahoo.co.jp）

## エラーハンドリング

- **DL失敗**: タイムアウト時はリトライを1回行う
- **スクレイピング失敗**: 
  - ページが存在しない(404)場合 → 在庫ステータス「売り切れ」、仕入れ価格「0」
  - 要素が見つからない場合 → 仕入れ価格「-1」、在庫ステータス「不明」
- **アップロード失敗**: 例外をキャッチし、ローカルにCSVを残したままログを出力

## ログ

実行ログは`logs/`ディレクトリに保存されます。
ファイル名は`scraper_YYYYMMDD_HHMMSS.log`形式です。

## 注意事項

- Chromeプロファイルを使用するため、Googleアカウントにログイン済みの状態で実行してください
- スクレイピング時は各アクセスごとに3-7秒のランダムな待機時間を入れています
- 大量のURLを処理する場合は、実行時間が長くなる可能性があります

## トラブルシューティング

### ChromeDriverのエラー

`webdriver-manager`が自動的にChromeDriverをダウンロードしますが、エラーが発生する場合は手動でインストールしてください。

### CSVダウンロードがタイムアウトする

`.env`の`DOWNLOAD_FOLDER`が正しく設定されているか確認してください。

### フォーム送信が失敗する

Googleフォームのセレクタが変更されている可能性があります。`uploader.py`のセレクタを確認してください。
