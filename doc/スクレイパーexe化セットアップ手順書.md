# スクレイパーexe化セットアップ手順書

## 概要

このドキュメントでは、在庫管理スクレイピングシステムをexe化し、Pythonがインストールされていない端末でも利用できるようにする手順を説明します。

## 目次

1. [exe化の準備（開発者側）](#exe化の準備開発者側)
2. [exe化の実行](#exe化の実行)
3. [配布パッケージの作成](#配布パッケージの作成)
4. [他端末でのセットアップ手順（利用者側）](#他端末でのセットアップ手順利用者側)
5. [トラブルシューティング](#トラブルシューティング)

---

## exe化の準備（開発者側）

### 1. 必要なツールのインストール

exe化にはPyInstallerを使用します。以下のコマンドでインストールしてください。

```bash
cd inventory_scraper
pip install pyinstaller
```

### 2. 依存関係の確認

`requirements.txt`に記載されているすべてのパッケージがインストールされていることを確認してください。

```bash
pip install -r requirements.txt
```

### 3. 動作確認

exe化する前に、Pythonスクリプトが正常に動作することを確認してください。

```bash
python main.py
```

---

## exe化の実行

### 方法1: 基本的なexe化（推奨）

以下のコマンドでexeファイルを生成します。

```bash
cd inventory_scraper
pyinstaller --onefile --name=InventoryScraper --icon=NONE main.py
```

**オプション説明:**
- `--onefile`: 単一のexeファイルにまとめる
- `--name=InventoryScraper`: 生成されるexeファイルの名前を指定
- `--icon=NONE`: アイコンなし（必要に応じて`.ico`ファイルを指定可能）

### 方法2: 詳細設定でのexe化

より詳細な設定が必要な場合は、以下のコマンドを使用してください。

```bash
pyinstaller --onefile ^
  --name=InventoryScraper ^
  --add-data "config;config" ^
  --hidden-import=selenium ^
  --hidden-import=pandas ^
  --hidden-import=webdriver_manager ^
  --hidden-import=dotenv ^
  --hidden-import=requests ^
  --collect-all=selenium ^
  --collect-all=webdriver_manager ^
  main.py
```

**オプション説明:**
- `--add-data "config;config"`: 設定ファイル（`config/scraper_config.json`）をexeに含める
- `--hidden-import`: 必要なモジュールを明示的に指定
- `--collect-all`: 指定したパッケージのすべてのサブモジュールを含める

### 方法3: specファイルを使用したexe化（推奨）

より細かい制御が必要な場合は、specファイルを作成して使用します。

#### 3-1. specファイルの生成

```bash
pyinstaller --onefile --name=InventoryScraper main.py
```

これにより`InventoryScraper.spec`ファイルが生成されます。

#### 3-2. specファイルの編集

生成された`InventoryScraper.spec`を編集して、必要な設定を追加します。

```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config/scraper_config.json', 'config'),
    ],
    hiddenimports=[
        'selenium',
        'pandas',
        'webdriver_manager',
        'dotenv',
        'requests',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='InventoryScraper',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # コンソールウィンドウを表示
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
```

#### 3-3. specファイルを使用してexe化

```bash
pyinstaller InventoryScraper.spec
```

### 4. exeファイルの生成場所

exeファイルは以下の場所に生成されます。

```
inventory_scraper/
└── dist/
    └── InventoryScraper.exe
```

---

## 配布パッケージの作成

### 1. 配布用フォルダの構成

配布用のフォルダを作成し、以下のファイルを含めます。

```
InventoryScraper_Package/
├── InventoryScraper.exe          # メインのexeファイル
├── config/                        # 設定ファイルフォルダ
│   └── scraper_config.json        # スクレイパー設定ファイル
├── .env.example                   # 環境変数設定のサンプル
├── README_セットアップ手順.txt    # セットアップ手順書（このドキュメントの利用者向け部分）
└── セットアップ手順書.md          # 詳細なセットアップ手順
```

### 2. 必要なファイルのコピー

```bash
# 配布用フォルダを作成
mkdir InventoryScraper_Package
mkdir InventoryScraper_Package\config

# exeファイルをコピー
copy dist\InventoryScraper.exe InventoryScraper_Package\

# 設定ファイルをコピー
xcopy /E /I config InventoryScraper_Package\config

# .env.exampleをコピー（存在する場合）
copy .env.example InventoryScraper_Package\
```

### 3. セットアップ手順書の作成

利用者向けのセットアップ手順書を作成します（後述の「他端末でのセットアップ手順」を参照）。

---

## 他端末でのセットアップ手順（利用者側）

### 前提条件

- **Windows 10以上**がインストールされていること
- **Google Chrome**がインストールされていること
- **インターネット接続**が利用可能であること
- **Googleアカウント**を持っていること（スプレッドシートへのアクセス権限が必要）

### ステップ1: 配布パッケージの展開

1. 配布された`InventoryScraper_Package`フォルダを任意の場所に展開します
   - 例: `C:\Program Files\InventoryScraper\` または `C:\Users\[ユーザー名]\InventoryScraper\`

2. フォルダ内のファイルを確認します
   - `InventoryScraper.exe`が存在することを確認
   - `config`フォルダが存在することを確認

### ステップ2: 環境変数設定ファイル（.env）の作成

1. 配布パッケージのフォルダ内に`.env`ファイルを作成します
   - メモ帳などのテキストエディタで新規ファイルを作成
   - ファイル名を`.env`として保存（拡張子なし）

2. `.env.example`が存在する場合は、それをコピーして`.env`として保存してください

3. `.env`ファイルに以下の内容を記述します

```ini
# スプレッドシート情報
SPREADSHEET_ID=あなたのスプレッドシートID
SHEET_GID=0

# Google Apps Script WebアプリURL（必須）
GAS_WEB_APP_URL=https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec

# ローカルChrome設定
CHROME_PROFILE_PATH=C:\Users\[ユーザー名]\AppData\Local\Google\Chrome\User Data
CHROME_PROFILE_NAME=Default

# ダウンロードフォルダ（Windowsの場合）
DOWNLOAD_FOLDER=C:\Users\[ユーザー名]\Downloads

# ログ設定
LOG_LEVEL=INFO
```

#### 各設定値の取得方法

##### SPREADSHEET_ID（スプレッドシートID）

1. 対象のGoogleスプレッドシートを開く
2. URLを確認
   - 例: `https://docs.google.com/spreadsheets/d/1ABC123xyz789/edit`
   - `SPREADSHEET_ID`は`/d/`と`/edit`の間の文字列（`1ABC123xyz789`）

##### SHEET_GID（シートGID）

1. スプレッドシートで「在庫管理」シートを開く
2. URLを確認
   - 例: `https://docs.google.com/spreadsheets/d/1ABC123xyz789/edit#gid=0`
   - `SHEET_GID`は`gid=`の後の数字（通常は`0`）

##### GAS_WEB_APP_URL（GAS WebアプリURL）

- GAS側でWebアプリとしてデプロイした際に取得できるURL
- 開発者から提供されるURLを使用してください

**GAS WebアプリURLの例:**
```
GAS_WEB_APP_URL=https://script.google.com/macros/s/AKfycbz1234567890abcdef/exec
```

##### CHROME_PROFILE_PATH（Chromeプロファイルパス）

**Windowsの場合:**
```
C:\Users\[ユーザー名]\AppData\Local\Google\Chrome\User Data
```

**確認方法:**
1. エクスプローラーを開く
2. アドレスバーに`%LOCALAPPDATA%\Google\Chrome\User Data`と入力してEnter
3. 表示されたパスを`.env`に設定

##### CHROME_PROFILE_NAME

通常は`Default`のままで問題ありません。複数のChromeプロファイルを使用している場合は、使用したいプロファイル名を指定してください。

##### DOWNLOAD_FOLDER（ダウンロードフォルダ）

通常は`C:\Users\[ユーザー名]\Downloads`で問題ありません。カスタムのダウンロードフォルダを使用している場合は、そのパスを指定してください。

### ステップ3: Google Chromeの設定

1. **Google Chromeを起動**して、Googleアカウントにログインしていることを確認します
   - スプレッドシートにアクセスできるアカウントでログインしてください

2. **Chromeのバージョンを確認**します
   - Chromeのメニュー（三点リーダー）→「ヘルプ」→「Google Chromeについて」
   - ChromeDriverは自動的にダウンロードされますが、Chromeのバージョンが古い場合は更新してください

### ステップ4: 初回実行と動作確認

1. **`InventoryScraper.exe`をダブルクリック**して実行します

2. **初回実行時の注意事項:**
   - 初回実行時、Windows Defenderなどのセキュリティソフトが警告を表示する場合があります
   - 「詳細情報」→「実行」をクリックして実行を許可してください
   - ChromeDriverが自動的にダウンロードされるため、初回実行には時間がかかる場合があります

3. **実行結果の確認:**
   - コンソールウィンドウが表示され、処理の進行状況が表示されます
   - エラーが発生した場合は、エラーメッセージを確認してください
   - ログファイルは`logs`フォルダに保存されます（自動的に作成されます）

### ステップ5: ショートカットの作成（オプション）

デスクトップにショートカットを作成すると便利です。

1. `InventoryScraper.exe`を右クリック
2. 「ショートカットの作成」を選択
3. 作成されたショートカットをデスクトップに移動

---

## トラブルシューティング

### 問題1: exeファイルが起動しない

**原因と対処法:**
- **Windows Defenderなどのセキュリティソフトがブロックしている**
  - セキュリティソフトの設定で、`InventoryScraper.exe`を例外に追加してください
- **必要なDLLが不足している**
  - Visual C++ Redistributableをインストールしてください
    - [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)

### 問題2: ChromeDriverのエラー

**エラーメッセージ例:**
```
selenium.common.exceptions.WebDriverException: Message: 'chromedriver' executable needs to be in PATH
```

**対処法:**
- ChromeDriverは自動的にダウンロードされますが、失敗する場合は以下を確認してください
  1. インターネット接続を確認
  2. Chromeのバージョンを最新に更新
  3. ファイアウォールがwebdriver-managerの通信をブロックしていないか確認

### 問題3: スプレッドシートにアクセスできない

**エラーメッセージ例:**
```
エラー: スプレッドシートへのアクセスに失敗しました
```

**対処法:**
1. `.env`ファイルの`SPREADSHEET_ID`が正しいか確認
2. Chromeでスプレッドシートにアクセスできることを確認
3. Chromeプロファイルパス（`CHROME_PROFILE_PATH`）が正しいか確認
4. Googleアカウントにログインしていることを確認

### 問題4: GAS Webアプリへの接続エラー

**エラーメッセージ例:**
```
エラー: GAS_WEB_APP_URLが設定されていません
```

**対処法:**
1. `.env`ファイルに`GAS_WEB_APP_URL`が正しく設定されているか確認
2. GAS Webアプリが正しくデプロイされているか確認
3. GAS Webアプリのアクセス権限が「全員（匿名ユーザーを含む）」に設定されているか確認

### 問題5: 設定ファイルが見つからない

**エラーメッセージ例:**
```
FileNotFoundError: config/scraper_config.json
```

**対処法:**
1. `config`フォルダがexeファイルと同じディレクトリに存在するか確認
2. `config/scraper_config.json`ファイルが存在するか確認
3. フォルダ構造が正しいか確認

### 問題8: ダウンロードが実行されない

**症状:**
- exeファイルを実行しても、スプレッドシートからのCSVダウンロードが実行されない
- エラーメッセージが表示されないが、処理が進まない

**原因:**
- exe実行時、作業ディレクトリやパス解決が正しく行われていない
- `.env`ファイルや`config`フォルダのパスが正しく解決されていない

**対処法:**
1. **exeファイルと同じディレクトリに`.env`ファイルが存在するか確認**
   - exeファイルを実行するディレクトリに`.env`ファイルを配置してください
   - `.env`ファイルの内容が正しく設定されているか確認してください

2. **exeファイルと同じディレクトリに`config`フォルダが存在するか確認**
   - `config/scraper_config.json`ファイルが存在するか確認してください

3. **フォルダ構造を確認**
   ```
   InventoryScraper.exe と同じディレクトリ/
   ├── InventoryScraper.exe
   ├── .env
   ├── config/
   │   └── scraper_config.json
   ├── data/          (自動生成)
   └── logs/          (自動生成)
   ```

4. **ログファイルを確認**
   - `logs`フォルダ内の最新のログファイルを確認して、エラーメッセージを確認してください

5. **再ビルド**
   - 最新のコードでexeファイルを再ビルドしてください（パス解決の修正が含まれています）

### 問題6: ログファイルが生成されない

**対処法:**
1. exeファイルを実行するディレクトリに書き込み権限があるか確認
2. `logs`フォルダが自動的に作成されることを確認
3. ウイルス対策ソフトがログファイルの作成をブロックしていないか確認

### 問題7: 実行が非常に遅い

**対処法:**
1. インターネット接続速度を確認
2. スクレイピング対象のURL数が多い場合は、処理に時間がかかります（正常な動作です）
3. ログファイルを確認して、エラーが発生していないか確認

---

## 注意事項

### セキュリティ

- `.env`ファイルには機密情報が含まれています。他のユーザーと共有しないでください
- exeファイルは信頼できるソースからのみ取得してください

### パフォーマンス

- 初回実行時はChromeDriverのダウンロードに時間がかかります
- 大量のURLを処理する場合は、実行時間が長くなる可能性があります

### 更新

- 新しいバージョンがリリースされた場合は、配布パッケージ全体を置き換えてください
- `.env`ファイルは新しいバージョンでもそのまま使用できます

---

## サポート

問題が解決しない場合は、以下の情報を含めて開発者に連絡してください。

1. エラーメッセージの全文
2. `logs`フォルダ内の最新のログファイル
3. `.env`ファイルの内容（機密情報はマスクして）
4. 実行環境の情報（Windowsのバージョン、Chromeのバージョンなど）

---

**作成日**: 2026-01-21  
**最終更新**: 2026-01-21
