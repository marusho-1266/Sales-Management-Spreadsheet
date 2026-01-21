#!/bin/bash
# 在庫管理スクレイピングシステム exe化スクリプト
# Linux/Mac用シェルスクリプト（Wine経由でWindows exeを生成する場合）

set -e

echo "========================================"
echo "在庫管理スクレイピングシステム exe化"
echo "========================================"
echo

# 現在のディレクトリを確認
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"
echo "作業ディレクトリ: $PWD"
echo

# PyInstallerがインストールされているか確認
if ! python -c "import PyInstaller" 2>/dev/null; then
    echo "[エラー] PyInstallerがインストールされていません"
    echo "以下のコマンドでインストールしてください:"
    echo "pip install pyinstaller"
    exit 1
fi

echo "[1/4] PyInstallerの確認完了"
echo

# 必要なパッケージがインストールされているか確認
echo "[2/4] 依存関係の確認中..."
if ! python -c "import selenium, pandas, webdriver_manager, dotenv, requests" 2>/dev/null; then
    echo "[警告] 一部のパッケージがインストールされていない可能性があります"
    echo "requirements.txtからインストールしてください:"
    echo "pip install -r requirements.txt"
    echo
    read -p "続行しますか? (Y/N): " continue
    if [[ ! "$continue" =~ ^[Yy]$ ]]; then
        exit 1
    fi
fi
echo "[2/4] 依存関係の確認完了"
echo

# 設定ファイルの存在確認
if [ ! -f "config/scraper_config.json" ]; then
    echo "[エラー] config/scraper_config.json が見つかりません"
    exit 1
fi
echo "[3/4] 設定ファイルの確認完了"
echo

# exe化の実行
echo "[4/4] exe化を実行中..."
echo "この処理には数分かかる場合があります..."
echo

pyinstaller InventoryScraper.spec

if [ $? -ne 0 ]; then
    echo
    echo "[エラー] exe化に失敗しました"
    exit 1
fi

echo
echo "========================================"
echo "exe化が完了しました！"
echo "========================================"
echo
echo "生成されたexeファイル:"
echo "  dist/InventoryScraper.exe"
echo
echo "配布用パッケージを作成する場合は、以下のファイルを含めてください:"
echo "  - dist/InventorySceaper.exe"
echo "  - config/scraper_config.json"
echo "  - .env.example (存在する場合)"
echo
