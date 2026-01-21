@echo off
REM 在庫管理スクレイピングシステム exe化スクリプト
REM Windows用バッチファイル

echo ========================================
echo 在庫管理スクレイピングシステム exe化
echo ========================================
echo.

REM 現在のディレクトリを確認
cd /d %~dp0
echo 作業ディレクトリ: %CD%
echo.

REM PyInstallerがインストールされているか確認
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [エラー] PyInstallerがインストールされていません
    echo 以下のコマンドでインストールしてください:
    echo pip install pyinstaller
    pause
    exit /b 1
)

echo [1/4] PyInstallerの確認完了
echo.

REM 必要なパッケージがインストールされているか確認
echo [2/4] 依存関係の確認中...
python -c "import selenium, pandas, webdriver_manager, dotenv, requests" 2>nul
if errorlevel 1 (
    echo [警告] 一部のパッケージがインストールされていない可能性があります
    echo requirements.txtからインストールしてください:
    echo pip install -r requirements.txt
    echo.
    echo 続行しますか? (Y/N)
    set /p continue=
    if /i not "%continue%"=="Y" (
        exit /b 1
    )
)
echo [2/4] 依存関係の確認完了
echo.

REM 設定ファイルの存在確認
if not exist "config\scraper_config.json" (
    echo [エラー] config\scraper_config.json が見つかりません
    pause
    exit /b 1
)
echo [3/4] 設定ファイルの確認完了
echo.

REM exe化の実行
echo [4/4] exe化を実行中...
echo この処理には数分かかる場合があります...
echo.

pyinstaller InventoryScraper.spec

if errorlevel 1 (
    echo.
    echo [エラー] exe化に失敗しました
    pause
    exit /b 1
)

echo.
echo ========================================
echo exe化が完了しました！
echo ========================================
echo.
echo 生成されたexeファイル:
echo   dist\InventoryScraper.exe
echo.
echo 配布用パッケージを作成する場合は、以下のファイルを含めてください:
echo   - dist\InventoryScraper.exe
echo   - config\scraper_config.json
echo   - .env.example (存在する場合)
echo.
pause
