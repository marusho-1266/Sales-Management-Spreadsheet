"""
設定管理モジュール
環境変数から設定を読み込む
"""
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# exe実行時と通常実行時でベースディレクトリを適切に取得
def get_base_dir():
    """
    exe実行時と通常実行時で適切なベースディレクトリを取得する
    
    Returns:
        Path: ベースディレクトリのPathオブジェクト
    """
    if getattr(sys, 'frozen', False):
        # exe実行時: sys.executableがexeファイルのパス
        # exeファイルのディレクトリをベースディレクトリとする
        base_dir = Path(sys.executable).parent.resolve()
    else:
        # 通常実行時: このファイルの親の親（inventory_scraper）をベースディレクトリとする
        base_dir = Path(__file__).parent.parent.resolve()
    return base_dir

# ベースディレクトリを取得
BASE_DIR = get_base_dir()

# .envファイルを読み込む
env_path = BASE_DIR / '.env'
load_dotenv(env_path)

# スプレッドシート情報
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID', '')
SHEET_GID = os.getenv('SHEET_GID', '0')  # 在庫管理シートのGID
SUPPLIER_SHEET_GID = os.getenv('SUPPLIER_SHEET_GID', '')  # 仕入れ元マスターシートのGID（空の場合は自動検出）

# Google Apps Script WebアプリURL（直接更新用）
GAS_WEB_APP_URL = os.getenv('GAS_WEB_APP_URL', '')

# ローカルChrome設定
CHROME_PROFILE_PATH = os.getenv('CHROME_PROFILE_PATH', '')
CHROME_PROFILE_NAME = os.getenv('CHROME_PROFILE_NAME', 'Default')

# Chrome User-Agent設定
# 環境変数 CHROME_USER_AGENT が設定されていない場合はデフォルト値を使用
DEFAULT_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
CHROME_USER_AGENT = os.getenv('CHROME_USER_AGENT', DEFAULT_USER_AGENT)

# ダウンロードフォルダ
DOWNLOAD_FOLDER = os.getenv('DOWNLOAD_FOLDER', str(Path.home() / 'Downloads'))

# ログ設定
LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')

# デバッグ設定
# デバッグモードを有効にする場合は 'true' または '1' を設定（デフォルト: 無効）
ENABLE_DEBUG_MODE = os.getenv('ENABLE_DEBUG_MODE', 'false').lower() in ('true', '1', 'yes')

# データ保存先
DATA_DIR = BASE_DIR / 'data'
LOGS_DIR = BASE_DIR / 'logs'

# ディレクトリが存在しない場合は作成
DATA_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)
