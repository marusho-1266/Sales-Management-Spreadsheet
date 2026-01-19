"""
在庫管理スクレイピングシステム メインエントリーポイント
"""
import sys
import logging
from pathlib import Path
from datetime import datetime

# ログ設定
from src.config import LOGS_DIR, LOG_LEVEL

log_file = LOGS_DIR / f"scraper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL.upper()),
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

from src.browser import create_browser
from src.downloader import download_spreadsheet_csv
from src.scraper import scrape_urls
from src.uploader import save_result_csv
from src.spreadsheet_updater import update_spreadsheet_via_gas


def main():
    """
    メイン処理
    1. スプレッドシートからCSVをダウンロード
    2. 各ECサイトをスクレイピング
    3. 結果をCSVに保存
    4. GAS Webアプリ経由でスプレッドシートを更新
    """
    browser = None
    
    try:
        logger.info("=== 在庫管理スクレイピングシステム 開始 ===")
        
        # 1. ブラウザを初期化
        logger.info("ブラウザを初期化しています...")
        browser = create_browser()
        logger.info("ブラウザの初期化が完了しました")
        
        # 2. スプレッドシートからCSVをダウンロード
        logger.info("スプレッドシートからCSVをダウンロードしています...")
        df = download_spreadsheet_csv(browser)
        logger.info(f"CSVダウンロード完了: {len(df)}件のデータを取得しました")
        
        if len(df) == 0:
            logger.warning("スクレイピング対象のURLが見つかりませんでした")
            return
        
        # 3. スクレイピングを実行
        logger.info("スクレイピングを開始します...")
        result_df = scrape_urls(df, browser)
        logger.info(f"スクレイピング完了: {len(result_df)}件の結果を取得しました")
        
        # 4. 結果をCSVに保存
        logger.info("結果をCSVファイルに保存しています...")
        csv_path = save_result_csv(result_df)
        logger.info(f"CSVファイルを保存しました: {csv_path}")
        
        # 5. スプレッドシートに反映（GAS Webアプリ経由）
        logger.info("Google Apps Script Webアプリ経由でスプレッドシートを更新しています...")
        from src.config import GAS_WEB_APP_URL
        if not GAS_WEB_APP_URL:
            raise Exception("GAS_WEB_APP_URLが設定されていません。.envファイルにGAS_WEB_APP_URLを設定してください。")
        
        update_spreadsheet_via_gas(browser, csv_path, GAS_WEB_APP_URL)
        logger.info("スプレッドシートの更新が完了しました")
        
        logger.info("=== 在庫管理スクレイピングシステム 正常終了 ===")
        
    except Exception as e:
        logger.error(f"エラーが発生しました: {e}", exc_info=True)
        sys.exit(1)
        
    finally:
        # ブラウザを閉じる
        if browser:
            logger.info("ブラウザを閉じています...")
            browser.quit()
            logger.info("ブラウザを閉じました")


if __name__ == '__main__':
    main()
