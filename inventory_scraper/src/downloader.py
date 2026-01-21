"""
スプレッドシートダウンロードモジュール
スプレッドシートの「在庫管理」シートをCSVとして取得する
"""
import time
import pandas as pd
from pathlib import Path
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from .config import SPREADSHEET_ID, SHEET_GID, DATA_DIR


def download_spreadsheet_csv(browser, retry_count=1):
    """
    スプレッドシートの「在庫管理」シートをCSVとして取得する
    
    Args:
        browser: Selenium WebDriverインスタンス
        retry_count: リトライ回数（デフォルト: 1）
    
    Returns:
        pd.DataFrame: スプレッドシートのデータをDataFrame形式で返す
    
    Raises:
        Exception: ダウンロードに失敗した場合
    """
    # エクスポートURLを生成
    export_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv&gid={SHEET_GID}"
    
    # ダウンロード前のファイル一覧を取得
    data_dir = Path(DATA_DIR).resolve()
    before_files = set(data_dir.glob('*.csv'))
    
    # ダウンロードフォルダの最新ファイルを記録
    try:
        # URLにアクセス（自動ダウンロード発生）
        browser.get(export_url)
        
        # ダウンロード完了を待機（最大60秒）
        max_wait_time = 60
        wait_time = 0
        download_complete = False
        latest_file = None
        stable_count = 0  # ファイルサイズが安定するまでのカウント
        
        while wait_time < max_wait_time:
            time.sleep(1)
            wait_time += 1
            
            # ダウンロード後のファイル一覧を取得
            after_files = set(data_dir.glob('*.csv'))
            new_files = after_files - before_files
            
            if new_files:
                # 最新のファイルを取得（更新日時でソート）
                current_latest = max(new_files, key=lambda f: f.stat().st_mtime)
                
                # ファイルサイズが安定しているか確認（2秒間同じサイズなら完了とみなす）
                try:
                    current_size = current_latest.stat().st_size
                    if latest_file is None or latest_file != current_latest:
                        latest_file = current_latest
                        stable_count = 0
                    elif latest_file.stat().st_size == current_size:
                        stable_count += 1
                        if stable_count >= 2:  # 2秒間サイズが変わらなければ完了
                            download_complete = True
                            break
                    else:
                        stable_count = 0  # サイズが変わったらリセット
                except (OSError, FileNotFoundError):
                    # ファイルがまだ書き込み中の可能性がある
                    continue
        
        if not download_complete or latest_file is None:
            if retry_count > 0:
                print(f"ダウンロードがタイムアウトしました。リトライします... (残り{retry_count}回)")
                return download_spreadsheet_csv(browser, retry_count - 1)
            else:
                raise Exception("CSVダウンロードがタイムアウトしました")
        
        # CSVファイルを読み込む（ファイルの書き込み完了を待つ）
        time.sleep(1)  # 念のため追加の待機
        df = pd.read_csv(latest_file, encoding='utf-8-sig')
        
        # 仕入れ元URL列が空でない行のみをフィルタリング
        supplier_url_col = '仕入れ元URL'
        if supplier_url_col in df.columns:
            df = df[df[supplier_url_col].notna() & (df[supplier_url_col] != '')]
        else:
            raise Exception(f"CSVに「{supplier_url_col}」列が見つかりません")
        
        print(f"CSVダウンロード完了: {len(df)}件のデータを取得しました")
        return df
        
    except Exception as e:
        raise Exception("CSVダウンロードに失敗しました") from e
