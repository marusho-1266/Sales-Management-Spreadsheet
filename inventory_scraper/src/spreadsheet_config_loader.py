"""
スプレッドシートからスクレイパー設定を読み込むモジュール
仕入れ元マスターシートから設定を読み込んでJSON形式に変換
"""
import json
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from .config import SPREADSHEET_ID, DATA_DIR, SUPPLIER_SHEET_GID


class SpreadsheetConfigLoader:
    """スプレッドシートからスクレイパー設定を読み込むクラス"""
    
    def __init__(self, browser=None):
        """
        Args:
            browser: Selenium WebDriverインスタンス（CSVダウンロード用）
        """
        self.browser = browser
        self.supplier_sheet_gid = None  # 仕入れ元マスターシートのGID（後で検出）
    
    def load_config_from_spreadsheet(self) -> Dict:
        """
        スプレッドシートの仕入れ元マスターシートから設定を読み込む
        
        Returns:
            Dict: JSON設定形式の辞書
        """
        if not self.browser:
            raise Exception("ブラウザインスタンスが必要です")
        
        # 仕入れ元マスターシートをCSVとしてダウンロード
        df = self._download_supplier_master_csv()
        
        # DataFrameをJSON設定形式に変換
        config = self._convert_dataframe_to_config(df)
        
        return config
    
    def _download_supplier_master_csv(self) -> pd.DataFrame:
        """
        仕入れ元マスターシートをCSVとしてダウンロード
        
        Returns:
            pd.DataFrame: 仕入れ元マスターシートのデータ
        """
        # 仕入れ元マスターシートのGIDを取得
        # .envで指定されている場合はそれを使用、なければ自動検出
        if SUPPLIER_SHEET_GID and SUPPLIER_SHEET_GID.strip():
            # .envでGIDが指定されている場合は直接使用
            print(f"仕入れ元マスターシートのGIDを.envから取得: {SUPPLIER_SHEET_GID}")
            return self._download_csv_by_gid(SUPPLIER_SHEET_GID)
        else:
            # GIDが指定されていない場合は自動検出（複数のシートを試行）
            print("仕入れ元マスターシートを検索中（GIDが.envで指定されていないため自動検出）...")
            return self._search_supplier_master_sheet()
    
    def _download_csv_by_gid(self, gid: str) -> pd.DataFrame:
        """
        指定されたGIDでCSVをダウンロード
        
        Args:
            gid: シートのGID
            
        Returns:
            pd.DataFrame: シートのデータ
        """
        data_dir = Path(DATA_DIR).resolve()
        before_files = set(data_dir.glob('*.csv'))
        before_files_mtime = {f: f.stat().st_mtime for f in before_files} if before_files else {}
        
        try:
            export_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv&gid={gid}"
            self.browser.get(export_url)
            
            # エラーページが表示されていないか確認
            import time
            time.sleep(2)  # ページ読み込み待機
            
            # エラーページの検出
            page_text = self.browser.page_source.lower()
            page_title = self.browser.title.lower()
            
            error_indicators = [
                '現在、ファイルを開くことができません',
                'ファイルを開くことができません',
                'cannot open the file',
                'unable to open the file',
                'アドレスを確認して、もう一度試してください',
                'check the address and try again'
            ]
            
            is_error_page = any(indicator in page_text or indicator in page_title for indicator in error_indicators)
            
            if is_error_page:
                raise Exception(f"GID {gid} のシートにアクセスできませんでした（エラーページが表示されました）")
            
            # CSVファイルがダウンロードされるのを待つ
            time.sleep(1)
            
            # ダウンロードされたCSVファイルを探す
            after_files = set(data_dir.glob('*.csv'))
            new_files = after_files - before_files
            
            if not new_files:
                for file in after_files:
                    current_mtime = file.stat().st_mtime
                    if file not in before_files_mtime or current_mtime > before_files_mtime[file]:
                        new_files.add(file)
            
            if new_files:
                latest_file = max(new_files, key=lambda f: f.stat().st_mtime)
                df = pd.read_csv(latest_file, encoding='utf-8-sig')
                
                # ヘッダーを確認して仕入れ元マスターシートかどうかを判定
                expected_headers = ['サイト名', 'URLパターン（カンマ区切り）', '価格セレクタ（カンマ区切り）']
                if all(header in df.columns for header in expected_headers):
                    print(f"✓ 仕入れ元マスターシートを読み込みました（GID: {gid}）")
                    self.supplier_sheet_gid = gid
                    
                    # CSVダウンロード完了後、ブラウザタブをクリーンアップ（スプレッドシートが開いたままになるのを防ぐ）
                    try:
                        self.browser.get('about:blank')
                    except:
                        pass
                    
                    return df
                else:
                    raise Exception(f"GID {gid} のシートは仕入れ元マスターシートではありません（期待されるヘッダーが見つかりませんでした）")
            else:
                raise Exception(f"GID {gid} のCSVファイルがダウンロードされませんでした")
        except Exception as e:
            # エラー発生時もブラウザタブをクリーンアップ
            try:
                self.browser.get('about:blank')
            except:
                pass
            raise Exception(f"GID {gid} のシートの読み込みに失敗しました: {e}")
    
    def _search_supplier_master_sheet(self) -> pd.DataFrame:
        """
        仕入れ元マスターシートを自動検出（複数のシートを試行）
        
        Returns:
            pd.DataFrame: 仕入れ元マスターシートのデータ
        """
        # GIDの検索範囲を拡大（0-20まで試行）
        possible_gids = [str(i) for i in range(21)]  # 0-20まで
        
        # ダウンロード前のファイル一覧を取得（最新ファイルの判定用）
        data_dir = Path(DATA_DIR).resolve()
        before_files = set(data_dir.glob('*.csv'))
        before_files_mtime = {f: f.stat().st_mtime for f in before_files} if before_files else {}
        
        for gid in possible_gids:
            try:
                print(f"  GID {gid} を試行中...")
                export_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv&gid={gid}"
                self.browser.get(export_url)
                
                # エラーページが表示されていないか確認
                import time
                time.sleep(2)  # ページ読み込み待機
                
                # エラーページの検出（日本語と英語の両方に対応）
                page_text = self.browser.page_source.lower()
                page_title = self.browser.title.lower()
                current_url = self.browser.current_url.lower()
                
                # エラーページの特徴を検出
                error_indicators = [
                    '現在、ファイルを開くことができません',
                    'ファイルを開くことができません',
                    'cannot open the file',
                    'unable to open the file',
                    'アドレスを確認して、もう一度試してください',
                    'check the address and try again'
                ]
                
                is_error_page = any(indicator in page_text or indicator in page_title for indicator in error_indicators)
                
                if is_error_page:
                    print(f"  × GID {gid}: エラーページが表示されました（シートが存在しない可能性があります）")
                    # エラーページを閉じて空白ページに移動
                    try:
                        self.browser.get('about:blank')
                    except:
                        pass
                    continue
                
                # CSVファイルがダウンロードされるのを待つ
                time.sleep(1)  # 追加の待機時間
                
                # ダウンロードされたCSVファイルを探す
                after_files = set(data_dir.glob('*.csv'))
                new_files = after_files - before_files
                
                # 新しいファイルが見つからない場合、更新されたファイルを探す
                if not new_files:
                    # 既存ファイルの更新時刻を確認
                    for file in after_files:
                        current_mtime = file.stat().st_mtime
                        if file not in before_files_mtime or current_mtime > before_files_mtime[file]:
                            new_files.add(file)
                
                if new_files:
                    # 最新のファイルを取得（更新日時でソート）
                    latest_file = max(new_files, key=lambda f: f.stat().st_mtime)
                    
                    # ヘッダーを確認して仕入れ元マスターシートかどうかを判定
                    try:
                        df = pd.read_csv(latest_file, encoding='utf-8-sig', nrows=1)
                        
                        # 期待されるヘッダーを確認
                        expected_headers = ['サイト名', 'URLパターン（カンマ区切り）', '価格セレクタ（カンマ区切り）']
                        if all(header in df.columns for header in expected_headers):
                            print(f"  ✓ 仕入れ元マスターシートを発見（GID: {gid}）")
                            # 全体を読み込む
                            df = pd.read_csv(latest_file, encoding='utf-8-sig')
                            self.supplier_sheet_gid = gid
                            
                            # CSVダウンロード完了後、ブラウザタブをクリーンアップ（スプレッドシートが開いたままになるのを防ぐ）
                            try:
                                self.browser.get('about:blank')
                            except:
                                pass
                            
                            return df
                        else:
                            print(f"  × GID {gid}: ヘッダーが一致しません（見つかったヘッダー: {list(df.columns)[:5]}...）")
                    except Exception as e:
                        print(f"  × GID {gid}: CSV読み込みエラー: {e}")
                        continue
                else:
                    print(f"  × GID {gid}: CSVファイルがダウンロードされませんでした")
            except Exception as e:
                print(f"  × GID {gid}: エラー: {e}")
                continue
        
        # 見つからない場合はエラー
        raise Exception(f"仕入れ元マスターシートが見つかりませんでした（GID 0-20を検索しました）。.envファイルにSUPPLIER_SHEET_GIDを指定してください。")
    
    def _convert_dataframe_to_config(self, df: pd.DataFrame) -> Dict:
        """
        DataFrameをJSON設定形式に変換
        
        Args:
            df: 仕入れ元マスターシートのDataFrame
            
        Returns:
            Dict: JSON設定形式の辞書
        """
        sites = []
        
        # 列名のマッピング（旧形式との互換性のため）
        column_mapping = {
            'サイト名': 'name',
            'URLパターン（カンマ区切り）': 'url_patterns',
            '価格セレクタ（カンマ区切り）': 'price_selectors',
            '価格除外セレクタ（カンマ区切り）': 'price_exclude_selectors',
            '在庫セレクタ（カンマ区切り）': 'stock_selectors',
            '在庫ありキーワード（カンマ区切り）': 'in_stock_keywords',
            '売り切れキーワード（カンマ区切り）': 'out_of_stock_keywords',
            '有効フラグ': 'active_flag'
        }
        
        for _, row in df.iterrows():
            # 有効フラグをチェック
            active_flag = str(row.get('有効フラグ', '')).strip()
            if active_flag != '有効':
                continue
            
            site_name = str(row.get('サイト名', '')).strip()
            if not site_name or site_name == 'nan':
                continue
            
            # URLパターンをリストに変換
            url_patterns_str = str(row.get('URLパターン（カンマ区切り）', '')).strip()
            url_patterns = [p.strip() for p in url_patterns_str.split(',') if p.strip()] if url_patterns_str != 'nan' else []
            
            # 価格セレクタをリストに変換
            price_selectors_str = str(row.get('価格セレクタ（カンマ区切り）', '')).strip()
            price_selectors = [s.strip() for s in price_selectors_str.split(',') if s.strip()] if price_selectors_str != 'nan' else []
            
            # 価格除外セレクタをリストに変換
            price_exclude_selectors_str = str(row.get('価格除外セレクタ（カンマ区切り）', '')).strip()
            price_exclude_selectors = [s.strip() for s in price_exclude_selectors_str.split(',') if s.strip()] if price_exclude_selectors_str != 'nan' else []
            
            # 在庫セレクタをリストに変換
            stock_selectors_str = str(row.get('在庫セレクタ（カンマ区切り）', '')).strip()
            stock_selectors = [s.strip() for s in stock_selectors_str.split(',') if s.strip()] if stock_selectors_str != 'nan' else []
            
            # 在庫ありキーワードをリストに変換
            in_stock_keywords_str = str(row.get('在庫ありキーワード（カンマ区切り）', '')).strip()
            in_stock_keywords = [k.strip() for k in in_stock_keywords_str.split(',') if k.strip()] if in_stock_keywords_str != 'nan' else []
            
            # 売り切れキーワードをリストに変換
            out_of_stock_keywords_str = str(row.get('売り切れキーワード（カンマ区切り）', '')).strip()
            out_of_stock_keywords = [k.strip() for k in out_of_stock_keywords_str.split(',') if k.strip()] if out_of_stock_keywords_str != 'nan' else []
            
            # サイト設定を作成
            site_config = {
                'name': site_name,
                'url_patterns': url_patterns,
                'price_selectors': price_selectors,
                'stock_selectors': stock_selectors,
                'stock_keywords': {
                    'in_stock': in_stock_keywords,
                    'out_of_stock': out_of_stock_keywords
                }
            }
            
            # 価格除外セレクタがある場合は追加
            if price_exclude_selectors:
                site_config['price_exclude_selectors'] = price_exclude_selectors
            
            sites.append(site_config)
        
        # デフォルト設定（空の設定）
        default_config = {
            'price_selectors': [
                "[class*='price']",
                "[class*='Price']",
                "[id*='price']",
                "[id*='Price']",
                ".price",
                "#price"
            ],
            'stock_selectors': [
                "[class*='stock']",
                "[class*='Stock']",
                "[id*='stock']",
                "[id*='Stock']",
                ".stock",
                "#stock"
            ],
            'stock_keywords': {
                'in_stock': ['在庫あり', '在庫', 'available', 'stock'],
                'out_of_stock': ['売り切れ', '在庫なし', 'out of stock', 'sold out']
            }
        }
        
        return {
            'sites': sites,
            'default': default_config
        }
    
    def merge_with_json_config(self, spreadsheet_config: Dict, json_config_path: Path) -> Dict:
        """
        スプレッドシート設定とJSON設定ファイルをマージする
        
        Args:
            spreadsheet_config: スプレッドシートから読み込んだ設定
            json_config_path: JSON設定ファイルのパス
            
        Returns:
            Dict: マージされた設定（スプレッドシート設定が優先）
        """
        # JSON設定ファイルを読み込む
        try:
            with open(json_config_path, 'r', encoding='utf-8') as f:
                json_config = json.load(f)
        except FileNotFoundError:
            json_config = {'sites': [], 'default': {}}
        
        # スプレッドシート設定を優先してマージ
        # 同じサイト名の場合はスプレッドシート設定で上書き
        merged_sites = {}
        
        # JSON設定のサイトを追加
        for site in json_config.get('sites', []):
            merged_sites[site['name']] = site
        
        # スプレッドシート設定のサイトで上書き
        for site in spreadsheet_config.get('sites', []):
            merged_sites[site['name']] = site
        
        # デフォルト設定はスプレッドシート設定を優先
        default_config = spreadsheet_config.get('default', json_config.get('default', {}))
        
        return {
            'sites': list(merged_sites.values()),
            'default': default_config
        }

