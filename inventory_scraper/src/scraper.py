"""
スクレイピングロジックモジュール
Strategyパターンを使用して各ECサイトのスクレイピングを実装
"""
import time
import random
import re
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Dict, Optional
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd


class BaseScraper(ABC):
    """スクレイパーの基底クラス"""
    
    def __init__(self, browser):
        """
        Args:
            browser: Selenium WebDriverインスタンス
        """
        self.browser = browser
    
    @abstractmethod
    def scrape(self, url: str) -> Dict[str, any]:
        """
        指定されたURLから価格と在庫情報を取得する
        
        Args:
            url: スクレイピング対象のURL
        
        Returns:
            Dict[str, any]: {
                '仕入れ価格': int,
                '在庫ステータス': str,
                '最終更新日時': str
            }
        """
        pass
    
    def wait_and_get_element(self, by, value, timeout=10):
        """
        要素が表示されるまで待機して取得する
        
        Args:
            by: SeleniumのByオブジェクト
            value: セレクタ
            timeout: タイムアウト時間（秒）
        
        Returns:
            WebElement: 見つかった要素
        """
        try:
            element = WebDriverWait(self.browser, timeout).until(
                EC.presence_of_element_located((by, value))
            )
            return element
        except TimeoutException:
            return None
    
    def extract_price(self, text: str) -> Optional[int]:
        """
        テキストから価格を抽出する
        
        Args:
            text: 価格が含まれるテキスト
        
        Returns:
            Optional[int]: 抽出された価格（円）、抽出できない場合はNone
        """
        if not text:
            return None
        
        # 数字とカンマを抽出
        price_match = re.search(r'[\d,]+', text.replace(',', ''))
        if price_match:
            try:
                price = int(price_match.group().replace(',', ''))
                return price
            except ValueError:
                return None
        return None


class AmazonScraper(BaseScraper):
    """Amazon用スクレイパー"""
    
    def scrape(self, url: str) -> Dict[str, any]:
        """
        Amazon商品ページから価格と在庫情報を取得する
        
        Args:
            url: Amazon商品ページのURL
        
        Returns:
            Dict[str, any]: スクレイピング結果
        """
        try:
            self.browser.get(url)
            time.sleep(random.uniform(3, 7))  # ランダムな待機時間
            
            result = {
                '仕入れ価格': 0,
                '在庫ステータス': '不明',
                '最終更新日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # 価格を取得（複数のセレクタを試行）
            price_selectors = [
                '#priceblock_ourprice',
                '#priceblock_dealprice',
                '.a-price-whole',
                '#price',
                '.a-price .a-offscreen'
            ]
            
            price = None
            for selector in price_selectors:
                try:
                    price_element = self.wait_and_get_element(By.CSS_SELECTOR, selector, timeout=5)
                    if price_element:
                        price_text = price_element.text
                        price = self.extract_price(price_text)
                        if price:
                            break
                except Exception:
                    continue
            
            if price:
                result['仕入れ価格'] = price
            else:
                result['仕入れ価格'] = -1
            
            # 在庫ステータスを取得
            stock_selectors = [
                '#availability span',
                '#availability',
                '#outOfStock',
                '.a-color-state'
            ]
            
            stock_status = '在庫あり'  # デフォルト
            for selector in stock_selectors:
                try:
                    stock_element = self.wait_and_get_element(By.CSS_SELECTOR, selector, timeout=3)
                    if stock_element:
                        stock_text = stock_element.text.lower()
                        if '在庫' in stock_text or 'stock' in stock_text:
                            if '残り' in stock_text or 'only' in stock_text:
                                stock_status = '在庫あり'
                            elif '売り切れ' in stock_text or 'out of stock' in stock_text or 'unavailable' in stock_text:
                                stock_status = '売り切れ'
                                break
                except Exception:
                    continue
            
            result['在庫ステータス'] = stock_status
            return result
            
        except Exception as e:
            # 404エラーやその他のエラーの場合
            return {
                '仕入れ価格': 0 if '404' in str(e) else -1,
                '在庫ステータス': '売り切れ' if '404' in str(e) else '不明',
                '最終更新日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }


class MercariScraper(BaseScraper):
    """メルカリ用スクレイパー"""
    
    def scrape(self, url: str) -> Dict[str, any]:
        """
        メルカリ商品ページから価格と在庫情報を取得する
        
        Args:
            url: メルカリ商品ページのURL
        
        Returns:
            Dict[str, any]: スクレイピング結果
        """
        try:
            self.browser.get(url)
            time.sleep(random.uniform(3, 7))  # ランダムな待機時間
            
            result = {
                '仕入れ価格': 0,
                '在庫ステータス': '不明',
                '最終更新日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # 価格を取得
            price_selectors = [
                '.item-price',
                '[data-testid="price"]',
                '.merPrice'
            ]
            
            price = None
            for selector in price_selectors:
                try:
                    price_element = self.wait_and_get_element(By.CSS_SELECTOR, selector, timeout=5)
                    if price_element:
                        price_text = price_element.text
                        price = self.extract_price(price_text)
                        if price:
                            break
                except Exception:
                    continue
            
            if price:
                result['仕入れ価格'] = price
            else:
                result['仕入れ価格'] = -1
            
            # 在庫ステータスを取得（メルカリは通常1点のみ）
            # 「売り切れ」や「取引中」の表示を確認
            status_selectors = [
                '.item-status',
                '[data-testid="status"]',
                '.sold-out'
            ]
            
            stock_status = '在庫あり'  # デフォルト
            for selector in status_selectors:
                try:
                    status_element = self.wait_and_get_element(By.CSS_SELECTOR, selector, timeout=3)
                    if status_element:
                        status_text = status_element.text.lower()
                        if '売り切れ' in status_text or '取引中' in status_text or 'sold' in status_text:
                            stock_status = '売り切れ'
                            break
                except Exception:
                    continue
            
            result['在庫ステータス'] = stock_status
            return result
            
        except Exception as e:
            # 404エラーやその他のエラーの場合
            return {
                '仕入れ価格': 0 if '404' in str(e) else -1,
                '在庫ステータス': '売り切れ' if '404' in str(e) else '不明',
                '最終更新日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }


class YahooScraper(BaseScraper):
    """Yahoo!ショッピング用スクレイパー"""
    
    def scrape(self, url: str) -> Dict[str, any]:
        """
        Yahoo!ショッピング商品ページから価格と在庫情報を取得する
        
        Args:
            url: Yahoo!ショッピング商品ページのURL
        
        Returns:
            Dict[str, any]: スクレイピング結果
        """
        try:
            self.browser.get(url)
            time.sleep(random.uniform(3, 7))  # ランダムな待機時間
            
            result = {
                '仕入れ価格': 0,
                '在庫ステータス': '不明',
                '最終更新日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # 価格を取得
            price_selectors = [
                '.elPriceNumber',
                '.Price__value',
                '[data-testid="price"]'
            ]
            
            price = None
            for selector in price_selectors:
                try:
                    price_element = self.wait_and_get_element(By.CSS_SELECTOR, selector, timeout=5)
                    if price_element:
                        price_text = price_element.text
                        price = self.extract_price(price_text)
                        if price:
                            break
                except Exception:
                    continue
            
            if price:
                result['仕入れ価格'] = price
            else:
                result['仕入れ価格'] = -1
            
            # 在庫ステータスを取得
            stock_selectors = [
                '.elStockStatus',
                '[data-testid="stock-status"]',
                '.StockStatus'
            ]
            
            stock_status = '在庫あり'  # デフォルト
            for selector in stock_selectors:
                try:
                    stock_element = self.wait_and_get_element(By.CSS_SELECTOR, selector, timeout=3)
                    if stock_element:
                        stock_text = stock_element.text.lower()
                        if '在庫' in stock_text:
                            if 'あり' in stock_text:
                                stock_status = '在庫あり'
                            elif 'なし' in stock_text or '売り切れ' in stock_text:
                                stock_status = '売り切れ'
                                break
                except Exception:
                    continue
            
            result['在庫ステータス'] = stock_status
            return result
            
        except Exception as e:
            # 404エラーやその他のエラーの場合
            return {
                '仕入れ価格': 0 if '404' in str(e) else -1,
                '在庫ステータス': '売り切れ' if '404' in str(e) else '不明',
                '最終更新日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }


def get_scraper(url: str, browser, config_loader=None) -> BaseScraper:
    """
    URLに基づいて適切なスクレイパーを返す
    
    優先順位:
    1. 設定ファイルベースのスクレイパー（新規サイト対応）
    2. 既存のハードコーディングされたスクレイパー（互換性のため）
    
    Args:
        url: スクレイピング対象のURL
        browser: Selenium WebDriverインスタンス
        config_loader: ScraperConfigLoaderインスタンス（省略可、再利用のため推奨）
    
    Returns:
        BaseScraper: 適切なスクレイパーインスタンス
    """
    # 設定ファイルベースのスクレイパーを試す（1回だけ実行）
    # config_loaderが渡されていない場合のみ新規作成（パフォーマンス最適化）
    if config_loader is None:
        try:
            from .configurable_scraper import ScraperConfigLoader, create_configurable_scraper
            config_loader = ScraperConfigLoader(browser=browser, use_spreadsheet=True)
        except Exception as e:
            # 設定ファイルの読み込みに失敗した場合は既存の方法にフォールバック
            print(f"設定ファイルの読み込みに失敗しました（既存の方法を使用）: {e}")
            config_loader = None
    
    # 設定ファイルベースのスクレイパーを作成（config_loaderが利用可能な場合）
    if config_loader:
        try:
            from .configurable_scraper import create_configurable_scraper
            scraper = create_configurable_scraper(url, browser, config_loader)
            if isinstance(scraper, BaseScraper):
                return scraper
        except Exception as e:
            # 設定ファイルベースのスクレイパー作成に失敗した場合は既存の方法にフォールバック
            pass
    
    # 既存のハードコーディングされたスクレイパー（互換性のため）
    # 注意: より具体的なURLパターンを先にチェックする必要がある
    url_lower = url.lower()
    
    if 'amazon.co.jp' in url_lower or 'amazon.com' in url_lower:
        return AmazonScraper(browser)
    elif 'mercari.com' in url_lower or 'mercari.jp' in url_lower:
        return MercariScraper(browser)
    elif 'shopping.yahoo.co.jp' in url_lower:
        # Yahoo!ショッピング専用
        return YahooScraper(browser)
    elif 'yahoo.co.jp' in url_lower:
        # その他のYahoo!サイト（オークション含む）は設定ファイルベースで処理済み
        # フォールバック: Yahoo!ショッピング用スクレイパーを使用
        return YahooScraper(browser)
    
    # 設定ファイルが使えない場合はAmazonスクレイパーを使用（デフォルト）
    return AmazonScraper(browser)


def scrape_urls(df: pd.DataFrame, browser) -> pd.DataFrame:
    """
    DataFrameの「仕入れ元URL」列に基づいてスクレイピングを実行する
    
    Args:
        df: スクレイピング対象のURLが含まれるDataFrame
        browser: Selenium WebDriverインスタンス
    
    Returns:
        pd.DataFrame: スクレイピング結果を含むDataFrame
    """
    results = []
    supplier_url_col = '仕入れ元URL'
    
    if supplier_url_col not in df.columns:
        raise Exception(f"DataFrameに「{supplier_url_col}」列が見つかりません")
    
    urls = df[supplier_url_col].tolist()
    total = len(urls)
    
    print(f"スクレイピング開始: {total}件のURLを処理します")
    
    # パフォーマンス最適化: ScraperConfigLoaderを1回だけ作成して全URLで再利用
    # これにより、スプレッドシート設定読み込みが各URLごとに実行されることを防ぐ
    config_loader = None
    try:
        from .configurable_scraper import ScraperConfigLoader
        print("設定ファイルローダーを初期化しています...")
        config_loader = ScraperConfigLoader(browser=browser, use_spreadsheet=True)
        print("設定ファイルローダーの初期化が完了しました")
    except Exception as e:
        print(f"警告: 設定ファイルローダーの初期化に失敗しました（各URLで個別に読み込みます）: {e}")
        config_loader = None
    
    for idx, url in enumerate(urls, 1):
        if pd.isna(url) or url == '':
            continue
        
        print(f"[{idx}/{total}] 処理中: {url}")
        
        try:
            scraper = get_scraper(url, browser, config_loader=config_loader)
            result = scraper.scrape(url)
            result['仕入れ元URL'] = url
            results.append(result)
        except Exception as e:
            print(f"エラーが発生しました ({url}): {e}")
            results.append({
                '仕入れ元URL': url,
                '仕入れ価格': -1,
                '在庫ステータス': '不明',
                '最終更新日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
    
    # 結果をDataFrameに変換
    result_df = pd.DataFrame(results)
    
    # 列の順序を調整
    columns_order = ['仕入れ元URL', '仕入れ価格', '在庫ステータス', '最終更新日時']
    result_df = result_df[columns_order]
    
    print(f"スクレイピング完了: {len(result_df)}件の結果を取得しました")
    return result_df
