"""
設定ファイルベースのスクレイパー
JSON設定ファイルからサイト別のセレクタを読み込んでスクレイピングを実行
"""
import json
import time
import random
import re
import logging
from pathlib import Path
from typing import Dict, Optional, List, Any
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from .scraper import BaseScraper

# ロガーを設定
logger = logging.getLogger(__name__)


class ConfigurableScraper(BaseScraper):
    """設定ファイルベースのスクレイパー"""
    
    def __init__(self, browser, config: Dict):
        """
        Args:
            browser: Selenium WebDriverインスタンス
            config: サイト設定（JSONから読み込んだ辞書）
        """
        super().__init__(browser)
        self.config = config
        self.name = config.get('name', 'Unknown')
    
    def scrape(self, url: str) -> Dict[str, Any]:
        """
        設定ファイルに基づいて価格と在庫情報を取得する
        
        Args:
            url: スクレイピング対象のURL
            
        Returns:
            Dict[str, any]: スクレイピング結果
        """
        try:
            self.browser.get(url)
            # Yahoo!オークションの場合は少し長めに待機（JavaScriptで動的に読み込まれる可能性があるため）
            if 'auctions.yahoo.co.jp' in url.lower():
                time.sleep(random.uniform(5, 10))  # 5-10秒待機
            else:
                time.sleep(random.uniform(3, 7))  # ランダムな待機時間
            
            result = {
                '仕入れ価格': 0,
                '在庫ステータス': '不明',
                '最終更新日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # 価格を取得
            price = None
            if 'auctions.yahoo.co.jp' in url.lower():
                price = self._extract_yahoo_auction_price_from_next_data()
                if price:
                    logger.info(f"  Yahoo!オークション: __NEXT_DATA__から価格を取得しました: {price}円")

            if not price and '/shops/product/' in url.lower():
                price = self._extract_mercari_shop_price()
                if price:
                    logger.info(f"  メルカリSHOP: 価格を取得しました: {price}円")

            if not price:
                price_selectors = self.config.get('price_selectors', [])
                price = self._extract_price_with_selectors(price_selectors, url)
            
            if price:
                result['仕入れ価格'] = price
            else:
                result['仕入れ価格'] = -1
                # デバッグ: 価格が見つからなかった場合、ページ内の価格要素を探す
                # デバッグモードが有効な場合のみ実行（パフォーマンス最適化）
                from .config import ENABLE_DEBUG_MODE
                if ENABLE_DEBUG_MODE:
                    logger.warning(f"  警告: 価格が見つかりませんでした（URL: {url[:80]}...）。デバッグ情報を出力します...")
                    self._debug_price_elements(url)
                else:
                    # デバッグモードが無効な場合は警告のみ出力
                    logger.warning(f"  警告: 価格が見つかりませんでした（URL: {url[:80]}...）")
            
            # 在庫ステータスを取得
            stock_selectors = self.config.get('stock_selectors', [])
            stock_keywords = self.config.get('stock_keywords', {})
            stock_status = self._extract_stock_status_with_selectors(
                stock_selectors, 
                stock_keywords
            )
            result['在庫ステータス'] = stock_status
            
            return result
            
        except Exception as e:
            # HTTP関連の例外をチェックしてステータスコードを取得
            status_code = None
            is_http_error = False
            
            # requestsライブラリの例外をチェック
            try:
                import requests
                if isinstance(e, requests.exceptions.HTTPError):
                    is_http_error = True
                    if hasattr(e, 'response') and e.response is not None:
                        status_code = e.response.status_code
            except ImportError:
                pass
            
            # urllib3ライブラリの例外をチェック
            try:
                import urllib3.exceptions
                if isinstance(e, urllib3.exceptions.HTTPError):
                    is_http_error = True
                    if hasattr(e, 'status') and e.status is not None:
                        status_code = e.status
            except (ImportError, AttributeError):
                pass
            
            # httpxライブラリの例外をチェック
            try:
                import httpx  # type: ignore
                if isinstance(e, httpx.HTTPStatusError):
                    is_http_error = True
                    if hasattr(e, 'response') and e.response is not None:
                        status_code = e.response.status_code
            except ImportError:
                pass
            
            # Seleniumの例外からHTTPステータスコードを抽出する試み
            if status_code is None:
                error_str = str(e)
                # エラーメッセージから404などのステータスコードを探す
                status_match = re.search(r'\b([4-5]\d{2})\b', error_str)
                if status_match:
                    try:
                        status_code = int(status_match.group(1))
                        is_http_error = True
                    except (ValueError, AttributeError):
                        pass
            
            # ステータスコード404の場合は在庫切れとして扱う
            if status_code == 404:
                logger.warning(f"  HTTP 404エラーが検出されました (URL: {url[:80]}...): {e}")
                return {
                    '仕入れ価格': 0,
                    '在庫ステータス': '売り切れ',
                    '最終更新日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
            else:
                # その他のエラーの場合
                error_msg = f"HTTP {status_code}エラー" if status_code else "エラー"
                logger.error(f"  {error_msg}が発生しました (URL: {url[:80]}...): {type(e).__name__}: {e}")
                return {
                    '仕入れ価格': -1,
                    '在庫ステータス': '不明',
                    '最終更新日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
    
    def _debug_price_elements(self, url: str):
        """
        デバッグ用: ページ内の価格要素を探して出力する
        
        Args:
            url: スクレイピング対象のURL
        """
        try:
            logger.warning(f"  デバッグ: {url} の価格要素を検索中...")
            
            # ページの読み込みを待つ（追加の待機時間）
            import time
            time.sleep(2)
            
            # 一般的な価格セレクタで全要素を取得
            debug_selectors = [
                "[class*='price']",
                "[id*='price']",
                "[class*='Price']",
                "[id*='Price']",
                ".price",
                "#price",
                "span[class*='price']",
                "div[class*='price']",
                "span[class*='Price']",
                "div[class*='Price']"
            ]
            
            found_elements = []
            all_elements_with_price_text = []  # 価格を含むすべての要素（価格抽出失敗も含む）
            
            for selector in debug_selectors:
                try:
                    elements = self.browser.find_elements(By.CSS_SELECTOR, selector)
                    logger.warning(f"    セレクタ '{selector}': {len(elements)}個の要素が見つかりました")
                    
                    for element in elements:
                        try:
                            text = element.text.strip()
                            if not text:
                                continue
                            
                            # 価格を含む可能性のあるテキストをすべて記録
                            if '円' in text or re.search(r'[\d,]+', text):
                                try:
                                    class_name = element.get_attribute('class') or ''
                                    element_id = element.get_attribute('id') or ''
                                    tag_name = element.tag_name
                                    
                                    price = self.extract_price(text)
                                    if price and price > 0:
                                        found_elements.append({
                                            'selector': selector,
                                            'tag': tag_name,
                                            'class': class_name,
                                            'id': element_id,
                                            'text': text[:200],  # 最初の200文字
                                            'price': price
                                        })
                                    else:
                                        # 価格抽出に失敗したが、価格らしいテキストを含む要素
                                        all_elements_with_price_text.append({
                                            'selector': selector,
                                            'tag': tag_name,
                                            'class': class_name,
                                            'id': element_id,
                                            'text': text[:200]
                                        })
                                except Exception as e:
                                    logger.warning(f"      要素処理エラー: {e}")
                                    continue
                        except:
                            continue
                except Exception as e:
                    logger.warning(f"    セレクタ '{selector}' の検索エラー: {e}")
                    continue
            
            if found_elements:
                logger.warning(f"  見つかった価格要素: {len(found_elements)}件")
                for i, elem in enumerate(found_elements[:10], 1):  # 最初の10件のみ表示
                    logger.warning(f"    [{i}] 価格: {elem['price']}円")
                    logger.warning(f"        セレクタ: {elem['selector']}")
                    logger.warning(f"        タグ: {elem['tag']}")
                    if elem['class']:
                        logger.warning(f"        クラス: {elem['class']}")
                    if elem['id']:
                        logger.warning(f"        ID: {elem['id']}")
                    logger.warning(f"        テキスト: {elem['text']}")
            else:
                logger.warning(f"  価格要素が見つかりませんでした")
                
                # 価格らしいテキストを含む要素を出力（価格抽出に失敗したものも含む）
                if all_elements_with_price_text:
                    logger.warning(f"  価格らしいテキストを含む要素: {len(all_elements_with_price_text)}件")
                    for i, elem in enumerate(all_elements_with_price_text[:5], 1):  # 最初の5件のみ表示
                        logger.warning(f"    [{i}] タグ: {elem['tag']}")
                        if elem['class']:
                            logger.warning(f"        クラス: {elem['class']}")
                        if elem['id']:
                            logger.warning(f"        ID: {elem['id']}")
                        logger.warning(f"        テキスト: {elem['text']}")
                
                # XPathを使用して「円」を含む要素を検索（より広範囲な検索）
                try:
                    logger.warning(f"  XPathで「円」を含む要素を検索中...")
                    xpath_elements = self.browser.find_elements(By.XPATH, "//*[contains(text(), '円')]")
                    logger.warning(f"  「円」を含む要素: {len(xpath_elements)}件")
                    for i, elem in enumerate(xpath_elements[:10], 1):  # 最初の10件のみ表示
                        try:
                            text = elem.text.strip()
                            if text:
                                tag_name = elem.tag_name
                                class_name = elem.get_attribute('class') or ''
                                element_id = elem.get_attribute('id') or ''
                                
                                # 親要素の情報も取得
                                try:
                                    parent = elem.find_element(By.XPATH, './..')
                                    parent_tag = parent.tag_name
                                    parent_class = parent.get_attribute('class') or ''
                                    parent_id = parent.get_attribute('id') or ''
                                    parent_text = parent.text.strip()[:100]
                                    logger.warning(f"    [{i}] タグ: {tag_name}, クラス: {class_name[:50]}, ID: {element_id}")
                                    logger.warning(f"        テキスト: {text[:100]}")
                                    logger.warning(f"        親要素: {parent_tag}, クラス: {parent_class[:50]}, ID: {parent_id}")
                                    logger.warning(f"        親要素のテキスト: {parent_text}")
                                    
                                    # 親要素から価格を抽出してみる
                                    parent_price = self.extract_price(parent_text)
                                    if parent_price and parent_price > 0:
                                        logger.warning(f"        → 親要素から価格を抽出: {parent_price}円")
                                except:
                                    logger.warning(f"    [{i}] タグ: {tag_name}, クラス: {class_name[:50]}, ID: {element_id}, テキスト: {text[:100]}")
                        except:
                            continue
                except Exception as e:
                    logger.warning(f"  XPath検索エラー: {e}")
                
                # ページのタイトルとURLを確認
                try:
                    page_title = self.browser.title
                    current_url = self.browser.current_url
                    logger.warning(f"  ページタイトル: {page_title}")
                    logger.warning(f"  現在のURL: {current_url}")
                    
                    # ページソースの一部を確認（デバッグ用）
                    page_source = self.browser.page_source
                    if 'Price' in page_source or 'price' in page_source:
                        # 価格関連の文字列を含む部分を抽出
                        import re
                        price_matches = re.findall(r'[0-9,]+円', page_source)
                        if price_matches:
                            logger.warning(f"  ページソース内の価格らしい文字列: {price_matches[:10]}")
                except Exception as e:
                    logger.warning(f"  ページ情報取得エラー: {e}")
                
        except Exception as e:
            logger.warning(f"  デバッグ中にエラー: {e}")
            import traceback
            logger.warning(f"  エラー詳細: {traceback.format_exc()}")

    def _select_best_price_candidate(
        self, 
        candidates: List[Dict], 
        min_price: int = 100, 
        reasonable_range: tuple = (1000, 10000)
    ) -> Optional[Dict]:
        """
        価格候補のリストから最適な候補を選択する
        
        Args:
            candidates: 価格候補のリスト（各候補は'price'と'text'キーを持つ辞書）
            min_price: 最小価格（デフォルト: 100円）
            reasonable_range: 合理的な価格範囲（デフォルト: (1000, 10000)）
            
        Returns:
            選択された候補の辞書（'price', 'text', 'selector'キーを含む）、見つからない場合はNone
        """
        if not candidates:
            return None
        
        # 重複を排除（同じ価格の候補を1つだけ残す）
        seen_prices = set()
        unique_candidates = []
        for candidate in candidates:
            if candidate['price'] not in seen_prices:
                seen_prices.add(candidate['price'])
                unique_candidates.append(candidate)
        
        # 関連商品セクションを除外（「この商品も注目されています」「おすすめ」などのキーワードを含む要素を除外）
        exclude_keywords = ['この商品も注目されています', 'おすすめ', '関連商品', 'この商品も', '注目されています', 'おすすめ商品']
        main_content_candidates = []
        for candidate in unique_candidates:
            text_lower = candidate['text'].lower()
            # 除外キーワードを含まない候補を追加
            if not any(keyword in text_lower for keyword in exclude_keywords):
                main_content_candidates.append(candidate)
            else:
                logger.warning(f"    関連商品セクションの候補を除外: {candidate['price']}円 (テキスト: {candidate['text'][:50]})")
        
        # メインコンテンツエリアの候補がない場合、すべての候補を使用
        if not main_content_candidates:
            logger.warning(f"  メインコンテンツエリアの候補が見つかりませんでした。すべての候補を使用します。")
            main_content_candidates = unique_candidates
        
        # メイン商品の価格を優先（「即決」「（税0円）」「送料」などのキーワードを含む価格を優先）
        priority_keywords = ['即決', '（税0円）', '税0円', '送料', '配送方法', '入札する', '今すぐ落札']
        priority_candidates = []
        for candidate in main_content_candidates:
            text = candidate['text']
            # 優先キーワードを含む候補を追加
            if any(keyword in text for keyword in priority_keywords):
                priority_candidates.append(candidate)
                logger.warning(f"    優先キーワードを含む候補を追加: {candidate['price']}円 (テキスト: {candidate['text'][:50]})")
        
        # 優先候補がある場合、それを使用。ない場合はすべての候補を使用
        if priority_candidates:
            logger.warning(f"  優先キーワードを含む候補が見つかりました（{len(priority_candidates)}件）。これらを優先します。")
            main_content_candidates = priority_candidates
        
        # 価格でソート
        main_content_candidates.sort(key=lambda x: x['price'])
        
        logger.warning(f"  価格候補（重複排除・関連商品除外・優先キーワード適用後、全{len(main_content_candidates)}件）: {[(c['price'], c['text'][:50]) for c in main_content_candidates]}")
        
        # 合理的な範囲内の価格を優先
        reasonable_min, reasonable_max = reasonable_range
        reasonable_candidates = [c for c in main_content_candidates if reasonable_min <= c['price'] <= reasonable_max]
        if reasonable_candidates:
            # 合理的な範囲内の価格から、最も小さい価格を選択
            best_candidate = reasonable_candidates[0]
            return {
                'price': best_candidate['price'],
                'text': best_candidate['text'][:100],
                'selector': 'direct_current_price_element_reasonable',
                **{k: v for k, v in best_candidate.items() if k not in ['price', 'text']}
            }
        else:
            # 合理的な範囲内に価格がない場合、最小価格以上の候補から最も小さい価格を選択
            if main_content_candidates and main_content_candidates[0]['price'] >= min_price:
                best_candidate = main_content_candidates[0]
                return {
                    'price': best_candidate['price'],
                    'text': best_candidate['text'][:100],
                    'selector': 'direct_current_price_element_all',
                    **{k: v for k, v in best_candidate.items() if k not in ['price', 'text']}
                }
        
        return None

    def _extract_yahoo_auction_price_from_next_data(self) -> Optional[int]:
        """
        Yahoo!オークションの__NEXT_DATA__から現在価格（税額込み）を取得する

        Returns:
            Optional[int]: 価格、取得できない場合はNone
        """
        try:
            next_data_text = self.browser.execute_script(
                "return document.querySelector('script#__NEXT_DATA__')?.textContent || null;"
            )
            if not next_data_text:
                return None

            next_data = json.loads(next_data_text)
            item = (
                next_data.get('props', {})
                .get('pageProps', {})
                .get('initialState', {})
                .get('item', {})
                .get('detail', {})
                .get('item', {})
            )
            if not item:
                return None

            candidates = [
                item.get('taxinPrice'),
                item.get('taxinStartPrice'),
                item.get('price'),
                item.get('initPrice')
            ]

            for candidate in candidates:
                if candidate is None:
                    continue
                try:
                    price = int(float(candidate))
                    if price > 0:
                        return price
                except (ValueError, TypeError):
                    continue

            tax_rate = item.get('taxRate')
            base_price = item.get('price')
            if tax_rate and base_price:
                try:
                    base_price_int = int(float(base_price))
                    tax_rate_float = float(tax_rate)
                    price_with_tax = int(round(base_price_int * (1 + (tax_rate_float / 100))))
                    if price_with_tax > 0:
                        return price_with_tax
                except (ValueError, TypeError):
                    return None
        except Exception as e:
            logger.debug(f"__NEXT_DATA__からの価格取得に失敗しました: {e}")
            return None

        return None
    
    def _extract_mercari_shop_price(self) -> Optional[int]:
        """
        メルカリSHOPページのメイン価格を取得する
        「送料込み」付近やタイトル付近から優先的に抽出
        「¥」記号を含む価格パターンを優先的に検索
        """
        exclude_selectors = self.config.get('price_exclude_selectors', [])
        exclude_keywords = [
            'このショップの商品',
            'おすすめ',
            'ランキング',
            'この商品を見ている人に',
            'shopsゲーム',
            '人気の商品'
        ]
        
        # 方法1: 「送料込み」を含む要素の親要素から「¥」を含む価格を取得
        try:
            shipping_elements = self.browser.find_elements(By.XPATH, "//*[contains(text(), '送料込み')]")
            for shipping_el in shipping_elements:
                try:
                    # 親要素を取得
                    parent = shipping_el.find_element(By.XPATH, './..')
                    parent_text = parent.text.strip()
                    
                    # 除外キーワードをチェック
                    if any(keyword in parent_text.lower() for keyword in exclude_keywords):
                        continue
                    
                    # 「¥」を含む価格パターンを先に探す（優先）
                    # 価格が見つかった場合は除外セレクタチェックをスキップ
                    price = self._extract_yen_price_from_text(parent_text)
                    if price:
                        # 価格が見つかった場合は除外セレクタチェックをスキップして返す
                        logger.debug(f"  メルカリSHOP: 「送料込み」親要素から価格を取得: {price}円")
                        return price
                    
                    # 価格が見つからなかった場合のみ除外セレクタをチェック
                    if self._is_excluded_element(parent, exclude_selectors):
                        continue
                except Exception:
                    continue
        except Exception:
            pass
        
        # 方法2: h1要素の近傍から「¥」を含む価格を取得
        try:
            h1_element = self.browser.find_element(By.TAG_NAME, 'h1')
            if h1_element:
                # h1の親の親要素（商品情報コンテナ）を取得
                h1_parent = h1_element.find_element(By.XPATH, './..')
                h1_grandparent = h1_parent.find_element(By.XPATH, './..')
                
                # 祖父要素の子要素を探索して「¥」を含む価格を探す
                children = h1_grandparent.find_elements(By.XPATH, './*')
                for child in children:
                    try:
                        child_text = child.text.strip()
                        if not child_text:
                            continue
                        
                        # 除外キーワードをチェック
                        if any(keyword in child_text.lower() for keyword in exclude_keywords):
                            continue
                        
                        # 除外セレクタをチェック
                        if self._is_excluded_element(child, exclude_selectors):
                            continue
                        
                        # 「¥」を含む価格パターンを探す
                        price = self._extract_yen_price_from_text(child_text)
                        if price:
                            logger.debug(f"  メルカリSHOP: h1近傍から価格を取得: {price}円")
                            return price
                    except Exception:
                        continue
        except Exception:
            pass
        
        # 方法3: フォールバック - 既存のロジック（「¥」を含む価格を優先）
        try:
            shipping_elements = self.browser.find_elements(By.XPATH, "//*[contains(text(), '送料込み')]")
            for shipping_el in shipping_elements:
                current = shipping_el
                for _ in range(4):
                    try:
                        text = current.text.strip()
                        text_lower = text.lower()
                        if text and not any(keyword in text_lower for keyword in exclude_keywords):
                            if not self._is_excluded_element(current, exclude_selectors):
                                # まず「¥」を含む価格を探す
                                price = self._extract_yen_price_from_text(text)
                                if price:
                                    return price
                                # 見つからなければ従来のロジック
                                price = self._extract_first_price_from_text(text)
                                if price:
                                    return price
                        current = current.find_element(By.XPATH, './..')
                    except Exception:
                        break
        except Exception:
            pass
        
        return None
    
    def _extract_yen_price_from_text(self, text: str) -> Optional[int]:
        """
        テキストから「¥」に続く価格を抽出する
        
        Args:
            text: 検索対象のテキスト
            
        Returns:
            Optional[int]: 抽出された価格、見つからない場合はNone
        """
        if not text:
            return None
        
        # 「¥」または「￥」に続く数字パターンをマッチ（例: ¥4,800 や ¥4800）
        matches = re.findall(r'[¥￥]\s*([0-9,]+)', text)
        if matches:
            try:
                return int(matches[0].replace(',', ''))
            except (ValueError, TypeError):
                pass
        
        return None

    def _extract_first_price_from_text(self, text: str) -> Optional[int]:
        """
        テキストから最初の価格を抽出する
        """
        if not text:
            return None
        matches = re.findall(r'\d{1,3}(?:,\d{3})*', text)
        if not matches:
            return None
        try:
            return int(matches[0].replace(',', ''))
        except (ValueError, TypeError):
            return None
    
    def _extract_price_from_container(self, container, exclude_selectors: List[str]) -> Optional[int]:
        """
        指定コンテナ内の価格を抽出（除外セレクタ考慮）
        """
        try:
            price_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '¥')]")
            for price_el in price_elements:
                try:
                    price_text = price_el.text.strip()
                    price = self.extract_price(price_text)
                    if not price or price <= 0:
                        continue
                    if self._is_excluded_element(price_el, exclude_selectors):
                        continue
                    return price
                except Exception:
                    continue
        except Exception:
            return None
        return None
    
    def _is_excluded_element(self, element, exclude_selectors: List[str]) -> bool:
        """
        要素が除外セレクタに該当するか判定
        """
        for exclude_selector in exclude_selectors:
            try:
                if element.find_elements(By.CSS_SELECTOR, exclude_selector):
                    return True
                current = element
                for _ in range(3):
                    try:
                        current = current.find_element(By.XPATH, './..')
                        if current.find_elements(By.CSS_SELECTOR, exclude_selector):
                            return True
                    except Exception:
                        break
            except Exception:
                continue
        return False
    
    def _extract_price_with_selectors(self, selectors: List[str], url: str = '') -> Optional[int]:
        """
        複数のセレクタを試行して価格を取得する
        
        Args:
            selectors: CSSセレクタのリスト
            
        Returns:
            Optional[int]: 抽出された価格、見つからない場合はNone
        """
        exclude_selectors = self.config.get('price_exclude_selectors', [])
        found_prices = []  # 見つかった価格をすべて保存
        
        # デバッグ: セレクタの試行状況を出力（DEBUGレベル）
        logger.debug(f"  価格セレクタを試行中: {len(selectors)}個のセレクタ")
        
        for selector in selectors:
            try:
                # セレクタで複数の要素が見つかる可能性があるため、すべて取得
                elements = self.browser.find_elements(By.CSS_SELECTOR, selector)
                
                if not elements:
                    logger.debug(f"    セレクタ '{selector}': 要素が見つかりませんでした")
                    continue
                
                logger.debug(f"    セレクタ '{selector}': {len(elements)}個の要素が見つかりました")
                
                for element in elements:
                    try:
                        # 除外セレクタのチェック（要素またはその親要素が除外セレクタに該当するか確認）
                        should_exclude = False
                        for exclude_selector in exclude_selectors:
                            try:
                                # 要素自体が除外セレクタに該当するか確認
                                if element.find_elements(By.CSS_SELECTOR, exclude_selector):
                                    should_exclude = True
                                    break
                                
                                # 親要素を再帰的にチェック（最大5階層まで）
                                current_element = element
                                for depth in range(5):
                                    try:
                                        parent = current_element.find_element(By.XPATH, './..')
                                        # 親要素自体が除外セレクタに該当するか確認
                                        if parent.find_elements(By.CSS_SELECTOR, exclude_selector):
                                            should_exclude = True
                                            break
                                        # 親要素のクラス名やIDが除外セレクタに含まれるか確認
                                        parent_class = parent.get_attribute('class') or ''
                                        parent_id = parent.get_attribute('id') or ''
                                        # [class*='xxx']形式のセレクタをチェック
                                        if '[class*=' in exclude_selector:
                                            class_pattern = exclude_selector.split("'")[1] if "'" in exclude_selector else exclude_selector.split('"')[1] if '"' in exclude_selector else ''
                                            if class_pattern and class_pattern in parent_class:
                                                should_exclude = True
                                                break
                                        # [id*='xxx']形式のセレクタをチェック
                                        if '[id*=' in exclude_selector:
                                            id_pattern = exclude_selector.split("'")[1] if "'" in exclude_selector else exclude_selector.split('"')[1] if '"' in exclude_selector else ''
                                            if id_pattern and id_pattern in parent_id:
                                                should_exclude = True
                                                break
                                        current_element = parent
                                    except:
                                        break
                                
                                if should_exclude:
                                    break
                                
                                # 要素自体のクラス名やIDが除外セレクタに含まれるか確認
                                element_class = element.get_attribute('class') or ''
                                element_id = element.get_attribute('id') or ''
                                # [class*='xxx']形式のセレクタをチェック
                                if '[class*=' in exclude_selector:
                                    class_pattern = exclude_selector.split("'")[1] if "'" in exclude_selector else exclude_selector.split('"')[1] if '"' in exclude_selector else ''
                                    if class_pattern and class_pattern in element_class:
                                        should_exclude = True
                                        break
                                # [id*='xxx']形式のセレクタをチェック
                                if '[id*=' in exclude_selector:
                                    id_pattern = exclude_selector.split("'")[1] if "'" in exclude_selector else exclude_selector.split('"')[1] if '"' in exclude_selector else ''
                                    if id_pattern and id_pattern in element_id:
                                        should_exclude = True
                                        break
                            except:
                                pass
                        
                        if should_exclude:
                            logger.debug(f"      除外セレクタに該当するためスキップ: {selector}")
                            continue
                        
                        # 要素のテキストを確認
                        price_text = element.text.strip()
                        if not price_text:
                            continue
                        
                        price_text_lower = price_text.lower()
                        
                        # IDセレクタ（#で始まる）の場合は信頼性が高いので、親要素チェックをスキップ
                        is_id_selector = selector.startswith('#') or '[id=' in selector or '[id*=' in selector
                        
                        # 除外キーワードチェック（IDセレクタの場合は緩和）
                        exclude_keywords = ['楽天カード', 'ポイント利用', 'special', 'offer', '送料別', '内訳', '倍']
                        # IDセレクタの場合は「送料別」のみチェック（「送料無料」は除外しない）
                        if is_id_selector:
                            exclude_keywords = ['楽天カード', 'ポイント利用', 'special', 'offer', '送料別', '内訳']
                        
                        if any(keyword in price_text_lower for keyword in exclude_keywords):
                            continue
                        
                        # 親要素もチェック（IDセレクタの場合はスキップ）
                        if not is_id_selector:
                            try:
                                parent = element.find_element(By.XPATH, './..')
                                parent_text = parent.text.lower()
                                if any(keyword in parent_text for keyword in exclude_keywords):
                                    continue
                            except:
                                pass
                        
                        # 価格を抽出
                        price = self.extract_price(price_text)
                        if price and price > 0:
                            logger.debug(f"      価格を発見: {price}円 (テキスト: {price_text[:50]})")
                            found_prices.append({
                                'price': price,
                                'text': price_text[:50],
                                'selector': selector
                            })
                        else:
                            logger.debug(f"      価格抽出失敗: テキスト='{price_text[:50]}'")
                    except Exception as e:
                        logger.debug(f"      要素処理エラー: {e}")
                        continue
                        
            except Exception:
                continue
        
        # Yahoo!オークションの場合、親要素から価格を抽出するフォールバック処理
        if not found_prices and 'auctions.yahoo.co.jp' in url.lower():
            try:
                # まず、設定ファイルのセレクタで価格を探す（再試行）
                logger.warning(f"  Yahoo!オークション: 設定ファイルのセレクタで価格を再検索中...")
                for selector in self.config.get('price_selectors', []):
                    try:
                        elements = self.browser.find_elements(By.CSS_SELECTOR, selector)
                        for elem in elements:
                            try:
                                text = elem.text.strip()
                                if text:
                                    price = self.extract_price(text)
                                    if price and price > 0:
                                        # 親要素に「現在」が含まれているか確認
                                        try:
                                            parent = elem.find_element(By.XPATH, './..')
                                            parent_text = parent.text.strip()
                                            if '現在' in parent_text or '現在' in text:
                                                found_prices.append({
                                                    'price': price,
                                                    'text': text[:100],
                                                    'selector': selector
                                                })
                                                logger.warning(f"  設定セレクタ '{selector}' から価格を発見: {price}円 (テキスト: {text[:100]})")
                                                break
                                        except:
                                            # 親要素チェックに失敗した場合でも、要素自体に「現在」が含まれていれば追加
                                            if '現在' in text:
                                                found_prices.append({
                                                    'price': price,
                                                    'text': text[:100],
                                                    'selector': selector
                                                })
                                                logger.warning(f"  設定セレクタ '{selector}' から価格を発見（現在含む）: {price}円 (テキスト: {text[:100]})")
                                                break
                            except:
                                continue
                        if found_prices:
                            break
                    except:
                        continue
                
                # 「現在」を含む要素を直接検索
                if not found_prices:
                    try:
                        # 「現在」を含むすべての要素を検索
                        current_price_elements = self.browser.find_elements(By.XPATH, "//*[contains(text(), '現在')]")
                        logger.warning(f"  「現在」を含む要素を{len(current_price_elements)}件発見")
                        
                        all_current_candidates = []  # すべての「現在」価格候補を保存
                        
                        for i, elem in enumerate(current_price_elements[:30]):  # 最初の30件を確認
                            try:
                                text = elem.text.strip()
                                if '現在' in text:
                                    price = self.extract_price(text)
                                    if price and price > 0:
                                        logger.warning(f"    [{i}] 価格: {price}円, テキスト: {text[:150]}")
                                        all_current_candidates.append({
                                            'price': price,
                                            'text': text[:200],
                                            'element': elem
                                        })
                            except Exception as e:
                                logger.debug(f"    要素[{i}]の処理エラー: {e}")
                                continue
                        
                        # すべての候補から最適な価格を選択
                        if all_current_candidates:
                            best_candidate = self._select_best_price_candidate(all_current_candidates)
                            if best_candidate:
                                found_prices.append(best_candidate)
                                selector_type = '合理的範囲' if best_candidate['selector'] == 'direct_current_price_element_reasonable' else '全候補'
                                logger.warning(f"  「現在」を含む要素から価格を抽出（{selector_type}）: {best_candidate['price']}円 (テキスト: {best_candidate['text'][:100]})")
                    except Exception as e:
                        logger.warning(f"  直接検索エラー: {e}")
                
                # 「円」を含む要素の親要素から価格を抽出（「現在」を含む要素から価格が見つからなかった場合のみ）
                if not found_prices:
                    yen_elements = self.browser.find_elements(By.XPATH, "//*[contains(text(), '円')]")
                    logger.warning(f"  Yahoo!オークション: 「円」を含む要素を{len(yen_elements)}件発見")
                    
                    # まず「現在」というキーワードを含む親要素を優先的に探す
                    current_price_found = False
                    current_price_candidates = []  # 「現在」を含むすべての価格候補を保存
                    
                    for yen_elem in yen_elements[:50]:  # 最初の50件を確認
                        try:
                            # 直接の親要素を確認
                            parent = yen_elem.find_element(By.XPATH, './..')
                            parent_text = parent.text.strip()
                            
                            # 親要素に「現在」が含まれている場合
                            if parent_text and '現在' in parent_text:
                                price = self.extract_price(parent_text)
                                if price and price > 0:
                                    current_price_candidates.append({
                                        'price': price,
                                        'text': parent_text[:200],
                                        'level': 'parent'
                                    })
                                    logger.warning(f"    親要素から価格候補を追加: {price}円 (テキスト: {parent_text[:150]})")
                            
                            # 親要素に「現在」が含まれていない場合、祖父要素も確認
                            try:
                                grandparent = parent.find_element(By.XPATH, './..')
                                grandparent_text = grandparent.text.strip()
                                if grandparent_text and '現在' in grandparent_text:
                                    price = self.extract_price(grandparent_text)
                                    if price and price > 0:
                                        current_price_candidates.append({
                                            'price': price,
                                            'text': grandparent_text[:200],
                                            'level': 'grandparent'
                                        })
                                        logger.warning(f"    祖父要素から価格候補を追加: {price}円 (テキスト: {grandparent_text[:150]})")
                            except:
                                pass
                            
                            # さらに上位の要素も確認（曽祖父要素）
                            try:
                                great_grandparent = grandparent.find_element(By.XPATH, './..')
                                great_grandparent_text = great_grandparent.text.strip()
                                if great_grandparent_text and '現在' in great_grandparent_text:
                                    price = self.extract_price(great_grandparent_text)
                                    if price and price > 0:
                                        current_price_candidates.append({
                                            'price': price,
                                            'text': great_grandparent_text[:200],
                                            'level': 'great_grandparent'
                                        })
                                        logger.warning(f"    曽祖父要素から価格候補を追加: {price}円 (テキスト: {great_grandparent_text[:150]})")
                            except:
                                pass
                        except:
                            continue
                    
                    # 「現在」を含む価格候補から、最適な価格を選択
                    if current_price_candidates:
                        best_candidate = self._select_best_price_candidate(current_price_candidates)
                        if best_candidate:
                            # levelキーがある場合は、selectorを上書き
                            if 'level' in best_candidate:
                                if best_candidate['selector'] == 'direct_current_price_element_reasonable':
                                    best_candidate['selector'] = f"{best_candidate['level']}_of_yen_element_with_current_reasonable"
                                else:
                                    best_candidate['selector'] = f"{best_candidate['level']}_of_yen_element_with_current"
                            
                            found_prices.append(best_candidate)
                            level_str = f"{best_candidate.get('level', '')}要素から" if 'level' in best_candidate else ''
                            selector_type = '合理的範囲' if 'reasonable' in best_candidate['selector'] else ''
                            logger.warning(f"  「現在」を含む{level_str}価格を抽出（{selector_type}）: {best_candidate['price']}円 (テキスト: {best_candidate['text'][:100]})")
                            current_price_found = True
                        else:
                            logger.warning(f"  「現在」を含む価格候補が見つかりましたが、価格が小さすぎます")
                
                # 「現在」を含む価格が見つからなかった場合、通常の親要素から価格を抽出
                if not current_price_found:
                    logger.warning(f"  「現在」を含む価格が見つかりませんでした。通常の親要素から価格を抽出します...")
                    all_parent_prices = []
                    for yen_elem in yen_elements[:50]:  # 最初の50件を確認
                        try:
                            parent = yen_elem.find_element(By.XPATH, './..')
                            parent_text = parent.text.strip()
                            if parent_text:
                                # 関連商品セクションを除外（「この商品も注目されています」など）
                                if 'この商品も注目されています' in parent_text or 'おすすめ' in parent_text or '関連商品' in parent_text:
                                    continue
                                
                                price = self.extract_price(parent_text)
                                # 価格が1000円以上の場合のみ有効（121円のような小さい価格を除外）
                                if price and price >= 1000:
                                    all_parent_prices.append({
                                        'price': price,
                                        'text': parent_text[:100],
                                        'selector': 'parent_of_yen_element'
                                    })
                        except:
                            continue
                    
                    if all_parent_prices:
                        # 価格でソートして、最も小さい価格を選択（通常、現在価格が最も小さい）
                        all_parent_prices.sort(key=lambda x: x['price'])
                        logger.warning(f"  見つかった価格候補: {[p['price'] for p in all_parent_prices[:10]]}")
                        found_prices.append(all_parent_prices[0])  # 最も小さい価格を追加
                        logger.warning(f"  最も小さい価格を選択: {all_parent_prices[0]['price']}円 (テキスト: {all_parent_prices[0]['text'][:100]})")
            except Exception as e:
                logger.warning(f"  親要素からの価格抽出エラー: {e}")
                import traceback
                logger.warning(f"  エラー詳細: {traceback.format_exc()}")
        
        # 見つかった価格から選択（Yahoo!オークションの場合は「現在」を含む価格を優先）
        if found_prices:
            # メルカリSHOPは事前の専用抽出で処理済み。ここでは通常の選択ロジックを適用する。
            
            # Yahoo!オークションの場合、「現在」を含む価格を優先
            if 'auctions.yahoo.co.jp' in url.lower():
                # 「現在」を含む価格を優先的に選択
                current_prices = [p for p in found_prices if '現在' in p.get('text', '')]
                if current_prices:
                    selected_price = current_prices[0]['price']
                    logger.warning(f"  Yahoo!オークション: 「現在」を含む価格を選択: {selected_price}円")
                    return selected_price
                
                # 「現在」を含む価格がない場合、関連商品セクションを除外
                # 「この商品も注目されています」「おすすめ」などのキーワードを含む価格を除外
                filtered_prices = []
                for p in found_prices:
                    text = p.get('text', '').lower()
                    if 'この商品も注目されています' not in text and 'おすすめ' not in text and '関連商品' not in text:
                        filtered_prices.append(p)
                
                if filtered_prices:
                    # フィルタリング後の価格から、最も小さい価格を選択（通常、現在価格が最も小さい）
                    filtered_prices.sort(key=lambda x: x['price'])
                    selected_price = filtered_prices[0]['price']
                    logger.warning(f"  Yahoo!オークション: フィルタリング後の価格を選択: {selected_price}円 (候補: {[p['price'] for p in filtered_prices[:5]]})")
                    return selected_price
            
            # デバッグ情報を出力（DEBUGレベル）
            if len(found_prices) > 1:
                logger.debug(f"  複数の価格が見つかりました:")
                for p in found_prices:
                    logger.debug(f"    - {p['price']}円 (セレクタ: {p['selector']}, テキスト: {p['text']})")
            
            # セレクタの優先順位に基づいて、最初に見つかった価格を返す
            # 設定ファイルのセレクタ順序を保持するため、セレクタの順序で最初の価格を選択
            selected_price = None
            selected_selector = None
            
            # 設定ファイルのセレクタ順序に従って、最初に見つかった価格を選択
            for selector in selectors:
                for price_info in found_prices:
                    if price_info['selector'] == selector:
                        selected_price = price_info['price']
                        selected_selector = selector
                        break
                if selected_price:
                    break
            
            # セレクタ順序で見つからない場合（通常は発生しない）、最初に見つかった価格を使用
            if not selected_price:
                selected_price = found_prices[0]['price']
                selected_selector = found_prices[0]['selector']
            
            logger.debug(f"  選択した価格: {selected_price}円 (セレクタ: {selected_selector})")
            return selected_price
        
        # すべてのセレクタで見つからなかった場合、すべての価格要素を取得して最大値を返す
        return self._extract_max_price_from_all_elements(exclude_selectors)
    
    def _extract_max_price_from_all_elements(self, exclude_selectors: List[str]) -> Optional[int]:
        """
        ページ内のすべての価格要素から最大値を取得する（フォールバック）
        
        Args:
            exclude_selectors: 除外するセレクタのリスト
            
        Returns:
            Optional[int]: 抽出された最大価格、見つからない場合はNone
        """
        try:
            # IDセレクタを優先的にチェック（信頼性が高い）
            id_selectors = [
                "#itemPrice",
                "[id*='itemPrice']",
                "[id*='price']",
                "#price"
            ]
            
            all_prices = []
            
            # まずIDセレクタをチェック
            for selector in id_selectors:
                try:
                    elements = self.browser.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        try:
                            price_text = element.text.strip()
                            if not price_text:
                                continue
                            
                            price_text_lower = price_text.lower()
                            
                            # IDセレクタの場合は「送料別」のみチェック（「送料無料」は除外しない）
                            exclude_keywords = ['楽天カード', 'ポイント利用', 'special', 'offer', '送料別', '内訳']
                            if any(keyword in price_text_lower for keyword in exclude_keywords):
                                continue
                            
                            # IDセレクタの場合は親要素チェックをスキップ
                            price = self.extract_price(price_text)
                            if price and price > 0:
                                all_prices.append(price)
                        except:
                            continue
                except:
                    continue
            
            # IDセレクタで見つかった場合はそれを返す
            if all_prices:
                return max(all_prices)
            
            # IDセレクタで見つからない場合、一般的な価格セレクタで全要素を取得
            all_price_selectors = [
                "[class*='price']",
                "[class*='Price']",
                ".price",
                "#price"
            ]
            
            for selector in all_price_selectors:
                try:
                    elements = self.browser.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        try:
                            price_text = element.text.strip()
                            if not price_text:
                                continue
                            
                            price_text_lower = price_text.lower()
                            
                            # 除外キーワードチェック
                            exclude_keywords = ['楽天カード', 'ポイント利用', 'special', 'offer', '送料別', '内訳']
                            if any(keyword in price_text_lower for keyword in exclude_keywords):
                                continue
                            
                            # 親要素もチェック
                            try:
                                parent = element.find_element(By.XPATH, './..')
                                parent_text = parent.text.lower()
                                if any(keyword in parent_text for keyword in exclude_keywords):
                                    continue
                            except:
                                pass
                            
                            price = self.extract_price(price_text)
                            if price and price > 0:
                                all_prices.append(price)
                        except:
                            continue
                except:
                    continue
            
            if all_prices:
                # 最大値を返す（通常価格の方が高いことが多い）
                return max(all_prices)
        except:
            pass
        
        return None
    
    def _extract_stock_status_with_selectors(
        self, 
        selectors: List[str], 
        keywords: Dict[str, List[str]]
    ) -> str:
        """
        複数のセレクタを試行して在庫ステータスを取得する
        
        Args:
            selectors: CSSセレクタのリスト
            keywords: 在庫ステータスのキーワード辞書
            
        Returns:
            str: 在庫ステータス（"在庫あり" or "売り切れ" or "不明"）
        """
        in_stock_keywords = keywords.get('in_stock', [])
        out_of_stock_keywords = keywords.get('out_of_stock', [])
        
        # デフォルトは「在庫あり」
        stock_status = '在庫あり'
        
        for selector in selectors:
            try:
                element = self.wait_and_get_element(By.CSS_SELECTOR, selector, timeout=3)
                if element:
                    stock_text = element.text.lower()
                    
                    # 売り切れキーワードをチェック
                    for keyword in out_of_stock_keywords:
                        if keyword.lower() in stock_text:
                            return '売り切れ'
                    
                    # 在庫ありキーワードをチェック
                    for keyword in in_stock_keywords:
                        if keyword.lower() in stock_text:
                            stock_status = '在庫あり'
                            break
            except Exception:
                continue
        
        return stock_status


class ScraperConfigLoader:
    """スクレイパー設定ファイルのローダー"""
    
    def __init__(self, config_path: Optional[Path] = None, browser=None, use_spreadsheet: bool = True):
        """
        Args:
            config_path: 設定ファイルのパス（省略時はデフォルトパスを使用）
            browser: Selenium WebDriverインスタンス（スプレッドシート読み込み用）
            use_spreadsheet: スプレッドシートから設定を読み込むかどうか（デフォルト: True）
        """
        if config_path is None:
            # exe実行時と通常実行時で適切なベースディレクトリを取得
            import sys
            if getattr(sys, 'frozen', False):
                # exe実行時: sys.executableがexeファイルのパス
                base_dir = Path(sys.executable).parent.resolve()
            else:
                # 通常実行時: このファイルの親の親（inventory_scraper）をベースディレクトリとする
                base_dir = Path(__file__).parent.parent.resolve()
            config_path = base_dir / 'config' / 'scraper_config.json'
        self.config_path = config_path
        self.browser = browser
        self.use_spreadsheet = use_spreadsheet
        self.config = self._load_config()
    
    def _load_config(self) -> Dict:
        """設定ファイルを読み込む（スプレッドシート設定を優先）"""
        json_config = self._load_json_config()
        
        # スプレッドシートから設定を読み込む
        if self.use_spreadsheet and self.browser:
            try:
                from .spreadsheet_config_loader import SpreadsheetConfigLoader
                spreadsheet_loader = SpreadsheetConfigLoader(self.browser)
                spreadsheet_config = spreadsheet_loader.load_config_from_spreadsheet()
                
                # スプレッドシート設定とJSON設定をマージ（スプレッドシート設定が優先）
                merged_config = spreadsheet_loader.merge_with_json_config(
                    spreadsheet_config, 
                    self.config_path
                )
                logger.info("スプレッドシートから設定を読み込みました")
                return merged_config
            except Exception as e:
                logger.warning(f"スプレッドシートから設定を読み込めませんでした（JSON設定を使用）: {e}")
                return json_config
        else:
            return json_config
    
    def _load_json_config(self) -> Dict:
        """JSON設定ファイルを読み込む"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            logger.warning(f"設定ファイルが見つかりません: {self.config_path}")
            return {"sites": [], "default": {}}
        except json.JSONDecodeError as e:
            logger.warning(f"設定ファイルのJSON解析に失敗しました: {e}")
            return {"sites": [], "default": {}}
    
    def find_site_config(self, url: str) -> Optional[Dict]:
        """
        URLに基づいて該当するサイト設定を検索する
        
        Args:
            url: スクレイピング対象のURL
            
        Returns:
            Optional[Dict]: サイト設定、見つからない場合はNone
        """
        url_lower = url.lower()
        
        for site_config in self.config.get('sites', []):
            url_patterns = site_config.get('url_patterns', [])
            for pattern in url_patterns:
                if pattern.lower() in url_lower:
                    return site_config
        
        return None
    
    def get_default_config(self) -> Dict:
        """デフォルト設定を取得する"""
        return self.config.get('default', {})


def create_configurable_scraper(url: str, browser, config_loader: ScraperConfigLoader) -> BaseScraper:
    """
    設定ファイルに基づいてスクレイパーを作成する
    
    Args:
        url: スクレイピング対象のURL
        browser: Selenium WebDriverインスタンス
        config_loader: 設定ローダー
        
    Returns:
        BaseScraper: 設定ファイルベースのスクレイパーまたはデフォルトスクレイパー
    """
    site_config = config_loader.find_site_config(url)
    
    if site_config:
        return ConfigurableScraper(browser, site_config)
    else:
        # デフォルト設定を使用
        default_config = config_loader.get_default_config()
        if default_config:
            return ConfigurableScraper(browser, default_config)
        else:
            # 設定ファイルがない場合は、既存のAmazonスクレイパーを使用
            from .scraper import AmazonScraper
            return AmazonScraper(browser)

