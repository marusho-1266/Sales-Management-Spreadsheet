"""
Googleフォームアップロードモジュール
結果CSVをGoogleフォーム経由で送信する
"""
import time
import logging
import pandas as pd
from pathlib import Path
from typing import Dict, Optional
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    NoSuchFrameException,
    WebDriverException
)
from .config import GOOGLE_FORM_URL, DATA_DIR

# ロガーを設定
logger = logging.getLogger(__name__)


def save_result_csv(df: pd.DataFrame, filename: str = 'upload_data.csv') -> Path:
    """
    スクレイピング結果をCSVファイルとして保存する
    
    Args:
        df: スクレイピング結果のDataFrame
        filename: 保存するファイル名（デフォルト: upload_data.csv）
    
    Returns:
        Path: 保存されたCSVファイルのパス
    
    Raises:
        Exception: CSVファイルの保存に失敗した場合
    """
    data_dir = Path(DATA_DIR)
    csv_path = data_dir / filename
    
    # 親ディレクトリが存在することを確認（存在しない場合は作成）
    data_dir.mkdir(parents=True, exist_ok=True)
    
    # UTF-8 BOM付きで保存（Excel互換性のため）
    try:
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    except Exception as e:
        error_message = f"CSVファイルの保存に失敗しました: {csv_path}, エラー: {e}"
        logger.error(error_message)
        raise Exception(error_message) from e
    
    print(f"CSVファイルを保存しました: {csv_path}")
    return csv_path


def upload_to_google_form(browser, csv_path: Path) -> Dict[str, str]:
    """
    GoogleフォームにCSVファイルをアップロードして送信する
    
    Args:
        browser: Selenium WebDriverインスタンス
        csv_path: アップロードするCSVファイルのパス
    
    Returns:
        Dict[str, str]: 送信結果を表す辞書
            - status: "success" または "unconfirmed"
            - message: 結果メッセージ
            - url: 送信後のURL（unconfirmedの場合）
            - title: 送信後のページタイトル（unconfirmedの場合）
    
    Raises:
        Exception: アップロードに失敗した場合
    """
    try:
        # CSVファイルのパスを絶対パスに変換
        absolute_csv_path = csv_path.resolve()
        
        # ファイルの存在確認（ブラウザ操作の前にチェック）
        if not absolute_csv_path.exists():
            raise Exception(f"CSVファイルが見つかりません: {absolute_csv_path}")
        
        # Googleフォームを開く
        print(f"Googleフォームを開きます: {GOOGLE_FORM_URL}")
        browser.get(GOOGLE_FORM_URL)
        print("ページの読み込みを待機しています...")
        time.sleep(5)  # ページ読み込み待機（少し長めに）
        
        # ページのタイトルを確認
        print(f"ページタイトル: {browser.title}")
        
        # ページが完全に読み込まれるまで待機（GoogleフォームはJavaScriptで動的に要素を生成する）
        print("ページの完全な読み込みを待機しています...")
        time.sleep(5)  # 追加の待機時間
        
        file_input = None
        
        # まず、直接ファイル入力要素を探す（ボタンクリック前に存在する可能性がある）
        print("直接ファイル入力要素を探しています...")
        try:
            browser.switch_to.default_content()
            file_input = WebDriverWait(browser, 3).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
            )
            if file_input:
                print("ファイル入力要素を直接見つけました（ボタンクリック不要）")
        except TimeoutException:
            print("直接ファイル入力要素は見つかりませんでした。ボタンクリックが必要です。")
            file_input = None
        
        # 直接見つからない場合、ボタンを探してクリックする
        if not file_input:
            # Googleフォームのファイルアップロードは、まず「ファイルを追加」ボタンをクリックする必要がある
            print("「ファイルを追加」ボタンを探しています...")
        
        # 「ファイルを追加」ボタンのセレクタ（日本語と英語の両方に対応）
        add_file_button_selectors = [
            "//span[contains(text(), 'ファイルを追加')]",
            "//span[contains(text(), 'Add file')]",
            "//button[contains(., 'ファイルを追加')]",
            "//button[contains(., 'Add file')]",
            "//div[contains(@aria-label, 'ファイルを追加')]",
            "//div[contains(@aria-label, 'Add file')]",
            "//span[contains(@aria-label, 'ファイルを追加')]",
            "//span[contains(@aria-label, 'Add file')]",
            "//*[@role='button' and contains(., 'ファイル')]",
            "//*[@role='button' and contains(., 'file')]"
        ]
        
        add_file_button = None
        for selector in add_file_button_selectors:
            try:
                print(f"  ボタンセレクタを試行中: {selector}")
                add_file_button = WebDriverWait(browser, 5).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                if add_file_button:
                    button_text = add_file_button.text
                    print(f"  「ファイルを追加」ボタンを見つけました: {button_text}")
                    break
            except TimeoutException:
                continue
        
        # ボタンが見つからない場合、role="button"の要素をすべて確認
        if not add_file_button:
            print("標準セレクタで見つからなかったため、role='button'の要素を確認します...")
            try:
                buttons = browser.find_elements(By.XPATH, "//*[@role='button']")
                print(f"  見つかったボタンの数: {len(buttons)}")
                for i, button in enumerate(buttons):
                    button_text = button.text.lower()
                    aria_label = (button.get_attribute('aria-label') or '').lower()
                    print(f"  Button[{i}]: text={button.text[:50]}, aria-label={aria_label[:50]}")
                    # 「ファイルを追加」または「Add file」を含むボタンを探す
                    # ただし、「ファイルを開く」や「削除」などの既存ファイル操作ボタンは除外
                    if (('ファイルを追加' in button_text or 'add file' in button_text or 
                         'ファイルを追加' in aria_label or 'add file' in aria_label) and
                        'ファイルを開く' not in button_text and 'open' not in button_text and
                        '削除' not in button_text and 'delete' not in button_text and
                        'remove' not in button_text):
                        add_file_button = button
                        print(f"  「ファイルを追加」ボタンを見つけました（全ボタン検索）")
                        break
            except Exception as e:
                print(f"  ボタン検索中にエラー: {e}")
        
        # 既にファイルがアップロードされている場合は削除する
        # 「削除」ボタンや「×」ボタンを探す（ファイル入力要素が見つからない場合、または既存ファイルがある場合）
        delete_button = None
        if not file_input or not add_file_button:
            print("既にファイルがアップロードされている可能性があります。削除ボタンを探します...")
            try:
                delete_selectors = [
                    "//button[contains(@aria-label, '削除')]",
                    "//button[contains(@aria-label, 'Delete')]",
                    "//button[contains(@aria-label, 'Remove')]",
                    "//*[@role='button' and contains(@aria-label, '削除')]",
                    "//*[@role='button' and contains(@aria-label, 'Delete')]",
                    "//*[@role='button' and contains(@aria-label, 'Remove')]",
                    "//button[contains(., '×')]",
                    "//*[@role='button' and contains(., '×')]",
                    "//*[@role='button' and contains(@aria-label, 'ファイルを削除')]",
                    "//*[@role='button' and contains(@aria-label, 'Remove file')]"
                ]
                for selector in delete_selectors:
                    try:
                        delete_button = WebDriverWait(browser, 2).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        if delete_button:
                            print(f"削除ボタンを見つけました（セレクタ: {selector}）。既存のファイルを削除します...")
                            browser.execute_script("arguments[0].click();", delete_button)
                            time.sleep(3)  # 削除処理の待機（長めに）
                            print("既存のファイルを削除しました")
                            # 削除後、再度「ファイルを追加」ボタンとファイル入力要素を探す
                            time.sleep(2)  # ページ更新待機
                            break
                    except TimeoutException:
                        continue
                
                # セレクタで見つからない場合、role="button"の要素をすべて確認
                if not delete_button:
                    print("標準セレクタで削除ボタンが見つからなかったため、role='button'の要素を確認します...")
                    try:
                        buttons = browser.find_elements(By.XPATH, "//*[@role='button']")
                        for i, button in enumerate(buttons):
                            try:
                                button_text = button.text.lower()
                                aria_label = (button.get_attribute('aria-label') or '').lower()
                                # 削除関連のボタンを探す
                                if (('削除' in button_text or 'delete' in button_text or 'remove' in button_text or
                                     '×' in button_text or 'x' in button_text) and
                                    ('ファイルを開く' not in button_text and 'open' not in button_text)):
                                    delete_button = button
                                    print(f"削除ボタンを見つけました（全ボタン検索）: text={button.text[:50]}, aria-label={aria_label[:50]}")
                                    browser.execute_script("arguments[0].click();", delete_button)
                                    time.sleep(3)  # 削除処理の待機
                                    print("既存のファイルを削除しました")
                                    time.sleep(2)  # ページ更新待機
                                    break
                            except Exception as e:
                                continue
                    except Exception as e:
                        print(f"削除ボタンの全ボタン検索中にエラー（無視して続行）: {e}")
            except Exception as e:
                print(f"削除ボタンの検索中にエラー（無視して続行）: {e}")
        
        # 削除後、再度「ファイルを追加」ボタンとファイル入力要素を探す
        if delete_button:
            print("削除後、「ファイルを追加」ボタンとファイル入力要素を再度探します...")
            time.sleep(2)  # ページ更新待機
            
            # まずファイル入力要素を探す
            try:
                browser.switch_to.default_content()
                file_input = WebDriverWait(browser, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
                )
                if file_input:
                    print("削除後、ファイル入力要素を見つけました")
            except TimeoutException:
                pass
            
            # ファイル入力要素が見つからない場合、「ファイルを追加」ボタンを探す
            if not file_input:
                for selector in add_file_button_selectors:
                    try:
                        add_file_button = WebDriverWait(browser, 5).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        if add_file_button:
                            button_text = add_file_button.text
                            print(f"  「ファイルを追加」ボタンを見つけました: {button_text}")
                            break
                    except TimeoutException:
                        continue
        
        # まだファイル入力要素が見つかっていない場合、ボタンをクリックしてファイルピッカーを開く
        if not file_input:
            if add_file_button:
                try:
                    print("「ファイルを追加」ボタンをクリックします...")
                    browser.execute_script("arguments[0].click();", add_file_button)
                    time.sleep(5)  # ファイルピッカーが開くまで待機（長めに）
                    print("ファイルピッカーが開きました")
                    
                    # ボタンクリック後、再度ファイル入力要素を探す
                    print("ボタンクリック後、ファイル入力要素を再度探します...")
                    try:
                        browser.switch_to.default_content()
                        file_input = WebDriverWait(browser, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
                        )
                        if file_input:
                            print("ボタンクリック後、ファイル入力要素を見つけました")
                    except TimeoutException:
                        pass
                except Exception as e:
                    print(f"ボタンクリック中にエラー: {e}")
                    raise Exception(f"「ファイルを追加」ボタンのクリックに失敗しました: {e}")
            else:
                # ボタンが見つからない場合、直接ファイル入力要素を探す
                print("「ファイルを追加」ボタンが見つかりませんでした。直接ファイル入力要素を探します...")
        
        # ファイルピッカーのiframeを探す
        # Googleフォームのファイルピッカーは、ボタンクリック後にダイアログとして開かれる
        print("ファイルピッカーのiframeを探しています...")
        time.sleep(2)  # iframeが表示されるまで待機
        
        # まず、すべてのiframeを確認
        all_iframes = browser.find_elements(By.TAG_NAME, 'iframe')
        print(f"  見つかったiframeの数: {len(all_iframes)}")
        
        file_picker_iframe = None
        iframe_selectors = [
            'iframe.picker-frame',
            'iframe[class*="picker"]',
            'iframe[src*="picker"]',
            'iframe[src*="drive.google.com"]',
            'iframe[src*="docs.google.com"]'
        ]
        
        # セレクタで探す
        for selector in iframe_selectors:
            try:
                print(f"  iframeセレクタを試行中: {selector}")
                file_picker_iframe = WebDriverWait(browser, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                if file_picker_iframe:
                    iframe_src = file_picker_iframe.get_attribute('src')
                    print(f"  iframeを見つけました: {selector}, src={iframe_src[:100] if iframe_src else 'None'}")
                    break
            except TimeoutException:
                continue
        
        # セレクタで見つからない場合、すべてのiframeを確認
        if not file_picker_iframe and len(all_iframes) > 0:
            print("  セレクタで見つからなかったため、すべてのiframeを確認します...")
            for i, iframe in enumerate(all_iframes):
                try:
                    iframe_src = iframe.get_attribute('src') or ''
                    iframe_class = iframe.get_attribute('class') or ''
                    print(f"  Iframe[{i}]: src={iframe_src[:100]}, class={iframe_class[:50]}")
                    # picker関連のiframeを探す
                    if ('picker' in iframe_src.lower() or 'picker' in iframe_class.lower() or
                        'drive.google.com' in iframe_src.lower() or 'docs.google.com' in iframe_src.lower()):
                        file_picker_iframe = iframe
                        print(f"  ファイルピッカーiframeを見つけました（全iframe検索）")
                        break
                except Exception as e:
                    print(f"  Iframe[{i}]の確認中にエラー: {e}")
                    continue
        
        # iframeが見つかった場合、iframe内でファイル入力要素を探す
        in_iframe = False
        if file_picker_iframe:
            try:
                print("ファイルピッカーのiframeに切り替えます...")
                browser.switch_to.frame(file_picker_iframe)
                in_iframe = True
                time.sleep(2)  # iframe内の読み込み待機
                
                # iframe内でファイル入力要素を探す
                print("iframe内でファイル入力要素を探しています...")
                file_input = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
                )
                if file_input:
                    print("iframe内でファイル入力要素を見つけました")
            except TimeoutException:
                print("iframe内でファイル入力要素が見つかりませんでした")
                browser.switch_to.default_content()
                in_iframe = False
            except Exception as e:
                print(f"iframe処理中にエラー: {e}")
                browser.switch_to.default_content()
                in_iframe = False
        
        # iframe内で見つからなかった場合、メインフレームで探す
        if not file_input:
            print("メインフレームでファイル入力要素を探しています...")
            try:
                browser.switch_to.default_content()
                in_iframe = False
                # 少し待機してから探す（ダイアログが開くまで）
                time.sleep(2)
                file_input = WebDriverWait(browser, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
                )
                if file_input:
                    print("メインフレームでファイル入力要素を見つけました")
            except TimeoutException:
                pass
        
        # まだ見つからない場合、すべてのinput要素を確認（iframe内も含む）
        if not file_input:
            print("標準セレクタで見つからなかったため、すべてのinput要素を確認します...")
            try:
                # まずメインフレームで確認
                browser.switch_to.default_content()
                in_iframe = False
                all_inputs = browser.find_elements(By.TAG_NAME, 'input')
                print(f"  メインフレームで見つかったinput要素の数: {len(all_inputs)}")
                for i, inp in enumerate(all_inputs):
                    input_type = inp.get_attribute('type')
                    input_class = inp.get_attribute('class')
                    input_id = inp.get_attribute('id')
                    aria_label = inp.get_attribute('aria-label')
                    print(f"  Input[{i}]: type={input_type}, class={input_class}, id={input_id}, aria-label={aria_label}")
                    if input_type == 'file':
                        file_input = inp
                        print(f"  ファイル入力要素を見つけました（メインフレーム全要素検索）")
                        break
                
                # メインフレームで見つからない場合、iframe内も確認
                if not file_input and file_picker_iframe:
                    try:
                        print("  iframe内でinput要素を確認します...")
                        browser.switch_to.frame(file_picker_iframe)
                        in_iframe = True
                        time.sleep(1)
                        iframe_inputs = browser.find_elements(By.TAG_NAME, 'input')
                        print(f"  iframe内で見つかったinput要素の数: {len(iframe_inputs)}")
                        for i, inp in enumerate(iframe_inputs):
                            input_type = inp.get_attribute('type')
                            if input_type == 'file':
                                file_input = inp
                                print(f"  ファイル入力要素を見つけました（iframe内全要素検索）")
                                break
                    except Exception as e:
                        print(f"  iframe内の検索中にエラー: {e}")
                        browser.switch_to.default_content()
                        in_iframe = False
            except Exception as e:
                print(f"  全要素検索中にエラー: {e}")
                try:
                    browser.switch_to.default_content()
                    in_iframe = False
                except (NoSuchFrameException, WebDriverException) as e:
                    # 予期されるSelenium例外: ログに記録して続行
                    logger.warning(f"メインフレームへの復帰中にSelenium例外が発生しました: {e}")
                except Exception as e:
                    # 予期しない例外: 再発生させる
                    logger.error(f"メインフレームへの復帰中に予期しない例外が発生しました: {e}")
                    raise
        
        if not file_input:
            # メインフレームに戻る
            try:
                browser.switch_to.default_content()
            except (NoSuchFrameException, WebDriverException) as e:
                # 予期されるSelenium例外: ログに記録して続行
                logger.warning(f"メインフレームへの復帰中にSelenium例外が発生しました: {e}")
            except Exception as e:
                # 予期しない例外: 再発生させる
                logger.error(f"メインフレームへの復帰中に予期しない例外が発生しました: {e}")
                raise
            
            # デバッグ情報: より詳細なHTML構造を出力
            print("ファイル入力要素が見つかりませんでした。ページのHTML構造を確認します...")
            # ファイルアップロード関連の要素を探す
            upload_elements = browser.find_elements(By.XPATH, "//*[contains(text(), 'ファイル') or contains(text(), 'file') or contains(text(), 'upload')]")
            print(f"  ファイル関連のテキストを含む要素の数: {len(upload_elements)}")
            for i, elem in enumerate(upload_elements[:10]):  # 最初の10個のみ表示
                print(f"  Element[{i}]: tag={elem.tag_name}, text={elem.text[:50]}")
            
            # より詳細なデバッグ情報を出力
            print("\n=== デバッグ情報 ===")
            print("ファイル入力要素が見つかりませんでした。")
            print("\n確認事項:")
            print("1. フォーム編集画面で「ファイルアップロード」項目が追加されているか確認してください")
            print("2. フォーム編集URL: フォームURLの'/viewform'を'/edit'に変更してアクセス")
            print("3. フォームに「ファイルアップロード」項目が表示されているか確認")
            print("\nファイルアップロード項目の追加手順:")
            print("- フォーム編集画面で「+」ボタンをクリック")
            print("- 「ファイルアップロード」を選択")
            print("- 質問タイトルを設定")
            print("- ファイルの種類を「CSV」に設定（オプション）")
            print("- 最大ファイル数を「1」に設定")
            print("- 必須項目にチェックを入れる")
            print("- 右上の「保存」ボタンをクリック")
            
            # ページソースの一部を出力（デバッグ用）
            page_source_snippet = browser.page_source[:3000]  # 最初の3000文字
            print(f"\nページソース（最初の3000文字）:\n{page_source_snippet}")
            
            raise Exception("ファイル入力要素が見つかりません。フォームにファイルアップロード項目が追加されているか確認してください。フォーム編集画面で「ファイルアップロード」項目を追加してください。")
        
        print(f"CSVファイルをアップロードします: {absolute_csv_path}")
        print(f"ファイルサイズ: {absolute_csv_path.stat().st_size} bytes")
        
        # ファイル入力要素が表示されているか確認（iframe内の場合はスキップ）
        try:
            if not in_iframe and not file_input.is_displayed():
                print("警告: ファイル入力要素が表示されていません。JavaScriptで表示する必要があるかもしれません。")
                # JavaScriptで表示を試みる
                browser.execute_script("arguments[0].style.display = 'block';", file_input)
                browser.execute_script("arguments[0].style.visibility = 'visible';", file_input)
                time.sleep(1)
        except Exception as e:
            print(f"表示確認中にエラー（無視して続行）: {e}")
        
        # ファイルをアップロード（iframe内で見つけた場合は、iframe内で実行）
        try:
            print("ファイルパスを送信します...")
            file_input.send_keys(str(absolute_csv_path))
            print("ファイルのアップロードを開始しました...")
            time.sleep(5)  # アップロード処理待機（少し長めに）
            
            # メインフレームに戻る（ファイルアップロード後）
            if in_iframe:
                print("メインフレームに戻ります...")
                browser.switch_to.default_content()
                time.sleep(2)  # メインフレームの読み込み待機
        except Exception as e:
            # エラーが発生した場合もメインフレームに戻る
            if in_iframe:
                try:
                    browser.switch_to.default_content()
                except (NoSuchFrameException, WebDriverException) as frame_e:
                    # 予期されるSelenium例外: ログに記録して続行
                    logger.warning(f"メインフレームへの復帰中にSelenium例外が発生しました: {frame_e}")
                except Exception as unexpected_e:
                    # 予期しない例外: ログに記録するが、元の例外を優先して再発生させない
                    # （このブロックは既に例外ハンドリング内にあるため）
                    logger.error(f"メインフレームへの復帰中に予期しない例外が発生しました: {unexpected_e}")
            print(f"ファイルアップロード中にエラーが発生しました: {e}")
            raise Exception(f"ファイルのアップロードに失敗しました: {e}")
        
        # 送信ボタンを探す（複数のセレクタを試行）
        # 注意: CSSセレクタでは:contains()が使えないため、XPathを使用
        submit_selectors = [
            '//button[@type="submit"]',
            '//span[contains(text(), "送信")]',
            '//span[contains(text(), "Submit")]',
            '//div[@role="button" and contains(., "送信")]',
            '//div[@role="button" and contains(., "Submit")]',
            '//*[@data-testid="submit-button"]',
            '//button[contains(., "送信")]',
            '//button[contains(., "Submit")]'
        ]
        
        submit_button = None
        for selector in submit_selectors:
            try:
                print(f"  送信ボタンセレクタを試行中: {selector}")
                submit_button = WebDriverWait(browser, 5).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                if submit_button:
                    button_text = submit_button.text
                    print(f"  送信ボタンを見つけました: {button_text}")
                    break
            except TimeoutException:
                continue
            except Exception as e:
                print(f"  セレクタエラー（無視して続行）: {e}")
                continue
        
        # 送信ボタンが見つからない場合、role="button"の要素を探す
        if not submit_button:
            print("標準セレクタで見つからなかったため、role='button'の要素を確認します...")
            try:
                buttons = browser.find_elements(By.XPATH, "//*[@role='button']")
                print(f"  見つかったボタンの数: {len(buttons)}")
                for i, button in enumerate(buttons):
                    try:
                        button_text = button.text.lower()
                        aria_label = (button.get_attribute('aria-label') or '').lower()
                        print(f"  Button[{i}]: text={button.text[:50]}, aria-label={aria_label[:50]}")
                        if '送信' in button_text or 'submit' in button_text or '送信' in aria_label or 'submit' in aria_label:
                            submit_button = button
                            print(f"  送信ボタンを見つけました（全ボタン検索）")
                            break
                    except Exception as e:
                        print(f"  ボタン[{i}]の処理中にエラー: {e}")
                        continue
            except Exception as e:
                print(f"  ボタン検索中にエラー: {e}")
        
        if not submit_button:
            raise Exception("送信ボタンが見つかりません")
        
        # 送信ボタンをクリック
        print("送信ボタンをクリックします")
        browser.execute_script("arguments[0].click();", submit_button)
        time.sleep(5)  # 送信処理待機（長めに）
        
        # 送信完了画面を確認
        try:
            # ページが更新されるまで待機
            time.sleep(2)
            
            # 送信完了を示す要素を探す（より多くのパターンを追加）
            success_indicators = [
                '送信しました',
                '送信が完了しました',
                '回答を記録しました',
                'Thank you',
                'Your response has been recorded',
                'Response recorded',
                '回答が記録されました',
                'ありがとうございました'
            ]
            
            # ページのタイトルとテキストを確認
            page_title = browser.title.lower()
            page_text = browser.page_source.lower()
            
            # タイトルで確認
            if 'thank' in page_title or 'ありがとう' in page_title:
                print("送信が完了しました（タイトルで確認）")
                return {"status": "success", "message": "送信が完了しました（タイトルで確認）"}
            
            # ページテキストで確認
            for indicator in success_indicators:
                if indicator.lower() in page_text:
                    print(f"送信が完了しました（'{indicator}'を検出）")
                    return {"status": "success", "message": f"送信が完了しました（'{indicator}'を検出）"}
            
            # URLが変更されたか確認（送信後、URLが変わる可能性がある）
            current_url = browser.current_url.lower()
            if 'formresponse' in current_url or 'thank' in current_url:
                print("送信が完了しました（URL変更で確認）")
                return {"status": "success", "message": "送信が完了しました（URL変更で確認）"}
            
            # 送信完了の確認ができない場合、警告ログを出力してunconfirmed状態を返す
            current_url_full = browser.current_url
            current_title = browser.title
            logger.warning(
                "送信完了の確認メッセージが検出されませんでした。送信状態が不明です。 "
                f"url={current_url_full}, title={current_title}"
            )
            return {
                "status": "unconfirmed",
                "message": "confirmation message not detected",
                "url": current_url_full,
                "title": current_title
            }
            
        except Exception as e:
            logger.warning(f"送信完了の確認中にエラーが発生しました: {e}")
            # エラーが発生した場合もunconfirmed状態を返す
            return {
                "status": "unconfirmed",
                "message": f"confirmation check failed: {str(e)}",
                "url": browser.current_url if hasattr(browser, 'current_url') else "unknown",
                "title": browser.title if hasattr(browser, 'title') else "unknown"
            }
        
    except Exception as e:
        error_message = f"Googleフォームへのアップロードに失敗しました: {e}"
        print(error_message)
        raise Exception(error_message) from e
