"""
スプレッドシート直接更新モジュール
Google Apps ScriptのWebアプリを直接呼び出してスプレッドシートを更新する

このモジュールは、GASのWebアプリとして公開されたエンドポイントに
URLパラメータ経由でCSVデータを送信し、スプレッドシートを更新します。
"""
import json
import time
import urllib.parse
from pathlib import Path


def update_spreadsheet_via_gas(browser, csv_path: Path, script_url: str = None):
    """
    Google Apps ScriptのWebアプリを呼び出してスプレッドシートを更新する
    
    CSVデータをURLパラメータとしてエンコードし、GASのWebアプリに送信します。
    
    Args:
        browser: Selenium WebDriverインスタンス
        csv_path: 更新データが含まれるCSVファイルのパス
        script_url: Google Apps ScriptのWebアプリURL（必須）
        
    Raises:
        Exception: 更新に失敗した場合
    """
    try:
        # Google Apps ScriptのWebアプリURLを確認
        if not script_url:
            raise Exception("Google Apps ScriptのWebアプリURLが指定されていません。.envファイルにGAS_WEB_APP_URLを設定してください。")
        
        # CSVファイルを読み込む
        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            csv_content = f.read()
        
        if not csv_content or len(csv_content.strip()) == 0:
            print("更新するデータがありません")
            return
        
        print(f"更新対象データ: {len(csv_content.splitlines()) - 1}件（ヘッダー除く）")
        
        # URLパラメータ経由でCSVデータを送信
        print("GAS WebアプリにCSVデータを送信しています...")
        encoded_csv = urllib.parse.quote(csv_content)
        full_url = f"{script_url}?csvData={encoded_csv}"
        print(f"Google Apps Scriptを呼び出します: {full_url[:100]}...")
        
        browser.get(full_url)
        time.sleep(5)
        
        # GAS Webアプリの処理完了後、ブラウザタブをクリーンアップ
        try:
            browser.get('about:blank')
        except:
            pass
        
        # レスポンスを確認（JSON形式で返される）
        page_text = browser.page_source
        
        # JSONレスポンスをパース
        try:
            # HTMLタグを除去してJSONを抽出
            json_start = page_text.find('{')
            json_end = page_text.rfind('}') + 1
            if json_start >= 0 and json_end > json_start:
                json_text = page_text[json_start:json_end]
                result = json.loads(json_text)
                
                if result.get('success'):
                    print(f"✅ スプレッドシートの更新が完了しました")
                    print(f"   - 更新行数: {result.get('updateCount', 0)}行")
                    print(f"   - 仕入れ価格更新: {result.get('priceUpdateCount', 0)}件")
                    print(f"   - 在庫ステータス更新: {result.get('statusUpdateCount', 0)}件")
                    print(f"   - 最終更新日時更新: {result.get('dateUpdateCount', 0)}件")
                    if result.get('notFoundCount', 0) > 0:
                        print(f"   ⚠️  マッチしなかったURL: {result.get('notFoundCount', 0)}件")
                else:
                    error_msg = result.get('error', '不明なエラー')
                    print(f"❌ スプレッドシートの更新に失敗しました: {error_msg}")
                    raise Exception(f"更新失敗: {error_msg}")
            else:
                # JSONが見つからない場合、テキストで確認
                if 'success' in page_text.lower() or '更新' in page_text:
                    print("スプレッドシートの更新が完了しました（JSON解析失敗）")
                else:
                    print("警告: 更新結果の確認ができませんでした")
                    print(f"ページ内容: {page_text[:500]}")
        except json.JSONDecodeError as e:
            print(f"警告: JSON解析に失敗しました: {e}")
            print(f"ページ内容: {page_text[:500]}")
            # JSON解析に失敗しても、ページにsuccessが含まれていれば成功とみなす
            if 'success' in page_text.lower():
                print("スプレッドシートの更新が完了しました（JSON解析失敗）")
        
    except Exception as e:
        error_message = f"スプレッドシートの更新に失敗しました: {e}"
        print(error_message)
        raise Exception(error_message)

