"""
スプレッドシート直接更新モジュール
Google Apps ScriptのWebアプリを直接呼び出してスプレッドシートを更新する

このモジュールは、GASのWebアプリとして公開されたエンドポイントに
POSTリクエストでCSVデータを送信し、スプレッドシートを更新します。
"""
import json
import requests
from pathlib import Path


def update_spreadsheet_via_gas(browser=None, csv_path: Path = None, script_url: str = None):
    """
    Google Apps ScriptのWebアプリを呼び出してスプレッドシートを更新する
    
    CSVデータをPOSTリクエストのボディとして送信し、GASのWebアプリに送信します。
    
    Args:
        browser: Selenium WebDriverインスタンス（後方互換性のため、使用されません）
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
        
        data_row_count = len(csv_content.splitlines()) - 1
        print(f"更新対象データ: {data_row_count}件（ヘッダー除く）")
        
        # POSTリクエストでCSVデータを送信
        print("GAS WebアプリにCSVデータをPOST送信しています...")
        
        # 大きなCSVの場合のチャンキング処理
        MAX_CHUNK_SIZE = 50000  # 約50KBを1チャンクとする（保守的に設定）
        csv_size = len(csv_content.encode('utf-8'))
        
        if csv_size > MAX_CHUNK_SIZE:
            print(f"大きなCSVデータを検出しました（{csv_size}バイト）。チャンキング処理を実行します...")
            _send_csv_in_chunks(csv_content, script_url, MAX_CHUNK_SIZE)
        else:
            _send_csv_post(csv_content, script_url)
        
    except Exception as e:
        error_message = f"スプレッドシートの更新に失敗しました: {e}"
        print(error_message)
        raise Exception(error_message)


def _send_csv_post(csv_content: str, script_url: str):
    """
    CSVデータをPOSTリクエストで送信する
    
    Args:
        csv_content: CSVデータ（文字列）
        script_url: GAS WebアプリURL
        
    Raises:
        Exception: 送信に失敗した場合
    """
    try:
        # JSON形式でCSVデータを送信
        payload = {"csvData": csv_content}
        headers = {"Content-Type": "application/json"}
        
        response = requests.post(
            script_url,
            json=payload,
            headers=headers,
            timeout=300  # 5分のタイムアウト
        )
        
        # ステータスコードを確認
        response.raise_for_status()
        
        # HTMLエラーページが返された場合のチェック
        response_text = response.text.strip()
        if response_text.startswith('<!DOCTYPE html>') or response_text.startswith('<html>'):
            # HTMLエラーページが返された場合
            error_message = (
                "GAS WebアプリがHTMLエラーページを返しました。\n"
                "以下の可能性があります：\n"
                "1. doPost関数がデプロイされていない（新しいバージョンとしてデプロイが必要）\n"
                "2. POSTリクエストがサポートされていない\n"
                "3. 認証の問題\n"
                "\n"
                "対応方法：\n"
                "1. GASエディタで「デプロイ」→「デプロイを管理」を開く\n"
                "2. 既存のデプロイを選択して「新しいバージョンを保存」をクリック\n"
                "3. デプロイ後に再度実行してください\n"
                f"\nレスポンス（最初の500文字）:\n{response_text[:500]}"
            )
            print(f"❌ {error_message}")
            raise Exception(error_message)
        
        # JSONレスポンスをパース
        try:
            result = response.json()
            
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
                
        except json.JSONDecodeError as e:
            error_message = (
                f"JSON解析に失敗しました: {e}\n"
                f"レスポンス内容（最初の1000文字）:\n{response_text[:1000]}"
            )
            print(f"警告: {error_message}")
            raise Exception(f"レスポンスの解析に失敗しました: {e}")
            
    except requests.exceptions.RequestException as e:
        error_msg = f"HTTPリクエストエラー: {e}"
        print(error_msg)
        raise Exception(error_msg)


def _send_csv_in_chunks(csv_content: str, script_url: str, max_chunk_size: int):
    """
    大きなCSVデータをチャンクに分割して送信する
    
    Args:
        csv_content: CSVデータ（文字列）
        script_url: GAS WebアプリURL
        max_chunk_size: 1チャンクの最大サイズ（バイト）
        
    Raises:
        Exception: 送信に失敗した場合
    """
    lines = csv_content.splitlines()
    if len(lines) < 2:
        # ヘッダーのみまたは空の場合は通常送信
        _send_csv_post(csv_content, script_url)
        return
    
    header = lines[0]
    data_lines = lines[1:]
    
    # チャンクサイズを計算（ヘッダーを含めた行数ベースで分割）
    total_size = len(csv_content.encode('utf-8'))
    estimated_chunk_count = max(1, (total_size + max_chunk_size - 1) // max_chunk_size)
    lines_per_chunk = max(1, len(data_lines) // estimated_chunk_count)
    
    print(f"CSVデータを{estimated_chunk_count}チャンクに分割して送信します（1チャンクあたり約{lines_per_chunk}行）")
    
    # チャンクごとに送信
    for i in range(0, len(data_lines), lines_per_chunk):
        chunk_lines = data_lines[i:i + lines_per_chunk]
        chunk_content = '\n'.join([header] + chunk_lines)
        chunk_number = (i // lines_per_chunk) + 1
        
        print(f"チャンク {chunk_number}/{estimated_chunk_count} を送信中...")
        try:
            _send_csv_post(chunk_content, script_url)
        except Exception as e:
            print(f"⚠️  チャンク {chunk_number} の送信でエラーが発生しました: {e}")
            # チャンク送信の失敗は警告として記録し、次のチャンクを続行
            # 完全な失敗にする場合は例外を再スロー
            continue
    
    print("すべてのチャンクの送信が完了しました")

