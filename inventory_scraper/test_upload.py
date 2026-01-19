# test_upload.py
from src.browser import create_browser
from src.uploader import upload_to_google_form
from pathlib import Path

print("ブラウザを起動しています...")
browser = create_browser()

try:
    csv_path = Path("data/test.csv")  # テスト用CSVファイルのパス
    print("Googleフォームにアップロードしています...")
    upload_to_google_form(browser, csv_path)
    print("アップロード成功")
finally:
    browser.quit()