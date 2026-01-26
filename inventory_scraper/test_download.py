# test_download.py
from src.browser import create_browser
from src.downloader import download_spreadsheet_csv

print("ブラウザを起動しています...")
browser = create_browser()

try:
    print("スプレッドシートからCSVをダウンロードしています...")
    df = download_spreadsheet_csv(browser)
    print(f"ダウンロード成功: {len(df)}行のデータを取得しました")
    print(df.head())
finally:
    browser.quit()