# test_browser.py
from src.browser import create_browser

print("ブラウザを起動しています...")
browser = create_browser()
print("ブラウザが正常に起動しました")
browser.get("https://www.google.com")
print("Googleにアクセスしました")
browser.quit()
print("ブラウザを閉じました")