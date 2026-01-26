"""
ブラウザ初期化モジュール
Selenium WebDriverのインスタンスを生成する
"""
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from .config import CHROME_PROFILE_PATH, CHROME_PROFILE_NAME, DATA_DIR, CHROME_USER_AGENT


def create_browser():
    """
    Selenium WebDriverのインスタンスを生成する
    
    Returns:
        webdriver.Chrome: Chrome WebDriverのインスタンス
    """
    chrome_options = Options()
    
    # 既存のChromeプロファイルを使用（Googleログイン状態の維持）
    if CHROME_PROFILE_PATH and CHROME_PROFILE_NAME:
        user_data_dir = os.path.join(CHROME_PROFILE_PATH, CHROME_PROFILE_NAME)
        chrome_options.add_argument(f'--user-data-dir={user_data_dir}')
        chrome_options.add_argument(f'--profile-directory={CHROME_PROFILE_NAME}')
    
    # Bot検知回避のためのオプション
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # その他のオプション
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    
    # ダウンロード設定
    # DATA_DIRを絶対パスに変換して設定
    download_dir = str(DATA_DIR.resolve())
    prefs = {
        'download.default_directory': download_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True
    }
    chrome_options.add_experimental_option('prefs', prefs)
    
    # WebDriverManagerを使用してChromeDriverを自動管理
    service = Service(ChromeDriverManager().install())
    
    # WebDriverインスタンスを作成
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    # User-Agentを設定（Bot検知回避）
    # 環境変数 CHROME_USER_AGENT から読み込む（設定されていない場合はデフォルト値を使用）
    driver.execute_cdp_cmd('Network.setUserAgentOverride', {
        "userAgent": CHROME_USER_AGENT
    })
    
    return driver
