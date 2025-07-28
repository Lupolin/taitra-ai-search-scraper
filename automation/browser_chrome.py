from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import config
import os

# 全域瀏覽器實例
_global_browser = None

class BrowserManager:
    def __init__(self, headless=None):
        self.driver = None
        self.headless = headless if headless is not None else config.BROWSER_HEADLESS
        
    def setup_driver(self):
        try:
            # Chrome 選項設定
            chrome_options = Options()
            
            if self.headless:
                chrome_options.add_argument("--headless")
            
            # 基本設定
            chrome_options.add_argument('--ignore-certificate-errors')  
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # # 使用現有的 Chrome 使用者資料夾（保留登入狀態）
            # chrome_options.add_argument("--user-data-dir=C:\\Users\\E100\\AppData\\Local\\Google\\Chrome\\User Data")
            # chrome_options.add_argument("--profile-directory=Default")
            # 使用臨時用戶資料目錄避免衝突
            import tempfile
            temp_dir = tempfile.mkdtemp()
            chrome_options.add_argument(f"--user-data-dir={temp_dir}")
            
            # 下載設定
            prefs = {
                "download.default_directory": config.DOWNLOAD_FOLDER,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            # 設定 WebDriver
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # 設定等待時間
            self.driver.implicitly_wait(config.BROWSER_TIMEOUT)
            
            # 執行腳本來隱藏 webdriver 屬性
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            print("Chrome WebDriver 設定完成")
            return self.driver
            
        except Exception as e:
            print(f"設定 Chrome WebDriver 時發生錯誤: {e}")
            raise
    
    def close_driver(self):
        """關閉瀏覽器"""
        if self.driver:
            self.driver.quit()
            print("瀏覽器已關閉")
    
    def __enter__(self):
        """Context manager 進入"""
        self.setup_driver()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager 退出"""
        self.close_driver()

# 全域瀏覽器控制函數
def set_global_browser(browser):
    """設定全域瀏覽器實例"""
    global _global_browser
    _global_browser = browser

def close_global_browser():
    """關閉全域瀏覽器"""
    global _global_browser
    if _global_browser and _global_browser.driver:
        _global_browser.close_driver()
        _global_browser = None
        print("全域瀏覽器已關閉")
    else:
        print("沒有開啟的瀏覽器") 