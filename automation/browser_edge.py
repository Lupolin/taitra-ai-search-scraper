from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import config
import tempfile
import os

# ✅ 定義 BASE_DIR（專案根目錄）
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

class EdgeBrowserManager:
    def __init__(self, headless=None):
        self.driver = None
        self.headless = headless if headless is not None else config.BROWSER_HEADLESS
        self.download_folder =  os.path.join(BASE_DIR, "downloads")

    def setup_driver(self):
        try:
            options = Options()

            # ─── Headless 模式 ───────────────────────────────
            if self.headless:
                options.add_argument("--headless")

            # ─── 通用瀏覽器優化 ─────────────────────────────
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--ignore-certificate-errors")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")

            # ─── WebDriver 偽裝 ──────────────────────────────
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)

            # ✅ 加入下載路徑設定
            prefs = {
                "download.default_directory": self.download_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            }
            options.add_experimental_option("prefs", prefs)

            # ✅ 可選：User-Agent 偽裝（可視需求啟用）
            # options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.67")

            # ─── 使用臨時使用者資料目錄 ──────────────────────
            temp_dir = tempfile.mkdtemp()
            options.add_argument(f"--user-data-dir={temp_dir}")

            # ─── 使用專案內的 EdgeDriver ──────────────────────
            base_dir = os.path.dirname(os.path.abspath(__file__))  # automation/
            project_root = os.path.abspath(os.path.join(base_dir, ".."))
            driver_path = os.path.join(project_root, "drivers", "msedgedriver.exe")
            service = Service(executable_path=driver_path)

            # ─── 啟動 WebDriver ─────────────────────────────
            self.driver = webdriver.Edge(service=service, options=options)
            self.driver.implicitly_wait(config.BROWSER_TIMEOUT)

            print("Edge WebDriver 設定完成")
            return self.driver

        except Exception as e:
            print(f"設定 Edge WebDriver 時發生錯誤: {e}")
            raise

    def close_driver(self):
        if self.driver:
            self.driver.quit()
            print("Edge 已關閉")
