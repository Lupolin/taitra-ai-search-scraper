from automation.browser_chrome import BrowserManager, set_global_browser, close_global_browser
from utils.logger import setup_logging
import time
import config
from automation.login import login

def chrome_simple():
    logger.info("開啟 Chrome 瀏覽器...")
    
    try:
        # 建立瀏覽器並設定為全域實例
        browser = BrowserManager(headless=False)
        browser.setup_driver()
        set_global_browser(browser)  # 設定為全域瀏覽器
        
        logger.info("✓ Chrome 瀏覽器成功啟動")
        
        # 開啟網頁
        browser.driver.maximize_window()
        browser.driver.get(config.WEBSITE_URL)
        logger.info("✓ 成功開啟網頁")
        
        return browser
        
    except Exception as e:
        logger.error(f"❌ 測試失敗: {e}")
        return None

if __name__ == "__main__":
    logger = setup_logging("啟動瀏覽器")
    browser = chrome_simple()
    
    if browser:
        logger = setup_logging("登入網頁")
        # 在這裡進行您的操作
        login(browser.driver, config.WEBSITE_USERNAME, config.WEBSITE_PASSWORD)
        # 例如：登入、下載、上傳等
        time.sleep(5)  # 等待 5 秒（您可以替換為實際操作）
        
        logger.info("操作完成，關閉瀏覽器...")
        # close_global_browser()
        logger.info("✓ 程式結束")
    else:
        logger = setup_logging("下載流程開始")
        logger.error("啟動失敗！") 