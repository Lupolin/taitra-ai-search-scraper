from automation.browser_edge import EdgeBrowserManager
from utils.logger import setup_logging
from automation.login import login
from automation.clear_folder import clear_downloads_folder
import time
import config

def launch_edge_with_profile():
    logger.info("啟動 Edge 瀏覽器（使用本地登入檔）...")
    browser = EdgeBrowserManager()
    driver = browser.setup_driver()
    browser.driver.maximize_window()
    driver.get(config.WEBSITE_URL)
    return browser

if __name__ == "__main__":
    logger = setup_logging("啟動瀏覽器")
    browser = launch_edge_with_profile()

    if browser:
        logger = setup_logging("登入網頁")
        # 在這裡進行您的操作
        login(browser.driver, config.WEBSITE_USERNAME, config.WEBSITE_PASSWORD, logger)
        # 例如：登入、下載、上傳等
        time.sleep(5)  # 等待 5 秒（您可以替換為實際操作）
        
        logger.info("操作完成，關閉瀏覽器...")
        # close_global_browser()
        logger.info("✓ 程式結束")
    else:
        logger = setup_logging("下載流程開始")
        logger.error("啟動失敗！") 
    
    clear_downloads_folder()
