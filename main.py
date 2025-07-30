from automation import (
    BrowserManager, set_global_browser, close_global_browser,
    login, download_excel, upload_temp_files_to_sharepoint,
    clear_temp_folder, report_download_status
)
from utils.logger import setup_logging
from utils.execution_logger import execution_logger
import time
import config

def chrome_simple():
    logger.info("開啟 Chrome 瀏覽器...")
    
    try:
        # 建立瀏覽器並設定為全域實例（RPA 模式）
        browser = BrowserManager(headless=False, rpa_mode=True)
        browser.setup_driver()
        set_global_browser(browser)  # 設定為全域瀏覽器
        
        logger.info("✓ Chrome 瀏覽器成功啟動")
        
        # 開啟網頁
        browser.driver.maximize_window()
        browser.driver.get(config.WEBSITE_URL)
        logger.info("✓ 成功開啟網頁")
        
        return browser
        
    except Exception as e:
        logger.error(f"❌ 測試失敗: {str(e)}")
        return None

if __name__ == "__main__":
    logger = setup_logging("啟動瀏覽器")
    browser = chrome_simple()

    if browser:
        logger = setup_logging("登入網頁")
        login(browser.driver, config.WEBSITE_USERNAME, config.WEBSITE_PASSWORD)

        logger = setup_logging("下載流程開始")

        # 設定時間區段：08:00 到 18:59，每 10 分鐘一組
        for hour in range(8, 9):  # 8 to 18
            for minute in range(0, 60, 10):  # 0, 10, 20, ..., 50
                success = download_excel(browser.driver, hour, minute, minute + 9, logger)
                if not success:
                    logger.warning(f"⚠️ 時段 {hour}:{minute:02} ~ {hour}:{minute+9:02} 處理失敗")
                time.sleep(2)

        logger.info("✓ 所有時間區段處理完畢，關閉瀏覽器...")
        close_global_browser()
        
        # 上傳到 SharePoint
        try:
            upload_result = upload_temp_files_to_sharepoint(logger)
            logger.info("✓ 上傳流程完成")
        except Exception as e:
            logger.error(f"❌ 上傳流程失敗: {str(e)}")
        
        # 發送結果到 Teams
        try:
            report_download_status()
            logger.info("✓ 發送結果到 Teams 完成")
        except Exception as e:
            logger.error(f"❌ Teams 通知失敗: {str(e)}")
        
        # 清除 temp 資料夾
        try:
            clear_temp_folder()
            logger.info("✓ 清除 temp 資料夾完成")
        except Exception as e:
            logger.error(f"❌ 清除 temp 資料夾失敗: {str(e)}")
        
        # 輸出執行摘要
        summary = execution_logger.get_summary_text()
        stats = execution_logger.get_statistics()
        
        logger.info("=" * 50)
        logger.info("📊 執行摘要:")
        logger.info(summary)
        logger.info(f"📈 統計: 總計 {stats['total']} 個區段，成功 {stats['success']} 個，失敗 {stats['failed']} 個")
        logger.info("=" * 50)
    else:
        logger.error("❌ 無法啟動瀏覽器！")