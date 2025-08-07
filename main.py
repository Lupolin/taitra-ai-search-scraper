from automation import (
    BrowserManager, set_global_browser, close_global_browser,
    login, download_excel, upload_temp_files_to_sharepoint,
    clear_temp_folder, report_download_status
)
from utils.logger import setup_logging
from utils.execution_logger import execution_logger
import time
import config

def chrome_simple(logger):
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
    # 啟動瀏覽器
    browser_logger = setup_logging("啟動瀏覽器")
    browser = chrome_simple(browser_logger)

    if browser:
        # 登入網頁
        login_logger = setup_logging("登入網頁")
        login(browser.driver, config.WEBSITE_USERNAME, config.WEBSITE_PASSWORD)

        # 下載流程
        download_logger = setup_logging("下載流程")
        
        # 設定時間區段：使用 config 中的時間設定
        for hour in range(config.START_HOUR, config.END_HOUR + 1):
            for minute in range(config.START_MINUTE, config.END_MINUTE, config.TIME_INTERVAL_MINUTES):
                success = download_excel(browser.driver, hour, minute, minute + 9, download_logger)
                if not success:
                    download_logger.warning(f"⚠️ 時段 {hour}:{minute:02} ~ {hour}:{minute+9:02} 處理失敗")
                    
                    # 檢查是否需要重新啟動瀏覽器
                    try:
                        browser.driver.current_url
                        download_logger.info("✅ 瀏覽器 session 仍然有效，繼續下一個時段")
                    except Exception as e:
                        download_logger.error(f"❌ 瀏覽器 session 已失效：{e}")
                        download_logger.info("🔄 嘗試重新啟動瀏覽器...")
                        
                        try:
                            # 關閉舊瀏覽器
                            close_global_browser()
                            
                            # 重新啟動瀏覽器
                            browser = chrome_simple(browser_logger)
                            if browser:
                                # 重新登入
                                login(browser.driver, config.WEBSITE_USERNAME, config.WEBSITE_PASSWORD)
                                download_logger.info("✅ 瀏覽器重新啟動成功")
                            else:
                                download_logger.error("❌ 無法重新啟動瀏覽器，跳過剩餘時段")
                                break
                        except Exception as restart_error:
                            download_logger.error(f"❌ 重新啟動瀏覽器失敗：{restart_error}")
                            break
                
                time.sleep(2)

        download_logger.info("✓ 所有時間區段處理完畢，關閉瀏覽器...")
        close_global_browser()
        
        # 上傳到 SharePoint
        upload_logger = setup_logging("上傳到 SharePoint")
        try:
            upload_result = upload_temp_files_to_sharepoint(upload_logger)
            upload_logger.info("✓ 上傳流程完成")
        except Exception as e:
            upload_logger.error(f"❌ 上傳流程失敗: {str(e)}")
        
        # 發送結果到 Teams
        teams_logger = setup_logging("發送結果到 Teams")
        try:
            report_download_status()
            teams_logger.info("✓ 發送結果到 Teams 完成")
        except Exception as e:
            teams_logger.error(f"❌ Teams 通知失敗: {str(e)}")
        
        # 輸出執行摘要
        summary = execution_logger.get_summary_text()
        stats = execution_logger.get_statistics()
        
        # 清除 temp 資料夾
        cleanup_logger = setup_logging("清除 temp 資料夾")
        try:
            clear_temp_folder()
            cleanup_logger.info("✓ 清除 temp 資料夾完成")
        except Exception as e:
            cleanup_logger.error(f"❌ 清除 temp 資料夾失敗: {str(e)}")
    else:
        browser_logger.error("❌ 無法啟動瀏覽器！")