from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from .rename_query_file import rename_query_file
import logging
import time
import config

def download_excel(driver, hour, start_minute, end_minute, logger=None):
    if logger is None:
        logger = logging.getLogger(__name__)

    try:
        # 檢查瀏覽器 session 是否有效
        try:
            driver.current_url
            logger.info("✅ 瀏覽器 session 有效")
        except Exception as e:
            logger.error(f"❌ 瀏覽器 session 無效：{e}")
            return False
        
        wait = WebDriverWait(driver, 10)
        logger.info(f"⏱ 開始下載：{hour}:{start_minute:02} 到 {hour}:{end_minute:02}")

        time.sleep(5)  # 等待首頁加載完成

        # 點擊主選單的「Log Data」
        log_data_btn = wait.until(EC.element_to_be_clickable((By.ID, "m0")))
        log_data_btn.click()
        time.sleep(3)  # 等待頁面切換

        # 🔄 切換到 iframe（根據你的 HTML 應該是 iframe id 為 iframepagef2）
        driver.switch_to.frame("iframepagef2")

        # 點擊查詢按鈕，觸發表單區塊出現
        driver.execute_script("document.querySelector('#query_btn').click();")
        time.sleep(1)

        # 準備昨天的日期
        date = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
        from_time = f"{hour:02}:{start_minute:02}:00"
        to_time = f"{hour:02}:{end_minute:02}:59"

        # 設定日期欄位（minDate, maxDate）
        wait.until(EC.presence_of_element_located((By.ID, "minDate")))
        wait.until(EC.presence_of_element_located((By.ID, "maxDate")))
        driver.execute_script("""
            const minInput = document.getElementById('minDate');
            const maxInput = document.getElementById('maxDate');
            if (minInput && maxInput) {
                minInput.value = arguments[0];
                maxInput.value = arguments[0];
                minInput.dispatchEvent(new Event('change'));
                maxInput.dispatchEvent(new Event('change'));
            }
        """, date)

        # 設定時間欄位（minTime, maxTime）
        wait.until(EC.presence_of_element_located((By.ID, "minTime")))
        wait.until(EC.presence_of_element_located((By.ID, "maxTime")))
        driver.execute_script("""
            const minT = document.getElementById('minTime');
            const maxT = document.getElementById('maxTime');
            if (minT && maxT) {
                minT.value = arguments[0];
                maxT.value = arguments[1];
                minT.dispatchEvent(new Event('change'));
                maxT.dispatchEvent(new Event('change'));
            }
        """, from_time, to_time)

        logger.info(f"✓ 已設定日期 {date}，時間區間 {from_time} ~ {to_time}")

        # ✅ 執行查詢（再次點擊查詢按鈕）
        driver.execute_script("document.querySelector('#query_btn').click();")
        time.sleep(2)

        # ✅ 點擊「匯出」按鈕
        driver.execute_script("document.querySelector('#export_btn').click();")
        time.sleep(2)

        driver.execute_script("""
            [...document.querySelectorAll('a.click-btn')]
                .find(el => el.textContent.trim() === 'Export to EXCEL')
                ?.click();
        """)
        time.sleep(2)


        # 1️⃣ 等待 JavaScript alert 彈出（例如：大量資料請稍候）
        WebDriverWait(driver, 10).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        logger.info(f"📢 彈跳視窗內容：{alert.text}")
        alert.accept()  # 點擊「確定」
        logger.info("✓ 成功點擊彈跳視窗的 OK")

        # 假設目前只有一個 Tab，點擊後會開新 Tab，這裡先記住原本的 Tab
        original_window = driver.current_window_handle
        logger.info("📌 記錄原本的瀏覽器 Tab")

        # 🔁 等待新 Tab 開啟（總共應變成 2 個）
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)

        # 切換到新開的 Tab（找到不是原本的那個）
        for handle in driver.window_handles:
            if handle != original_window:
                driver.switch_to.window(handle)
                logger.info("🆕 已切換到新開的下載頁面 Tab")
                break

        # ✅ 等待「Download File」的連結出現並點擊
        download_link = WebDriverWait(driver, config.DOWNLOAD_TIMEOUT).until(
            EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Download File"))
        )
        download_link.click()
        logger.info("📥 成功點擊下載連結")

        # 🔙 可選：下載完畢後關閉新 Tab，並切回原本頁面
        driver.close()
        driver.switch_to.window(original_window)
        logger.info("🔄 關閉下載 Tab 並切回原本頁面")
        
        # 等待檔案下載完成
        time.sleep(5)  # 給檔案下載 5 秒時間
        
        date = (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")
        rename_query_file(driver, date, hour, start_minute, end_minute, logger)
        return True

    except Exception as e:
        logger.error(f"❌ 下載流程錯誤: {str(e)}")
        try:
            driver.close()
            driver.switch_to.window(original_window)
            logger.info("🔄 關閉下載 Tab 並切回原本頁面")
        except:
            pass
        return False
