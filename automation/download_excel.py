from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from automation.captcha_solver import solve_captcha_from_element
from datetime import datetime, timedelta
import config
import logging
import time

def download_excel(driver, logger=None):
    """執行登入後下載資料的流程"""
    if logger is None:
        logger = logging.getLogger(__name__)

    try:
        wait = WebDriverWait(driver, 10)

        logger.info("開始下載流程...")
        time.sleep(5)
        
        # 點擊 log 資料按鈕
        log_data_btn = wait.until(EC.element_to_be_clickable((By.ID, "m0")))
        log_data_btn.click()
        time.sleep(1)

        # 點擊篩選按鈕
        wait.until(EC.presence_of_element_located((By.ID, "query_btn")))
        filter_btn = wait.until(EC.element_to_be_clickable((By.ID, "query_btn")))
        filter_btn.click()

        # 前一天的日期
        yesterday = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
        logger.info(f"設定查詢日期為：{yesterday}")

        # 等待日期欄位出現
        wait.until(EC.presence_of_element_located((By.ID, "minDate")))
        wait.until(EC.presence_of_element_located((By.ID, "maxDate")))

        # 設定開始日期（minDate）
        driver.execute_script("""
            const minInput = document.getElementById('minDate');
            if (minInput) {
                minInput.value = arguments[0];
                minInput.dispatchEvent(new Event('change'));
            }
        """, yesterday)

        # 設定結束日期（maxDate）
        driver.execute_script("""
            const maxInput = document.getElementById('maxDate');
            if (maxInput) {
                maxInput.value = arguments[0];
                maxInput.dispatchEvent(new Event('change'));
            }
        """, yesterday)

        logger.info("日期設定完成。")

        return True

    except Exception as e:
        logger.error(f"下載流程發生錯誤: {e}")
        return False