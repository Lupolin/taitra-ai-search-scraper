from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import config
from automation.captcha_solver import solve_captcha_from_element
import logging

def login(driver, username, password, logger=None):
    """執行登入流程"""
    if logger is None:
        logger = logging.getLogger(__name__)
        
    try:
        wait = WebDriverWait(driver, 10)
        
        logger.info("開始登入流程...")
        
        # 等待並輸入帳號
        logger.info("等待帳號輸入欄位...")
        username_input = wait.until(EC.presence_of_element_located((By.NAME, "user_name")))
        logger.info("找到帳號輸入欄位，開始輸入帳號...")
        username_input.send_keys(username)
        logger.info(f"已輸入帳號: {username}")

        # 輸入密碼
        password_input = driver.find_element(By.NAME, "user_password")
        password_input.send_keys(password)
        
        # # 檢查是否有驗證碼圖片，如果有才進行辨識
        # logger.info("檢查是否有驗證碼圖片...")
        # captcha_elements = driver.find_elements(By.ID, "captcha_img")
        # if captcha_elements:
        #     captcha_element = captcha_elements[0]
        #     logger.info("發現驗證碼圖片，開始進行 OCR 辨識")
        #     try:
        #         captcha_text = solve_captcha_from_element(driver, captcha_element)
                
        #         if captcha_text:
        #             captcha_input = driver.find_element(By.NAME, "vercode") 
        #             captcha_input.send_keys(captcha_text)
        #             logger.info(f"已輸入驗證碼: {captcha_text}")
        #         else:
        #             logger.warning("驗證碼辨識失敗，跳過驗證碼輸入")
        #     except Exception as ocr_error:
        #         logger.warning(f"OCR 處理失敗: {ocr_error}，跳過驗證碼輸入")
        # else:
        #     logger.info("未發現驗證碼圖片，跳過驗證碼輸入")
        
        # 點擊登入按鈕（用 id 最穩）
        login_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "new_btn")))
        login_btn.click()
        return True
    except Exception as e:
        logger.error(f"登入流程發生錯誤: {e}")
        return False