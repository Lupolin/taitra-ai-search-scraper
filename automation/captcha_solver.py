from PIL import Image
import cv2
import logging
import base64
import requests
import json
import os
import config

def solve_captcha_from_element(driver, element, save_path="captcha.png"):
    """從 WebElement 擷取圖片並進行 OCR"""
    logging.info(f"擷取驗證碼圖片到 {save_path}")
    element.screenshot(save_path)
    return solve_captcha_with_azure_openai(save_path)

def solve_captcha_with_azure_openai(image_path):
    """使用 Azure OpenAI Vision API 辨識驗證碼"""
    try:
        # Azure OpenAI 設定
        endpoint = config.AZURE_OPENAI_ENDPOINT
        api_key = config.AZURE_OPENAI_API_KEY
        deployment_name = config.AZURE_OPENAI_DEPLOYMENT_NAME
        
        if not all([endpoint, api_key, deployment_name]):
            logging.error("缺少 Azure OpenAI 設定，請檢查環境變數")
            return None
        
        # 讀取並編碼圖片
        with open(image_path, "rb") as image_file:
            image_data = base64.b64encode(image_file.read()).decode('utf-8')
        
        # 準備 API 請求
        url = f"{endpoint}/openai/deployments/{deployment_name}/chat/completions?api-version=2024-02-15-preview"
        
        headers = {
            "Content-Type": "application/json",
            "api-key": api_key
        }
        
        payload = {
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "請辨識這個驗證碼圖片中的文字。這是一個簡單的驗證碼，請只回傳辨識出的文字，不要包含任何其他說明或標點符號。"
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{image_data}"
                            }
                        }
                    ]
                }
            ],
            "max_tokens": 10,
            "temperature": 0.1
        }
        
        logging.info("發送請求到 Azure OpenAI Vision API...")
        response = requests.post(url, headers=headers, json=payload, timeout=30)
        
        if response.status_code == 200:
            result = response.json()
            captcha_text = result['choices'][0]['message']['content'].strip()
            logging.info(f"Azure OpenAI 辨識結果: {captcha_text}")
            return captcha_text
        else:
            logging.error(f"Azure OpenAI API 請求失敗: {response.status_code} - {response.text}")
            return None
            
    except Exception as e:
        logging.error(f"使用 Azure OpenAI 辨識驗證碼時發生錯誤: {e}")
        return None

def solve_captcha(image_path):
    """保留原有的 Tesseract 方法作為備用"""
    try:
        import pytesseract
        logging.info(f"使用 Tesseract 讀取驗證碼圖片: {image_path}")
        image = cv2.imread(image_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        logging.info("轉為灰階")
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
        logging.info("二值化處理")
        text = pytesseract.image_to_string(thresh, config='--psm 8 -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyz')
        logging.info(f"Tesseract OCR 結果: {text.strip()}")
        return text.strip()
    except ImportError:
        logging.error("Tesseract 未安裝，無法使用備用 OCR 方法")
        return None
    except Exception as e:
        logging.error(f"Tesseract OCR 處理失敗: {e}")
        return None
