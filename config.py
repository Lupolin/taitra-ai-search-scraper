import os
from datetime import datetime, timedelta
from dotenv import load_dotenv
import os

# 載入 .env 檔案
load_dotenv()

# 網站設定
WEBSITE_URL = os.getenv('WEBSITE_URL')
LOGIN_URL = f"{WEBSITE_URL}/login"
DOWNLOAD_URL = f"{WEBSITE_URL}/download"

# 網站登入憑證
WEBSITE_USERNAME = os.getenv('WEBSITE_USERNAME')
WEBSITE_PASSWORD = os.getenv('WEBSITE_PASSWORD')

# Graph API 設定
GRAPH_API_CLIENT_ID = os.getenv('GRAPH_API_CLIENT_ID')
GRAPH_API_CLIENT_SECRET = os.getenv('GRAPH_API_CLIENT_SECRET')
GRAPH_API_TENANT_ID = os.getenv('GRAPH_API_TENANT_ID')
GRAPH_API_DRIVE_ID = os.getenv('GRAPH_API_DRIVE_ID')
GRAPH_API_FOLDER_ID = os.getenv('GRAPH_API_FOLDER_ID')

# Azure OpenAI 設定
AZURE_OPENAI_ENDPOINT = os.getenv('AZURE_OPENAI_ENDPOINT')
AZURE_OPENAI_API_KEY = os.getenv('AZURE_OPENAI_API_KEY')
AZURE_OPENAI_DEPLOYMENT_NAME = os.getenv('AZURE_OPENAI_DEPLOYMENT_NAME')

# 資料夾設定
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_FOLDER = os.path.join(BASE_DIR, "temp")

# 時間設定
TIME_INTERVAL_MINUTES = 10  # 每10分鐘一個區段
START_TIME = "08:00"
END_TIME = "18:00"

# SharePoint 設定
SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL')

# 瀏覽器設定
BROWSER_HEADLESS = False  # 設為 True 可隱藏瀏覽器視窗
BROWSER_TIMEOUT = 30
DOWNLOAD_TIMEOUT = 300  # 5分鐘下載超時

# 重試設定
MAX_RETRIES = 3
RETRY_DELAY = 5  # 秒

# 日誌設定
LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_FILE = os.path.join(BASE_DIR, "rpa.log")

# 確保必要資料夾存在
def ensure_directories():
    """確保必要的資料夾存在"""
    directories = [DOWNLOAD_FOLDER]
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"建立資料夾: {directory}")

# 初始化時建立資料夾
ensure_directories() 