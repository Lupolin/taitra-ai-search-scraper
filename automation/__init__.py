"""
AI Tool Search Automation Package

這個套件包含所有自動化相關的功能模組：
- browser: 瀏覽器管理
- download: 檔案下載
- upload: SharePoint 上傳
- notification: Teams 通知
- utils: 工具函數
"""

# Browser 模組
from .browser.browser_chrome import BrowserManager, set_global_browser, close_global_browser
from .browser.login import login

# Download 模組
from .download.download_excel import download_excel
from .download.rename_query_file import rename_query_file

# Upload 模組
from .upload.upload_sharepoint import upload_temp_files_to_sharepoint

# Notification 模組
from .notification.notify_teams_result import report_download_status

# Utils 模組
from .utils.clear_folder import clear_temp_folder
from .utils.captcha_solver import solve_captcha

__version__ = "1.0.0"
__author__ = "AI Tool Search Team"

# 主要功能函數
__all__ = [
    # Browser
    "BrowserManager",
    "set_global_browser", 
    "close_global_browser",
    "login",
    
    # Download
    "download_excel",
    "rename_query_file",
    
    # Upload
    "upload_temp_files_to_sharepoint",
    
    # Notification
    "report_download_status",
    
    # Utils
    "clear_temp_folder",
    "solve_captcha"
] 