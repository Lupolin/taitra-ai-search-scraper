"""
AI Tool Search Automation Package

這個套件包含所有自動化相關的功能模組：
- 瀏覽器管理
- 檔案下載
- SharePoint 上傳
- Teams 通知
- 工具函數
"""

from .browser_chrome import BrowserManager, set_global_browser, close_global_browser
from .login import login
from .download_excel import download_excel
from .upload_sharepoint import upload_temp_files_to_sharepoint
from .notify_teams_result import report_download_status
from .clear_folder import clear_temp_folder
from .rename_query_file import rename_query_file

__version__ = "1.0.0"
__author__ = "AI Tool Search Team"

# 主要功能函數
__all__ = [
    "BrowserManager",
    "set_global_browser", 
    "close_global_browser",
    "login",
    "download_excel",
    "upload_temp_files_to_sharepoint",
    "report_download_status",
    "clear_temp_folder",
    "rename_query_file"
] 