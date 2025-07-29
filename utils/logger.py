"""
日誌模組
提供可選的日誌功能
"""

import logging
import os
from datetime import datetime
import config

class RPALogger:
    """RPA 日誌管理器"""
    
    def __init__(self, name="RPA", log_file=None, log_level=None):
        """
        初始化日誌管理器
        
        Args:
            name (str): 日誌名稱
            log_file (str): 日誌檔案路徑，None 則使用 config 設定
            log_level (str): 日誌等級，None 則使用 config 設定
        """
        self.name = name
        self.log_file = log_file or config.LOG_FILE
        self.log_level = log_level or config.LOG_LEVEL
        
        # 建立 logger
        self.logger = logging.getLogger(name)
        self.logger.setLevel(getattr(logging, self.log_level.upper()))
        
        # 清除現有的 handlers
        self.logger.handlers.clear()
        
        # 設定格式
        formatter = logging.Formatter(config.LOG_FORMAT)
        
        # 檔案 handler
        if self.log_file:
            # 確保日誌目錄存在
            log_dir = os.path.dirname(self.log_file)
            if log_dir and not os.path.exists(log_dir):
                os.makedirs(log_dir)
            
            file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
            file_handler.setLevel(getattr(logging, self.log_level.upper()))
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)
        
        # 控制台 handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(getattr(logging, self.log_level.upper()))
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)
    
    def __str__(self):
        """字串表示"""
        return f"RPALogger({self.name})"
    
    def __format__(self, format_spec):
        """格式化方法"""
        return self.__str__()
    
    def info(self, message):
        """記錄資訊訊息"""
        self.logger.info(message)
    
    def warning(self, message):
        """記錄警告訊息"""
        self.logger.warning(message)
    
    def error(self, message):
        """記錄錯誤訊息"""
        self.logger.error(message)
    
    def debug(self, message):
        """記錄除錯訊息"""
        self.logger.debug(message)
    
    def critical(self, message):
        """記錄嚴重錯誤訊息"""
        self.logger.critical(message)
    
    def log_download_start(self, start_time, end_time, date=None):
        """記錄下載開始"""
        date_str = date or datetime.now().strftime("%Y-%m-%d")
        message = f"開始下載時間區段: {date_str} {start_time}-{end_time}"
        self.info(message)
    
    def log_download_success(self, file_path, start_time, end_time):
        """記錄下載成功"""
        message = f"下載成功: {file_path} ({start_time}-{end_time})"
        self.info(message)
    
    def log_download_failure(self, start_time, end_time, error):
        """記錄下載失敗"""
        message = f"下載失敗: {start_time}-{end_time} - {error}"
        self.error(message)
    
    def log_upload_start(self, file_path):
        """記錄上傳開始"""
        message = f"開始上傳檔案: {file_path}"
        self.info(message)
    
    def log_upload_success(self, file_path, sharepoint_url=None):
        """記錄上傳成功"""
        message = f"上傳成功: {file_path}"
        if sharepoint_url:
            message += f" -> {sharepoint_url}"
        self.info(message)
    
    def log_upload_failure(self, file_path, error):
        """記錄上傳失敗"""
        message = f"上傳失敗: {file_path} - {error}"
        self.error(message)
    
    def log_session_start(self, date=None):
        """記錄會話開始"""
        date_str = date or datetime.now().strftime("%Y-%m-%d")
        message = f"=== RPA 會話開始: {date_str} ==="
        self.info(message)
    
    def log_session_end(self, total_downloads=0, total_uploads=0, errors=0):
        """記錄會話結束"""
        message = f"=== RPA 會話結束 ==="
        message += f" 下載: {total_downloads}, 上傳: {total_uploads}, 錯誤: {errors}"
        self.info(message)
    
    def log_config_info(self):
        """記錄配置資訊"""
        self.info("=== 配置資訊 ===")
        self.info(f"網站 URL: {config.WEBSITE_URL}")
        self.info(f"下載資料夾: {config.DOWNLOAD_FOLDER}")
        self.info(f"時間間隔: {config.TIME_INTERVAL_MINUTES} 分鐘")
        self.info(f"檔案格式: {config.FILE_EXTENSION}")
        self.info(f"SharePoint 資料夾: {config.SHAREPOINT_FOLDER}")

# 全域日誌實例
_logger = None

def get_logger(name="RPA", log_file=None, log_level=None):
    """
    取得日誌實例
    
    Args:
        name (str): 日誌名稱
        log_file (str): 日誌檔案路徑
        log_level (str): 日誌等級
        
    Returns:
        RPALogger: 日誌實例
    """
    global _logger
    if _logger is None:
        _logger = RPALogger(name, log_file, log_level)
    return _logger

def setup_logging(name="RPA", log_file=None, log_level=None):
    """
    設定日誌
    
    Args:
        name (str): 日誌名稱
        log_file (str): 日誌檔案路徑
        log_level (str): 日誌等級
        
    Returns:
        RPALogger: 日誌實例
    """
    return get_logger(name, log_file, log_level)

def log_function_call(func_name, *args, **kwargs):
    """
    記錄函數呼叫的裝飾器
    
    Args:
        func_name (str): 函數名稱
        *args: 位置參數
        **kwargs: 關鍵字參數
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            logger = get_logger()
            logger.debug(f"呼叫函數: {func_name}")
            try:
                result = func(*args, **kwargs)
                logger.debug(f"函數 {func_name} 執行成功")
                return result
            except Exception as e:
                logger.error(f"函數 {func_name} 執行失敗: {str(e)}")
                raise
        return wrapper
    return decorator

def log_execution_time(func_name):
    """
    記錄執行時間的裝飾器
    
    Args:
        func_name (str): 函數名稱
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            logger = get_logger()
            start_time = datetime.now()
            logger.debug(f"開始執行: {func_name}")
            
            try:
                result = func(*args, **kwargs)
                end_time = datetime.now()
                execution_time = (end_time - start_time).total_seconds()
                logger.info(f"執行完成: {func_name} (耗時: {execution_time:.2f} 秒)")
                return result
            except Exception as e:
                end_time = datetime.now()
                execution_time = (end_time - start_time).total_seconds()
                logger.error(f"執行失敗: {func_name} (耗時: {execution_time:.2f} 秒) - {str(e)}")
                raise
        return wrapper
    return decorator 