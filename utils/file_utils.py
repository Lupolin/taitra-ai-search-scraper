"""
檔案工具模組
負責檢查下載檔案、命名、確認成功等檔案操作
"""

import os
import time
import shutil
from datetime import datetime
import config

def wait_for_download(existing_files, timeout=300):
    """
    等待檔案下載完成
    
    Args:
        existing_files (set): 下載前的檔案列表
        timeout (int): 等待超時時間（秒）
        
    Returns:
        str: 下載的檔案路徑，超時則返回 None
    """
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        try:
            # 檢查下載資料夾中的新檔案
            if os.path.exists(config.DOWNLOAD_FOLDER):
                current_files = set(os.listdir(config.DOWNLOAD_FOLDER))
                new_files = current_files - existing_files
                
                # 檢查是否有新檔案且不是臨時檔案
                for filename in new_files:
                    file_path = os.path.join(config.DOWNLOAD_FOLDER, filename)
                    
                    # 跳過臨時檔案
                    if filename.endswith('.tmp') or filename.endswith('.crdownload'):
                        continue
                    
                    # 檢查檔案是否完整（不是正在下載中）
                    if is_file_complete(file_path):
                        return file_path
            
            time.sleep(1)  # 等待1秒後再檢查
            
        except Exception as e:
            print(f"檢查下載檔案時發生錯誤: {e}")
            time.sleep(1)
    
    print(f"下載超時 ({timeout} 秒)")
    return None

def is_file_complete(file_path, check_interval=1):
    """
    檢查檔案是否下載完成
    
    Args:
        file_path (str): 檔案路徑
        check_interval (int): 檢查間隔（秒）
        
    Returns:
        bool: 檔案是否完整
    """
    try:
        if not os.path.exists(file_path):
            return False
        
        # 檢查檔案大小是否穩定
        initial_size = os.path.getsize(file_path)
        time.sleep(check_interval)
        final_size = os.path.getsize(file_path)
        
        # 如果檔案大小沒有變化，表示下載完成
        return initial_size == final_size and initial_size > 0
        
    except Exception as e:
        print(f"檢查檔案完整性時發生錯誤: {e}")
        return False

def verify_download_file(file_path):
    """
    驗證下載的檔案
    
    Args:
        file_path (str): 檔案路徑
        
    Returns:
        bool: 檔案是否有效
    """
    try:
        if not os.path.exists(file_path):
            print(f"檔案不存在: {file_path}")
            return False
        
        # 檢查檔案大小
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            print(f"檔案大小為0: {file_path}")
            return False
        
        # 檢查檔案副檔名
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in ['.xlsx', '.xls', '.csv']:
            print(f"不支援的檔案格式: {file_ext}")
            return False
        
        # 嘗試開啟檔案（基本驗證）
        try:
            with open(file_path, 'rb') as f:
                # 讀取檔案開頭來驗證檔案格式
                header = f.read(8)
                
                # Excel 檔案驗證
                if file_ext in ['.xlsx', '.xls']:
                    if not header.startswith(b'\xd0\xcf\x11\xe0') and not header.startswith(b'PK\x03\x04'):
                        print(f"檔案格式可能損壞: {file_path}")
                        return False
                
                # CSV 檔案驗證
                elif file_ext == '.csv':
                    if len(header) == 0:
                        print(f"CSV 檔案為空: {file_path}")
                        return False
            
            print(f"檔案驗證成功: {file_path} ({file_size} bytes)")
            return True
            
        except Exception as e:
            print(f"檔案開啟失敗: {file_path} - {e}")
            return False
            
    except Exception as e:
        print(f"驗證檔案時發生錯誤: {e}")
        return False

def generate_filename(start_time, end_time, date=None, prefix=None, extension=None):
    """
    生成檔案名稱
    
    Args:
        start_time (str): 開始時間 (HH:MM 格式)
        end_time (str): 結束時間 (HH:MM 格式)
        date (str): 日期 (YYYY-MM-DD 格式)，None 則使用今天
        prefix (str): 檔案前綴，None 則使用 config 設定
        extension (str): 檔案副檔名，None 則使用 config 設定
        
    Returns:
        str: 生成的檔案名稱
    """
    # 使用預設值
    prefix = prefix or config.FILE_PREFIX
    extension = extension or config.FILE_EXTENSION
    
    # 格式化時間
    start_formatted = start_time.replace(':', '')
    end_formatted = end_time.replace(':', '')
    
    # 日期
    if date is None:
        date = datetime.now().strftime("%Y-%m-%d")
    date_formatted = date.replace('-', '')
    
    # 生成檔案名稱
    filename = f"{prefix}_{date_formatted}_{start_formatted}-{end_formatted}{extension}"
    
    return filename

def organize_downloads_by_date(date=None):
    """
    按日期整理下載的檔案
    
    Args:
        date (str): 日期 (YYYY-MM-DD 格式)，None 則使用今天
        
    Returns:
        list: 整理後的檔案路徑列表
    """
    if date is None:
        date = datetime.now().strftime("%Y-%m-%d")
    
    # 建立日期資料夾
    date_folder = os.path.join(config.DOWNLOAD_FOLDER, date)
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"建立日期資料夾: {date_folder}")
    
    # 移動檔案到日期資料夾
    moved_files = []
    
    if os.path.exists(config.DOWNLOAD_FOLDER):
        for filename in os.listdir(config.DOWNLOAD_FOLDER):
            file_path = os.path.join(config.DOWNLOAD_FOLDER, filename)
            
            # 跳過資料夾
            if os.path.isdir(file_path):
                continue
            
            # 檢查是否為目標檔案格式
            if filename.endswith(config.FILE_EXTENSION):
                # 移動到日期資料夾
                new_path = os.path.join(date_folder, filename)
                try:
                    shutil.move(file_path, new_path)
                    moved_files.append(new_path)
                    print(f"移動檔案: {filename} -> {date_folder}")
                except Exception as e:
                    print(f"移動檔案失敗: {filename} - {e}")
    
    return moved_files

def cleanup_temp_files():
    """
    清理臨時檔案
    
    Returns:
        int: 清理的檔案數量
    """
    cleaned_count = 0
    
    if os.path.exists(config.DOWNLOAD_FOLDER):
        for filename in os.listdir(config.DOWNLOAD_FOLDER):
            file_path = os.path.join(config.DOWNLOAD_FOLDER, filename)
            
            # 檢查是否為臨時檔案
            if filename.endswith(('.tmp', '.crdownload', '.part')):
                try:
                    os.remove(file_path)
                    cleaned_count += 1
                    print(f"清理臨時檔案: {filename}")
                except Exception as e:
                    print(f"清理臨時檔案失敗: {filename} - {e}")
    
    return cleaned_count

def get_file_info(file_path):
    """
    取得檔案資訊
    
    Args:
        file_path (str): 檔案路徑
        
    Returns:
        dict: 檔案資訊字典
    """
    try:
        if not os.path.exists(file_path):
            return None
        
        stat = os.stat(file_path)
        
        return {
            'name': os.path.basename(file_path),
            'path': file_path,
            'size': stat.st_size,
            'created': datetime.fromtimestamp(stat.st_ctime),
            'modified': datetime.fromtimestamp(stat.st_mtime),
            'extension': os.path.splitext(file_path)[1]
        }
        
    except Exception as e:
        print(f"取得檔案資訊時發生錯誤: {e}")
        return None

def rename_file_with_timestamp(file_path, timestamp=None):
    """
    使用時間戳重新命名檔案
    
    Args:
        file_path (str): 原始檔案路徑
        timestamp (str): 時間戳，None 則使用當前時間
        
    Returns:
        str: 新的檔案路徑，失敗則返回 None
    """
    try:
        if not os.path.exists(file_path):
            return None
        
        # 生成時間戳
        if timestamp is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 取得檔案資訊
        dir_path = os.path.dirname(file_path)
        filename = os.path.basename(file_path)
        name, ext = os.path.splitext(filename)
        
        # 生成新檔名
        new_filename = f"{name}_{timestamp}{ext}"
        new_path = os.path.join(dir_path, new_filename)
        
        # 重新命名
        os.rename(file_path, new_path)
        print(f"檔案重新命名: {filename} -> {new_filename}")
        
        return new_path
        
    except Exception as e:
        print(f"重新命名檔案時發生錯誤: {e}")
        return None 