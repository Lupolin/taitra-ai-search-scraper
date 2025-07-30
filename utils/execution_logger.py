from datetime import datetime, timedelta
from typing import Dict, List, Optional
import logging
import os
import config

class ExecutionLogger:
    """執行結果記錄器，整合現有的檔案檢查機制"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def generate_expected_filenames(self) -> list[str]:
        """產生預期的檔案名稱列表"""
        filenames = []
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
        for hour in range(8, 18):
            for i in range(6):
                start_min = i * 10
                end_min = start_min + 9
                start_str = f"{hour:02d}{start_min:02d}"
                end_str = f"{hour:02d}{end_min:02d}"
                interval = f"{start_str}-{end_str}"
                filename = f"logger_url_{yesterday}_{interval}.xls"
                filenames.append(filename)
        return filenames
    
    def check_download_files(self) -> Dict:
        """檢查下載檔案並回傳詳細結果"""
        download_folder = config.DOWNLOAD_FOLDER
        expected_files = self.generate_expected_filenames()
        
        results = {
            "success_count": 0,
            "fail_count": 0,
            "total_count": len(expected_files),
            "success_files": [],
            "failed_files": [],
            "details": []
        }
        
        for filename in expected_files:
            full_path = os.path.join(download_folder, filename)
            interval = filename.replace(".xls", "")
            exists = os.path.exists(full_path)
            
            file_result = {
                "filename": filename,
                "interval": interval,
                "exists": exists,
                "path": full_path
            }
            
            results["details"].append(file_result)
            
            if exists:
                results["success_count"] += 1
                results["success_files"].append(filename)
            else:
                results["fail_count"] += 1
                results["failed_files"].append(filename)
        
        return results
    
    def get_summary_text(self) -> str:
        """取得格式化的摘要文字"""
        results = self.check_download_files()
        log_date = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        
        lines = [f"{log_date} AI上網統計紀錄：\n"]
        
        for detail in results["details"]:
            status = "✅" if detail["exists"] else "❌"
            lines.append(f"{detail['interval']}-{status}")
        
        lines.append("")  # 空行
        lines.append(f"✅ 成功：{results['success_count']}，❌ 失敗：{results['fail_count']}")
        
        return "\n".join(lines)
    
    def get_statistics(self) -> Dict:
        """取得統計資訊"""
        results = self.check_download_files()
        total = results["total_count"]
        success = results["success_count"]
        failed = results["fail_count"]
        
        return {
            "total": total,
            "success": success,
            "failed": failed,
            "success_rate": (success / total * 100) if total > 0 else 0,
            "failed_files": results["failed_files"]
        }
    
    def get_failed_segments(self) -> List[str]:
        """取得失敗的區段列表"""
        results = self.check_download_files()
        return results["failed_files"]

# 全域執行記錄器實例
execution_logger = ExecutionLogger() 