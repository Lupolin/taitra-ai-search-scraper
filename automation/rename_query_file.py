import os
import time
import glob
import config
from datetime import date, datetime

def rename_query_file(driver, date, hour, start_minute, end_minute, logger=None):
    download_dir = config.DOWNLOAD_FOLDER
    print(f"🔍 下載資料夾路徑: {download_dir}")

    matched_file = None
    wait_time = 0
    while wait_time < 60:
        candidates = glob.glob(os.path.join(download_dir, "query*.xls"))
        valid_files = [f for f in candidates if not f.endswith(".crdownload")]
        print(f"⏱️ 嘗試第 {wait_time+1} 秒, 找到檔案：{valid_files}")
        if valid_files:
            matched_file = max(valid_files, key=os.path.getctime)
            break
        time.sleep(1)
        wait_time += 1

    if not matched_file:
        error_msg = "❌ 找不到有效的 query.xls 檔案或下載超時"
        if logger:
            logger.error(error_msg)
        raise FileNotFoundError(error_msg)

    start_str = f"{hour:02}{start_minute:02}"
    end_str = f"{hour:02}{end_minute:02}"
    new_filename = f"logger_url_{date}_{start_str}-{end_str}.xls"
    new_path = os.path.join(download_dir, new_filename)

    try:
        os.rename(matched_file, new_path)
    except Exception as e:
        error_msg = f"❌ 檔案改名失敗: {e}"
        if logger:
            logger.error(error_msg)
        raise

    success_msg = f"✅ 檔案重新命名為：{new_filename}"
    if logger:
        logger.info(success_msg)
    else:
        print(success_msg)

    return new_path
