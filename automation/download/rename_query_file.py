import os
import time
import glob
import config
import pandas as pd
import win32com.client
from datetime import datetime

def convert_xls_to_csv_trimmed(xls_file, output_csv_file, logger=None):
    """
    將 .xls 檔案轉換為指定格式的 .csv 並刪除前 6 列，調整欄位格式
    """
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        if logger:
            logger.info("✅ Excel COM 物件建立成功")
        else:
            print("✅ Excel COM 物件建立成功")

        wb = excel.Workbooks.Open(xls_file)
        temp_csv = xls_file.replace('.xls', '_temp.csv')
        wb.SaveAs(temp_csv, FileFormat=6, Local=True)
        wb.Close(False)

        df = None
        encodings = ['utf-8', 'big5', 'gbk', 'cp950', 'latin1']
        for encoding in encodings:
            try:
                df = pd.read_csv(temp_csv, header=None, encoding=encoding)
                if logger:
                    logger.info(f"使用編碼：{encoding}")
                else:
                    print(f"使用編碼：{encoding}")
                break
            except UnicodeDecodeError:
                continue

        if df is None:
            raise Exception("❌ 無法使用任何編碼讀取 CSV 檔案")

        # 移除前 6 行
        df_trimmed = df.iloc[6:].reset_index(drop=True)

        # 指定原始欄位名稱
        df_trimmed.columns = [
            "Index", "User", "Group Name", "Host IP", "Target IP",
            "App Type", "Specific applications", "Time", "Action/Result"
        ]

        # 轉換為指定格式的新表格
        df_final = pd.DataFrame({
            "Time": df_trimmed["Time"],
            "Src. IP": df_trimmed["User"],
            "Src. Port": "",
            "Dst. IP": df_trimmed["Target IP"],
            "Dst. Port": "",
            "User": df_trimmed["User"],
            "Show Name": df_trimmed["Host IP"],
            "User Group": df_trimmed["Group Name"],
            "URL": "",
            "Title": "",
            "Domain Cate.": df_trimmed["Specific applications"],
            "Action": df_trimmed["Action/Result"],
            "Src. Location": ""
        })

        df_final.to_csv(output_csv_file, index=False, encoding='utf-8-sig')

        if os.path.exists(temp_csv):
            os.remove(temp_csv)
        try:
            excel.Quit()
        except:
            pass

        if logger:
            logger.info(f"✅ 轉換完成：{os.path.basename(output_csv_file)}")
        else:
            print(f"✅ 轉換完成：{os.path.basename(output_csv_file)}")

        return True

    except Exception as e:
        if logger:
            logger.error(f"❌ 轉換失敗：{e}")
        else:
            print(f"❌ 轉換失敗：{e}")
        return False

def rename_query_file(driver, date, hour, start_minute, end_minute, logger=None):
    """
    從下載資料夾中尋找 query*.xls 檔，轉成新格式 CSV 並命名
    """
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
    new_filename = f"logger_urlLog_{date}_{start_str}-{end_str}.csv"
    new_path = os.path.join(download_dir, new_filename)

    try:
        if convert_xls_to_csv_trimmed(matched_file, new_path, logger):
            os.remove(matched_file)
            success_msg = f"✅ 檔案轉換並重新命名為：{new_filename}"
        else:
            error_msg = "❌ 檔案轉換失敗"
            if logger:
                logger.error(error_msg)
            raise Exception(error_msg)

    except Exception as e:
        error_msg = f"❌ 檔案處理失敗: {e}"
        if logger:
            logger.error(error_msg)
        raise

    if logger:
        logger.info(success_msg)
    else:
        print(success_msg)

    return new_path