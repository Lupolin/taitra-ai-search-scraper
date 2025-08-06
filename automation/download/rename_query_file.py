import os
import time
import glob
import config
import pandas as pd
import win32com.client
from datetime import date, datetime

def convert_xls_to_csv_trimmed(xls_file, output_csv_file, logger=None):
    """
    將 .xls 檔案轉換為 .csv 並刪除前 6 列
    """
    try:
        # 建立 Excel COM 物件
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        if logger:
            logger.info("✅ Excel COM 物件建立成功")
        else:
            print("✅ Excel COM 物件建立成功")
        
        # 開啟 .xls 並轉存成 .csv，指定編碼
        wb = excel.Workbooks.Open(xls_file)
        # 使用 FileFormat=6 (CSV) 並指定編碼
        temp_csv = xls_file.replace('.xls', '_temp.csv')
        wb.SaveAs(temp_csv, FileFormat=6, Local=True)  # Local=True 使用系統預設編碼
        wb.Close(False)
        
        # 用 pandas 刪除前 6 列，嘗試不同的編碼
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
            raise Exception("無法使用任何編碼讀取 CSV 檔案")
        
        # 檢查是否有足夠的列數
        if len(df) <= 6:
            if logger:
                logger.warning(f"警告：檔案只有 {len(df)} 列，無法刪除前 6 列")
            else:
                print(f"⚠️ 警告：檔案只有 {len(df)} 列，無法刪除前 6 列")
            df_trimmed = df
        else:
            df_trimmed = df.iloc[6:].reset_index(drop=True)
        
        # 儲存為 UTF-8 BOM 格式，確保 Excel 能正確顯示中文
        df_trimmed.to_csv(output_csv_file, index=False, header=False, encoding='utf-8-sig')
        
        # 清理臨時檔案
        if os.path.exists(temp_csv):
            os.remove(temp_csv)
        
        # 清理 Excel 物件
        try:
            excel.Quit()
        except:
            pass
        
        if logger:
            logger.info(f"✅ 轉換完成：{os.path.basename(output_csv_file)}")
            logger.info(f"原始列數：{len(df)}，處理後列數：{len(df_trimmed)}")
        else:
            print(f"✅ 轉換完成：{os.path.basename(output_csv_file)}")
            print(f"原始列數：{len(df)}，處理後列數：{len(df_trimmed)}")
        
        return True
        
    except Exception as e:
        if logger:
            logger.error(f"❌ 轉換失敗：{e}")
        else:
            print(f"❌ 轉換失敗：{e}")
        return False

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
    
    # 新的檔案名稱（CSV 格式）
    new_filename = f"logger_url_{date}_{start_str}-{end_str}.csv"
    new_path = os.path.join(download_dir, new_filename)

    try:
        # 轉換 XLS 到 CSV 並刪除前 6 列
        if convert_xls_to_csv_trimmed(matched_file, new_path, logger):
            # 轉換成功後刪除原始 XLS 檔案
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
