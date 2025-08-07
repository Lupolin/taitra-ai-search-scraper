#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
診斷 Excel 檔案轉換問題
"""

import os
import win32com.client
import pandas as pd
import config

def test_excel_file():
    """測試 Excel 檔案是否可以正常開啟"""
    xls_file = os.path.join(config.DOWNLOAD_FOLDER, "query.xls")
    
    print(f"🔍 檢查檔案：{xls_file}")
    
    # 檢查檔案是否存在
    if not os.path.exists(xls_file):
        print("❌ 檔案不存在")
        return False
    
    # 檢查檔案大小
    file_size = os.path.getsize(xls_file)
    print(f"📁 檔案大小：{file_size} bytes")
    
    if file_size == 0:
        print("❌ 檔案大小為 0")
        return False
    
    # 檢查檔案是否可讀取
    try:
        with open(xls_file, 'rb') as f:
            header = f.read(8)
            print(f"📄 檔案頭部：{header}")
    except Exception as e:
        print(f"❌ 無法讀取檔案：{e}")
        return False
    
    # 嘗試用 Excel COM 開啟
    try:
        print("🔄 嘗試建立 Excel COM 物件...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        print("✅ Excel COM 物件建立成功")
        
        print("🔄 嘗試開啟 Excel 檔案...")
        wb = excel.Workbooks.Open(xls_file)
        print("✅ Excel 檔案開啟成功")
        
        # 檢查工作表數量
        sheet_count = wb.Worksheets.Count
        print(f"📊 工作表數量：{sheet_count}")
        
        # 檢查第一個工作表的資料
        ws = wb.Worksheets(1)
        used_range = ws.UsedRange
        row_count = used_range.Rows.Count
        col_count = used_range.Columns.Count
        print(f"📈 資料範圍：{row_count} 行 x {col_count} 列")
        
        # 讀取前幾行資料
        print("📋 前 5 行資料：")
        for i in range(1, min(6, row_count + 1)):
            row_data = []
            for j in range(1, min(6, col_count + 1)):
                cell_value = ws.Cells(i, j).Value
                row_data.append(str(cell_value) if cell_value is not None else "")
            print(f"  第 {i} 行：{row_data}")
        
        wb.Close(False)
        excel.Quit()
        print("✅ Excel 檔案測試完成")
        return True
        
    except Exception as e:
        print(f"❌ Excel COM 錯誤：{e}")
        try:
            if 'excel' in locals():
                excel.Quit()
        except:
            pass
        return False

def test_pandas_read():
    """測試用 pandas 直接讀取 Excel 檔案"""
    xls_file = os.path.join(config.DOWNLOAD_FOLDER, "query.xls")
    
    print(f"\n🔄 嘗試用 pandas 讀取 Excel 檔案...")
    
    try:
        # 嘗試用 pandas 讀取
        df = pd.read_excel(xls_file, header=None)
        print(f"✅ pandas 讀取成功")
        print(f"📊 資料形狀：{df.shape}")
        print(f"📋 前 5 行資料：")
        print(df.head())
        return True
    except Exception as e:
        print(f"❌ pandas 讀取失敗：{e}")
        return False

if __name__ == "__main__":
    print("🚀 開始診斷 Excel 檔案問題")
    print("=" * 50)
    
    # 測試 Excel COM
    excel_success = test_excel_file()
    
    # 測試 pandas
    pandas_success = test_pandas_read()
    
    print("\n" + "=" * 50)
    print("📊 診斷結果：")
    print(f"Excel COM 測試：{'✅ 成功' if excel_success else '❌ 失敗'}")
    print(f"pandas 測試：{'✅ 成功' if pandas_success else '❌ 失敗'}")
    
    if not excel_success and not pandas_success:
        print("\n💡 建議：")
        print("1. 檢查檔案是否損壞")
        print("2. 檢查檔案格式是否正確")
        print("3. 嘗試重新下載檔案")
    elif excel_success and not pandas_success:
        print("\n💡 建議：使用 Excel COM 方式轉換")
    elif not excel_success and pandas_success:
        print("\n💡 建議：改用 pandas 方式轉換")
    else:
        print("\n✅ 檔案正常，兩種方式都可以使用")
