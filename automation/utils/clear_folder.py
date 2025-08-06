import os
import shutil
import config

def clear_temp_folder():
    folder_path = config.DOWNLOAD_FOLDER

    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)
        print(f"🗑 已整個刪除資料夾：{folder_path}")

    os.makedirs(folder_path)  # 重建資料夾
    print(f"📁 已重建空的 temp 資料夾")

