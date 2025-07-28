import os
import shutil

def clear_downloads_folder():
    folder_path = os.path.join(os.getcwd(), "downloads")

    if not os.path.exists(folder_path):
        print("❌ downloads 資料夾不存在")
        return

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)  # 刪除檔案或捷徑
                print(f"🗑 已刪除檔案：{filename}")
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # 刪除子資料夾
                print(f"🗑 已刪除資料夾：{filename}")
        except Exception as e:
            print(f"⚠️ 刪除 {filename} 時發生錯誤：{e}")

if __name__ == "__main__":
    clear_downloads_folder()
