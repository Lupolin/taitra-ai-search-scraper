import os
from datetime import datetime, timedelta
import requests
import config

def generate_expected_filenames() -> list[str]:
    filenames = []
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
    for hour in range(8, 18):
        for i in range(6):
            start_min = i * 10
            end_min = start_min + 9
            start_str = f"{hour:02d}{start_min:02d}"
            end_str = f"{hour:02d}{end_min:02d}"
            interval = f"{start_str}-{end_str}"
            filename = f"logger_url_{yesterday}_{interval}.csv"
            filenames.append(filename)
    return filenames

def check_download_files() -> str:
    """
    檢查昨天所有應該下載的檔案是否存在，並回傳格式化文字（含統計）
    """
    download_folder = config.DOWNLOAD_FOLDER
    expected_files = generate_expected_filenames()

    log_date = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    lines = [f"{log_date} AI上網統計紀錄：\n"]

    success_count = 0
    fail_count = 0

    for filename in expected_files:
        full_path = os.path.join(download_folder, filename)
        interval = filename.replace(".csv", "")
        exists = os.path.exists(full_path)
        status = "✅" if exists else "❌"
        lines.append(f"{interval}-{status}")
        if exists:
            success_count += 1
        else:
            fail_count += 1

    lines.append("")  # 空行
    lines.append(f"✅ 成功：{success_count}，❌ 失敗：{fail_count}")
    return "\n".join(lines)

def notify_teams_result(message_text: str) -> None:
    """
    發送訊息給 Power Automate，格式為 { "text": "..." }
    """
    webhook_url = config.POWER_AUTOMATE_WEBHOOK_URL
    payload = { "text": message_text }
    try:
        response = requests.post(webhook_url, json=payload)
        if response.status_code in [200, 202]:
            print("✅ 成功發送給 Power Automate")
        else:
            print(f"❌ 發送失敗: {response.status_code}\n{response.text}")
    except Exception as e:
        print(f"❌ 發送發生錯誤：{e}")

def report_download_status():
    """
    主入口：檢查下載檔案 → 發送執行結果
    """
    message = check_download_files()
    notify_teams_result(message)