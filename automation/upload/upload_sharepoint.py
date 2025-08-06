import os
import glob
import requests
from utils.logger import setup_logging
import config

def get_graph_token():
    """取得 Graph API 存取權杖"""
    token_url = f"https://login.microsoftonline.com/{config.GRAPH_API_TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "client_id": config.GRAPH_API_CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": config.GRAPH_API_CLIENT_SECRET,
        "grant_type": "client_credentials"
    }
    response = requests.post(token_url, data=payload)
    response.raise_for_status()
    return response.json()["access_token"]

def upload_file(file_path, file_name, token, logger):
    """上傳單一檔案到 SharePoint"""
    url = f"https://graph.microsoft.com/v1.0/drives/{config.GRAPH_API_DRIVE_ID}/items/{config.GRAPH_API_FOLDER_ID}:/{file_name}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream"
    }

    try:
        with open(file_path, "rb") as f:
            response = requests.put(url, headers=headers, data=f)
            response.raise_for_status()
        logger.info(f"✅ 成功上傳: {file_name}")
        return True
    except Exception as e:
        logger.error(f"❌ 上傳失敗 {file_name}: {e}")
        return False

def scan_files(folder, patterns=["*.csv"]):
    """掃描指定資料夾中符合格式的檔案"""
    files = []
    for pattern in patterns:
        files.extend(glob.glob(os.path.join(folder, pattern)))
    return sorted(files)

def upload_files(file_paths, logger, token):
    """批次上傳檔案"""
    success, failed = 0, 0
    for path in file_paths:
        if not os.path.exists(path):
            logger.error(f"❌ 檔案不存在: {path}")
            failed += 1
            continue
        name = os.path.basename(path)
        logger.info(f"📤 正在上傳: {name}")
        if upload_file(path, name, token, logger):
            success += 1
        else:
            failed += 1
    return success, failed

def log_summary(logger, success, failed):
    total = success + failed
    logger.info("=== 上傳完成 ===")
    logger.info(f"📊 總計: {total}，✅ 成功: {success}，❌ 失敗: {failed}")

def upload_temp_files_to_sharepoint(logger=None):
    """主流程：上傳 temp 資料夾中所有檔案到 SharePoint"""
    logger = logger or setup_logging("SharePoint 上傳")
    logger.info("=== 開始上傳 temp 資料夾中的檔案 ===")

    files = scan_files(config.DOWNLOAD_FOLDER)
    if not files:
        logger.warning("⚠️ 找不到任何檔案")
        return {"success": 0, "failed": 0, "total": 0}

    try:
        token = get_graph_token()
        logger.info("✅ 成功取得 Graph API 存取權杖")
    except Exception as e:
        logger.error(f"❌ 存取權杖取得失敗: {e}")
        return {"success": 0, "failed": len(files), "total": len(files)}

    success, failed = upload_files(files, logger, token)
    log_summary(logger, success, failed)

    return {"success": success, "failed": failed, "total": len(files)}