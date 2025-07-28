import os
import requests
import config

# ------- 取得 Token -------
def get_graph_token():
    token_url = f"https://login.microsoftonline.com/{config.GRAPH_API_TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "client_id": config.GRAPH_API_CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": config.GRAPH_API_CLIENT_SECRET,
        "grant_type": "client_credentials"
    }

    response = requests.post(token_url, data=payload)
    response.raise_for_status()
    token = response.json()["access_token"]
    return token

# ------- 上傳檔案 -------
def upload_file_to_sharepoint(file_path, file_name, access_token):
    url = f"https://graph.microsoft.com/v1.0/drives/{config.GRAPH_API_DRIVE_ID}/items/{config.GRAPH_API_FOLDER_ID}:/{file_name}:/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/octet-stream"
    }

    with open(file_path, "rb") as f:
        response = requests.put(url, headers=headers, data=f)
        response.raise_for_status()
        print(f"✅ 成功上傳: {file_name}")
        return response.json()

# ------- 主流程 -------
if __name__ == "__main__":
    access_token = get_graph_token()

    # ✅ 測試單一檔案
    file_path = "downloads/20250711092511.csv"        # 換成你的實際檔案
    file_name = os.path.basename(file_path)

    try:
        upload_file_to_sharepoint(file_path, file_name, access_token)
    except Exception as e:
        print(f"❌ 上傳失敗: {e}")
