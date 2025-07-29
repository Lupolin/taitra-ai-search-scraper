# AI Tool Search - 自動化資料下載與上傳工具

一個基於 Python 和 Selenium 的自動化工具，用於從網站下載 Excel 檔案並自動上傳到 SharePoint。

## 🚀 功能特色

- **自動化瀏覽器操作**: 使用 Chrome 瀏覽器自動化登入和下載
- **時間區段處理**: 支援按時間區段（每 10 分鐘）下載資料
- **檔案重新命名**: 自動將下載的檔案重新命名為指定格式
- **SharePoint 上傳**: 自動將檔案上傳到 SharePoint 指定資料夾
- **完整日誌記錄**: 詳細的操作日誌，方便追蹤和除錯
- **錯誤處理**: 完善的錯誤處理和重試機制

## 📁 專案結構

```
ai_tool_search/
├── automation/           # 自動化模組
│   ├── browser_chrome.py    # Chrome 瀏覽器管理
│   ├── login.py             # 網站登入功能
│   ├── download_excel.py    # Excel 檔案下載
│   ├── rename_query_file.py # 檔案重新命名
│   ├── upload_sharepoint.py # SharePoint 上傳
│   ├── clear_folder.py      # 資料夾清理
│   └── captcha_solver.py    # 驗證碼處理
├── utils/               # 工具模組
│   └── logger.py           # 日誌管理
├── temp/                # 暫存檔案資料夾
├── config.py            # 配置檔案
├── main.py              # 主程式
├── requirements.txt     # Python 依賴套件
└── README.md           # 專案說明
```

## 🛠️ 安裝步驟

### 1. 克隆專案

```bash
git clone <repository-url>
cd ai_tool_search
```

### 2. 建立虛擬環境

```bash
python -m venv venv
```

### 3. 啟動虛擬環境

```bash
# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

### 4. 安裝依賴套件

```bash
pip install -r requirements.txt
```

### 5. 設定環境變數

建立 `.env` 檔案並填入以下設定：

```env
# 網站設定
WEBSITE_URL=https://your-website.com
WEBSITE_USERNAME=your_username
WEBSITE_PASSWORD=your_password

# Graph API 設定 (SharePoint 上傳用)
GRAPH_API_CLIENT_ID=your_client_id
GRAPH_API_CLIENT_SECRET=your_client_secret
GRAPH_API_TENANT_ID=your_tenant_id
GRAPH_API_DRIVE_ID=your_drive_id
GRAPH_API_FOLDER_ID=your_folder_id

# Azure OpenAI 設定 (可選)
AZURE_OPENAI_ENDPOINT=your_openai_endpoint
AZURE_OPENAI_API_KEY=your_openai_api_key
AZURE_OPENAI_DEPLOYMENT_NAME=your_deployment_name

# SharePoint 設定
SHAREPOINT_SITE_URL=your_sharepoint_site_url
```

## 🚀 使用方法

### 基本使用

```bash
python main.py
```

### 功能流程

1. **啟動瀏覽器**: 自動開啟 Chrome 瀏覽器
2. **網站登入**: 使用設定檔中的帳號密碼登入
3. **下載資料**: 按時間區段（08:00-18:00，每 10 分鐘）下載 Excel 檔案
4. **檔案處理**: 自動重新命名下載的檔案
5. **上傳 SharePoint**: 將檔案上傳到指定的 SharePoint 資料夾
6. **清理暫存**: 清理暫存檔案和資料夾

## ⚙️ 配置說明

### 時間設定

- `START_TIME`: 開始時間 (預設: "08:00")
- `END_TIME`: 結束時間 (預設: "18:00")
- `TIME_INTERVAL_MINUTES`: 時間間隔 (預設: 10 分鐘)

### 瀏覽器設定

- `BROWSER_HEADLESS`: 是否隱藏瀏覽器視窗 (預設: False)
- `BROWSER_TIMEOUT`: 瀏覽器操作超時時間 (預設: 30 秒)
- `DOWNLOAD_TIMEOUT`: 下載超時時間 (預設: 300 秒)

### 重試設定

- `MAX_RETRIES`: 最大重試次數 (預設: 3 次)
- `RETRY_DELAY`: 重試間隔時間 (預設: 5 秒)

## 📝 日誌記錄

程式會自動產生詳細的日誌檔案 `rpa.log`，包含：

- 操作步驟記錄
- 錯誤訊息
- 執行時間統計
- 檔案處理狀態

## 🔧 故障排除

### 常見問題

1. **Chrome 驅動程式問題**

   - 確保已安裝最新版本的 Chrome 瀏覽器
   - 程式會自動下載對應的 ChromeDriver

2. **登入失敗**

   - 檢查 `.env` 檔案中的帳號密碼是否正確
   - 確認網站 URL 是否可正常存取

3. **下載超時**

   - 調整 `DOWNLOAD_TIMEOUT` 設定
   - 檢查網路連線狀況

4. **SharePoint 上傳失敗**
   - 確認 Graph API 設定是否正確
   - 檢查 SharePoint 權限設定

## 🤝 貢獻指南

1. Fork 本專案
2. 建立功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交變更 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

## 📄 授權條款

本專案採用 MIT 授權條款 - 詳見 [LICENSE](LICENSE) 檔案

## 📞 聯絡資訊

如有問題或建議，請透過以下方式聯絡：

- 建立 Issue
- 發送 Email

---

**注意**: 使用前請確保已正確設定所有必要的環境變數，並確認相關服務的存取權限。
