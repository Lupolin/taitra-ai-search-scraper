# AI 上網統計 RPA

這是一個自動化的 RPA（Robotic Process Automation）系統，用於從網站下載 AI 上網統計資料並上傳至 SharePoint。

## 專案結構

```
AI上網統計RPA/
├── main_download.py             # 主程式：下載階段
├── main_upload.py               # 主程式：上傳階段
├── requirements.txt             # 套件清單
├── .env                         # 機密帳密設定
├── config.py                    # 網址、資料夾等基本設定
│
├── automation/                  # Selenium 操作模組
│   ├── browser.py               # Chrome driver 設定
│   ├── login.py                 # 登入邏輯
│   ├── download_one.py          # 單次時間區段下載操作
│   └── upload_sharepoint.py     # 上傳至 SharePoint
│
├── utils/                       # 輔助工具模組
│   ├── time_utils.py            # 產生每 10 分鐘的時間區段
│   ├── file_utils.py            # 檢查下載檔案、命名、確認成功
│   └── logger.py                # 可選的 log 模組
│
└── downloads/                   # 實體下載資料存放區（自動建立）
    └── 2025-07-23/              # 每日資料夾，內含 60 檔
```

## 功能特色

- **自動化下載**: 自動從網站下載每 10 分鐘的統計資料
- **時間區段管理**: 自動產生 144 個時間區段（00:00-23:59）
- **檔案管理**: 自動整理檔案到日期資料夾
- **SharePoint 整合**: 自動上傳檔案至 SharePoint
- **錯誤處理**: 完整的重試機制和錯誤處理
- **日誌記錄**: 詳細的操作日誌
- **彈性配置**: 可自訂各種設定參數

## 安裝需求

### 系統需求

- Python 3.8 或以上
- Windows 10/11
- Chrome 瀏覽器
- 網路連線

### Python 套件

```bash
pip install -r requirements.txt
```

## 設定說明

### 1. 環境變數設定 (.env)

複製 `.env` 檔案並填入實際的帳密資訊：

```env
# 網站登入憑證
WEBSITE_USERNAME=your_actual_username
WEBSITE_PASSWORD=your_actual_password

# SharePoint 設定
SHAREPOINT_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
SHAREPOINT_CLIENT_ID=your_client_id
SHAREPOINT_CLIENT_SECRET=your_client_secret
SHAREPOINT_TENANT_ID=your_tenant_id

# 下載設定
DOWNLOAD_FOLDER=downloads
LOG_LEVEL=INFO
```

### 2. 配置檔案 (config.py)

根據實際需求修改 `config.py` 中的設定：

```python
# 網站設定
WEBSITE_URL = "https://your-statistics-website.com"
LOGIN_URL = f"{WEBSITE_URL}/login"
DOWNLOAD_URL = f"{WEBSITE_URL}/download"

# 時間設定
TIME_INTERVAL_MINUTES = 10  # 每10分鐘一個區段
START_TIME = "00:00"
END_TIME = "23:59"

# SharePoint 設定
SHAREPOINT_FOLDER = "AI上網統計"
SHAREPOINT_FILE_PREFIX = "AI統計_"
```

## 使用方法

### 下載階段

#### 下載今天的資料

```bash
python main_download.py
```

#### 下載指定日期的資料

```bash
python main_download.py 2025-01-15
```

### 上傳階段

#### 上傳今天的檔案

```bash
python main_upload.py
```

#### 上傳指定日期的檔案

```bash
python main_upload.py 2025-01-15
```

#### 列出 SharePoint 中的檔案

```bash
python main_upload.py --list
```

#### 刪除 SharePoint 中的檔案

```bash
python main_upload.py --delete filename.xlsx
```

#### 上傳指定檔案

```bash
python main_upload.py file1.xlsx file2.xlsx
```

## 工作流程

### 下載流程

1. 啟動 Chrome 瀏覽器
2. 自動登入網站
3. 產生 144 個時間區段（每 10 分鐘）
4. 逐一下載每個時間區段的資料
5. 驗證下載的檔案
6. 整理檔案到日期資料夾
7. 清理臨時檔案

### 上傳流程

1. 連接到 SharePoint
2. 掃描指定日期的檔案
3. 逐一上傳檔案
4. 記錄上傳結果
5. 生成上傳報告

## 錯誤處理

### 重試機制

- 下載失敗時自動重試（預設 3 次）
- 每次重試間隔 5 秒
- 記錄重試次數和失敗原因

### 錯誤類型

- 網路連線錯誤
- 登入失敗
- 檔案下載失敗
- SharePoint 連線錯誤
- 檔案格式錯誤

## 日誌系統

### 日誌等級

- DEBUG: 詳細除錯資訊
- INFO: 一般操作資訊
- WARNING: 警告訊息
- ERROR: 錯誤訊息
- CRITICAL: 嚴重錯誤

### 日誌檔案

- 位置: `rpa.log`
- 格式: `時間 - 模組 - 等級 - 訊息`
- 編碼: UTF-8

## 檔案命名規則

### 下載檔案

```
ai_stats_YYYYMMDD_HHMM-HHMM.xlsx
```

範例：

```
ai_stats_20250115_0000-0010.xlsx
ai_stats_20250115_0010-0020.xlsx
```

### SharePoint 檔案

```
AI統計_YYYYMMDD_01_filename.xlsx
AI統計_YYYYMMDD_02_filename.xlsx
```

## 監控與維護

### 定期檢查項目

1. 日誌檔案大小
2. 下載資料夾空間
3. SharePoint 連線狀態
4. 網站登入狀態

### 清理作業

- 定期清理舊的日誌檔案
- 清理下載資料夾中的臨時檔案
- 備份重要的配置檔案

## 故障排除

### 常見問題

#### 1. Chrome Driver 錯誤

```bash
# 重新安裝 webdriver-manager
pip uninstall webdriver-manager
pip install webdriver-manager
```

#### 2. 登入失敗

- 檢查 `.env` 檔案中的帳密是否正確
- 確認網站是否正常運作
- 檢查網路連線

#### 3. SharePoint 連線失敗

- 檢查 SharePoint 設定是否正確
- 確認應用程式權限
- 檢查網路連線

#### 4. 檔案下載失敗

- 檢查下載資料夾權限
- 確認磁碟空間是否足夠
- 檢查檔案格式設定

## 開發指南

### 新增功能

1. 在適當的模組中新增功能
2. 更新配置檔案
3. 新增測試案例
4. 更新文件

### 修改設定

1. 修改 `config.py` 中的相關設定
2. 測試修改後的設定
3. 更新相關文件

## 授權

本專案僅供內部使用，請勿外流。

## 聯絡資訊

如有問題或建議，請聯絡開發團隊。
