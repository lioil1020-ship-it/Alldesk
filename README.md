# Alldesk

Alldesk 是一個針對 Windows 的輕量 GUI 應用程式，用於快速啟動與預先設定常見的遠端桌面客戶端（RustDesk、AnyDesk、TightVNC）。藉由集中化的客戶資料庫與自動化的設定初始化，大幅簡化了遠端連線流程，減少技術人員的手動操作步驟。

## 核心特性

- **JSON 檔案庫**：使用 `Alldesk.json` 作為單一真實來源（SSOT），集中管理所有遠端連線設定
- **批量資料匯入/匯出**：支援 CSV 格式，便於與其他系統集成或分享設定
- **圖形化管理介面**：直接在應用程式 UI 內新增、編輯、刪除客戶資料，實時同步至資料庫
- **自動設定初始化**：在啟動遠端客戶端前，自動建立或更新 `%APPDATA%` 下的相關設定檔
- **密碼加密轉換**：內建輕量級 DES 實作，用於 TightVNC 密碼的安全格式轉換

## 支援的遠端工具

| 工具 | 功能 | 配置位置 |
|------|------|---------|
| **RustDesk** | 啟動連線、預設檢視模式、自訂 rendezvous server | `%APPDATA%\RustDesk\config` |
| **AnyDesk** | 啟動連線、視圖模式設定 | `%APPDATA%\AnyDesk` |
| **TightVNC** | 啟動連線、主機/埠/密碼配置 | 動態生成 `vnc.vnc` 檔 |

## 系統需求

- **作業系統**：Windows 7 或更新版本
- **Python 版本**：Python 3.12 或更新版本
- **必要套件**：見 `requirements.txt`

### 依賴套件

```
comtypes==1.4.14
pywin32==311
pywinauto==0.6.9
setuptools==80.9.0
```

## 安裝與啟動

### 1. 環境設定

建立並啟動虛擬環境：

```powershell
# 建立虛擬環境
python -m venv .venv

# 啟動虛擬環境
.\.venv\Scripts\Activate.ps1

# 安裝依賴套件
pip install -r requirements.txt
```

### 2. 執行應用程式

```powershell
python Alldesk.py
```

首次啟動時，應用程式會自動在執行目錄建立 `Alldesk.json` 檔案。

## 資料格式

### Alldesk.json 結構

```json
{
  "rustdesk": [
    {
      "tag": "生產環境伺服器 1",
      "id": "123456789",
      "pwd": "password123",
      "port": ""
    }
  ],
  "anydesk": [
    {
      "tag": "客戶端 A",
      "id": "987654321",
      "pwd": "securepass",
      "port": ""
    }
  ],
  "tightvnc": [
    {
      "tag": "監控伺服器",
      "id": "192.168.1.100",
      "pwd": "vncpass",
      "port": "5900"
    }
  ]
}
```

### CSV 匯入/匯出格式

推薦的 CSV 欄位結構（使用 UTF-8 編碼）：

```csv
tag,id,pwd,port
設備名稱,連線ID或主機,密碼,埠號(VNC使用)
```

**範例 - RustDesk CSV**：
```csv
tag,id,pwd,port
辦公室電腦,123456789,pass123,
遠端伺服器,987654321,securepass,
```

**範例 - TightVNC CSV**：
```csv
tag,id,pwd,port
監控系統,192.168.1.50,vncpass,5900
備份伺服器,10.0.0.30,bkpass,5901
```

## 使用指南

### 新增遠端連線

1. 在對應的分頁（RustDesk / AnyDesk / TightVNC）中，**右鍵點擊空白區域**
2. 選擇「新增客戶」
3. 填入必要資訊：設備名稱、連線 ID、密碼、埠號（VNC 需要）
4. 點擊「儲存」

### 編輯連線設定

1. **右鍵點擊**現有的連線按鈕
2. 選擇「編輯客戶」
3. 修改所需的欄位
4. 點擊「儲存」

### 刪除連線

1. **右鍵點擊**連線按鈕
2. 選擇「刪除客戶」
3. 確認刪除

### 批量匯入 CSV

1. 準備 CSV 檔案（遵守上述格式）
2. 在對應分頁點擊「匯入」按鈕
3. 選擇 CSV 檔案
4. 確認匯入（將覆蓋現有資料）

### 批量匯出 CSV

1. 在對應分頁點擊「匯出」按鈕
2. 指定匯出檔案位置
3. 應用程式將產生 CSV 檔案供分享或備份

### 發起遠端連線

- **直接點擊**連線按鈕即可立即啟動對應的遠端工具
- 程式會自動傳遞 ID、密碼等參數
- 密碼輸入會透過多層機制嘗試（UIA、剪貼簿、鍵盤模擬等）

## 環境變數配置

可透過環境變數自訂遠端工具的執行檔路徑：

```powershell
# 設定環境變數
$env:RUSTDESK_APP = "C:\Program Files\RustDesk\rustdesk.exe"
$env:ANYDESK_APP = "C:\Program Files\AnyDesk\AnyDesk.exe"
$env:TIGHTVNC_APP = "C:\Program Files\TightVNC\TightVNC.exe"

# 執行應用程式
python Alldesk.py
```

## 安全性考量

⚠️ **重要**：

- `Alldesk.json` 和 CSV 檔案可能包含明文密碼，請確保：
  - 檔案存放位置的存取權限受到限制
  - 不要將包含密碼的檔案提交至公開版控系統
  - 定期檢查和更新密碼

- 內建 DES 實作僅用於 TightVNC 密碼格式轉換，**不適合用作通用加密方案**

### TightVNC 範本檔

程式會根據 `exe/vnc.vnc` 範本檔動態生成連線設定。如需自訂，編輯此檔案並調整相關參數。

## 打包為可執行檔

使用 PyInstaller 將應用程式打包為 Windows 可執行檔：

```powershell
pip install pyinstaller>=6.18.0
pyinstaller --onefile --windowed --icon=lioil.ico Alldesk.py
```

生成的可執行檔位於 `dist/` 資料夾。

## 常見問題

### Q: 啟動遠端工具失敗

**A**: 檢查以下項目：
- 遠端工具是否已正確安裝
- `exe/` 目錄或環境變數中的路徑是否正確指向可執行檔
- 防毒軟體是否攔截了程式執行

### Q: CSV 匯入後資料未更新

**A**: 
- 確認 CSV 檔案使用 UTF-8 編碼
- 檢查欄位名稱是否為 `tag, id, pwd, port`
- 至少 `tag` 或 `id` 欄位不能為空

### Q: 密碼輸入失敗

**A**:
- 某些 VNC 用戶端可能需要手動輸入密碼
- 檢查密碼是否包含特殊字元（某些系統可能不支援）
- 嘗試透過應用程式 UI 直接編輯和測試連線

## 授權

本專案遵循 LICENSE 檔案中的授權條款。

## 技術棧

- **GUI 框架**：Tkinter
- **密碼加密**：內建輕量 DES 演算法
- **Windows 整合**：Win32 API、comtypes、pywinauto
- **資料序列化**：JSON、CSV

## 貢獻與支援

如發現 Bug 或有功能建議，歡迎提交 issue 或 PR。