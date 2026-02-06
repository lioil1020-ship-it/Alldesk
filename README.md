# Alldesk

<<<<<<< HEAD
Alldesk 是一個針對 Windows 的輕量 GUI 應用程式，用於快速啟動與預先設定常見的遠端桌面客戶端（RustDesk、AnyDesk、TightVNC）。藉由集中化的客戶資料庫與自動化的設定初始化，大幅簡化了遠端連線流程，減少技術人員的手動操作步驟。
=======
Alldesk 是一個針對 Windows 的輕量 GUI 管理工具，用於快速啟動與預先設定常見的遠端桌面客戶端（RustDesk、AnyDesk、TightVNC）。

核心資料來源已改為 `Alldesk.json` 與 CSV 匯入/匯出；使用者也可在應用程式 UI 內直接新增、編輯、刪除客戶資料。

---

**主要功能**

- 使用 `Alldesk.json` 作為主要儲存格式（程式會在啟動時自動建立 `Alldesk.json` 若不存在）。
- 支援從 CSV 匯入與匯出客戶清單，用於批次管理或與其他系統整合。
- 應用程式 UI 支援：新增 / 編輯 / 刪除 客戶項目，並能即時儲存至 `Alldesk.json`。
- 啟動並預先設定外部客戶端（RustDesk、AnyDesk、TightVNC），在啟動前會自動建立或更新必要的本機設定檔（位於 `%APPDATA%`）。
- TightVNC 密碼會轉換為 TightVNC 相容格式（程式內含用於此轉換的輕量 DES 實作）。

---

**支援平台**

- Windows（程式會使用 `%APPDATA%` 與 Win32 相關行為）。

---

**需求（執行 / 開發）**

請以專案根目錄的 `requirements.txt` 為依據安裝相依套件。若需 COM 自動化或進階 Windows 操作，可選安裝：

- `pywin32` / `comtypes` / `pywinauto`（可選，若需要與 Microsoft Excel 或進階 UI 自動化）
- 其餘套件請參考 `requirements.txt` 或 `pyproject.toml`。

建立虛擬環境範例：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

---

**資料檔說明**

- `Alldesk.json`：程式主要儲存格式，包含多個區段（例如 rustdesk / anydesk / vnc），程式會在 UI 操作時讀寫該檔案。
- CSV：匯入（import）會嘗試對應常見欄位並加入到指定區段；匯出（export）會將指定區段匯成 CSV 供外部使用。

建議 CSV 欄位（依服務類型）：
- RustDesk / AnyDesk：`name, id, password`
- TightVNC：`name, host, port, password`

---

**使用說明（快速開始）**

1. 若尚無 `Alldesk.json`，程式啟動時會自動建立空結構；或手動建立並放在程式執行目錄。
2. 執行：

```powershell
python Alldesk.py
```

3. 在 UI 中新增 / 編輯 / 刪除 客戶，或使用匯入 CSV 批次新增；匯出可產生可分享的 CSV 檔。
4. 選擇連線按鈕時，程式會依項目類型（RustDesk / AnyDesk / TightVNC）預先寫入所需設定並啟動對應外部程式。

---

**安全性注意事項**

- `Alldesk.json` 或 CSV 檔可能包含明文密碼，請妥善保管與存取權限控管。
- 內建的 DES 實作僅用於 TightVNC 密碼格式轉換，不應用作安全加密或機密保存。

---

**疑難排解**

- 外部客戶端啟動失敗：檢查 `exe/` 目錄或對應環境變數（`RUSTDESK_APP`、`ANYDESK_APP`、`TIGHTVNC_APP`）是否正確。
- 匯入 CSV 欄位錯誤：確認 CSV 欄位名稱符合建議格式或手動在 UI 補齊必要欄位。

---

**授權**

參考專案根目錄的 `LICENSE` 檔案。

---

若要我代為提交並推送此變更，回覆確認即可，我會執行 commit + push。 
# Alldesk

Alldesk 是一個針對 Windows 的輕量 GUI 工具，用來從 Excel 清單快速啟動與預先設定常見的遠端桌面客戶端（RustDesk、AnyDesk、TightVNC）。
>>>>>>> a28f6d45781bebc1ca60861304681310837cb02b

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

## 打包與發行

已在專案中加入簡易打包腳本，方便在 Windows 上使用 PyInstaller 建置可執行檔：

- Windows 批次檔：[scripts/build-exe.bat](scripts/build-exe.bat)
- PowerShell：[scripts/build-exe.ps1](scripts/build-exe.ps1)

使用方式（建議先安裝開發依賴）：

```powershell
# 安裝開發依賴（如果尚未安裝）
uv sync --group dev

# 使用批次檔
.\scripts\build-exe.bat

# 或使用 PowerShell 腳本（需要允許執行策略）
.
\scripts\build-exe.ps1

# PowerShell 範例：先清理再打包
.\scripts\build-exe.ps1 -Clean
```

生成的可執行檔會放在 `dist/` 資料夾，請確認 `lioil.ico` 與 `exe/` 資料夾存在於專案根目錄。
