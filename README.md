# Alldesk

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

核心目標：減少支援人員在現場或遠端操作時的手動設定步驟，並在啟動外部客戶端前自動建立或更新必要的本機設定檔，以達到較一致的連線行為。

---

**主要功能**

- 由 Excel（建議檔名 `Alldesk.xlsx`）讀取連線清單，動態生成快速連線按鈕。
- 支援預先寫入或調整應用程式設定（寫入 `%APPDATA%` 下對應的設定檔）以控制：
  - RustDesk：可預載 ID/密碼與設定檔
  - AnyDesk：可寫入 `user.conf` 並以命令列傳入密碼
  - TightVNC：產生 `vnc.vnc` 選項檔並將密碼轉換為 TightVNC 相容格式（程式內含專為此用途的輕量 DES 實作）
- 若系統安裝 Microsoft Excel 且安裝 `pywin32`，程式會嘗試使用 COM automation 讀取 Excel；否則以資料匯入或系統預設方式開啟檔案。

---

**支援平台**

- Windows（程式會使用 `%APPDATA%` 與 Win32 相關行為）。

---

**需求（開發 / 執行）**

建議使用 Python 3.9+ 並在虛擬環境中安裝相依：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

核心套件（請以專案中的 `requirements.txt` 或 `pyproject.toml` 為準）：
- `openpyxl`：讀取 Excel
- `pywin32`、`comtypes`、`pywinauto`（可選）：若需 COM automation 或進階 Windows 自動化

---

**快速開始（開發模式）**

在專案根目錄執行：

```powershell
python Alldesk.py
```

執行前：
- 將 `Alldesk.xlsx` 放在與 `Alldesk.py` 相同目錄，或與可執行檔放在同資料夾。
- 若欲使用本專案內的外部執行檔，請將它們放到 `exe/` 資料夾（例如 `exe/rustdesk.exe`、`exe/AnyDesk.exe`、`exe/TightVNC.exe`），或透過環境變數覆寫路徑：

```powershell
$env:RUSTDESK_APP = 'C:\path\to\rustdesk.exe'
$env:ANYDESK_APP = 'C:\path\to\AnyDesk.exe'
$env:TIGHTVNC_APP = 'C:\path\to\vncviewer.exe'
python Alldesk.py
```

---

**Excel 清單建議格式**

程式會嘗試以不區分大小寫的欄位名匹配資料，建議三張常用工作表（或以欄位區分）：

- RustDesk 表（sheet 名稱可自由）： `設備名稱`, `ID`, `密碼`
- AnyDesk 表： `設備名稱`, `ID`, `密碼`
- TightVNC 表： `設備名稱`, `URL` 或 `HOST`, `埠`, `密碼`

程式具備欄位容錯處理，但依照上述格式建立能取得最穩定的結果。

---

**安全性說明**

- `Alldesk.xlsx` 可能包含明文密碼，請妥善管理與傳輸該檔案。
- 程式內的 DES 實作僅用於將密碼轉為 TightVNC 相容格式，非通用加密庫，勿用於安全機密存放。

---

**打包建議**

- 可使用 PyInstaller 產生 `onefile` 或 `onedir`：

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed Alldesk.py
```

- 注意 `onefile` 與 `onedir` 在資源解析與 single-instance 行為上可能不同；請測試兩種模式以決定合適打包方式。

---

**疑難排解**

- 外部客戶端啟動失敗：檢查 `exe/` 目錄或對應環境變數是否正確。
- Excel 無法以 COM 開啟：確認本機是否安裝 Microsoft Excel 並安裝 `pywin32`。
- 打包後行為異常：檢查 PyInstaller 的參數與資源包含設定。

---

**授權**

參考專案根目錄的 `LICENSE` 檔案。

---

若要我幫你：

- 以 `git` 將新的 `README.md` commit 並 push（我可以代為執行），或
- 將 README 翻成英文 / 補上範例圖片與操作教學。

# Alldesk

## 專案簡介
Alldesk 是一個針對 Windows 的輕量 GUI 工具，主要用途為快速啟動並協助配置三種常見的遠端桌面工具：

- `RustDesk`：透過事先在 `%APPDATA%` 下寫入設定（例如 `RustDesk2.toml` 與 `peers/<id>.toml`）來預載密碼與視窗設定，然後啟動 RustDesk 連線。
- `AnyDesk`：在啟動前寫入 `%APPDATA%\AnyDesk\user.conf` 以設定 viewmode，並以命令列管道傳入密碼啟動 AnyDesk。
- `TightVNC`：根據第 3 張工作表的 host/port/password 生成 `vnc.vnc` 選項檔（輸出到 `./exe/vnc.vnc`），並啟動 TightVNC；若有密碼則以 TightVNC 相容格式加密（程式內含輕量 DES 實作）。

程式以 `Alldesk.xlsx` 為唯讀資料來源，提供手動輸入欄位與由 Excel 載入的快速連線按鈕，並在啟動外部程式前自動處理所需的本機設定檔。

> 注意：內含的 DES 實作僅用於 TightVNC 密碼格式轉換，非通用密碼庫或安全加密套件。

## 支援平台
- Windows（程式會使用 `%APPDATA%` 與可能的 COM Automation for Excel）

## 需求
建立虛擬環境後安裝相依套件：

```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

專案所需套件（請使用 `requirements.txt` 安裝以確保版本一致）：
- comtypes==1.4.14
- et-xmlfile==2.0.0
- openpyxl==3.1.5
- packaging==25.0
- python-dateutil==2.9.0.post0
- pywin32==311
- pywinauto==0.6.9
- setuptools==80.9.0
- six==1.17.0
- tzdata==2025.3
- wheel==0.45.1

說明：
- `openpyxl`：用於讀取 `Alldesk.xlsx`（必備）。
- `pywin32` / `comtypes` / `pywinauto`：在 Windows 下提供 COM Automation、系統互動與進階 GUI 自動化（若你需要程式自動使用 Excel 或更進階的 Windows 操作，請安裝這些套件）。
- 其他列出套件多為依賴或打包/相容性相關套件，建議直接使用 `pip install -r requirements.txt` 一次安裝。
# Alldesk

Alldesk 是一款以 Windows 為目標的輕量 GUI 工具，用於從 Excel 清單快速啟動並協助配置常見的遠端桌面客戶端（RustDesk、AnyDesk、TightVNC）。

核心設計理念：簡化遠端支援流程、減少手動設定步驟，並在必要時自動建立或更新應用程式的本機設定檔以達到一致的連線行為。

---

## 主要功能

- 由 `Alldesk.xlsx` 讀取客戶端清單（唯讀、data-only 模式），自動產生快速連線按鈕。
- 啟動 RustDesk / AnyDesk / TightVNC，並在啟動前自動寫入必要的應用設定（例如 `%APPDATA%` 下的設定檔）。
- TightVNC 密碼會依 TightVNC 格式做轉換（內含一個僅用於此轉換的輕量 DES 實作）。
- 支援以 COM automation 嘗試控制 Excel（若系統安裝 pywin32），否則使用系統關聯開啟檔案。

---

## 支援平台

- Windows（測試與開發皆針對 Windows；程式會讀寫 `%APPDATA%`，並使用 Win32 API 沿用系統行為）。

---

## 需求

請使用 Python 3.9+（或與專案相容的 Python 版本），並於虛擬環境中安裝相依：

```powershell
python -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
pip install -r requirements.txt
```

推薦安裝套件（請以專案內 `requirements.txt` 為準）：

- openpyxl（讀取 Excel）
- pywin32 / comtypes / pywinauto（可選，提供 COM 與 UI 自動化能力）

---

## 快速開始

開發模式下執行：

```powershell
python Alldesk.py
```

建議操作流程：

1. 將 `Alldesk.xlsx` 放在與程式相同的資料夾，或在打包後與可執行檔同目錄。
2. 若使用內附外部工具（如 `rustdesk.exe`、`AnyDesk.exe`、`TightVNC.exe`），請放到專案 `exe/` 資料夾，或以環境變數覆寫路徑：

- `RUSTDESK_APP`：RustDesk 可執行檔路徑
- `ANYDESK_APP`：AnyDesk 可執行檔路徑
- `TIGHTVNC_APP`：TightVNC 可執行檔路徑

例如（PowerShell）：

```powershell
$env:RUSTDESK_APP = 'C:\\path\\to\\rustdesk.exe'
python Alldesk.py
```

---

## Excel 清單格式（建議）

程式會嘗試以不區分大小寫的欄位名稱匹配資料，常見表格配置：

- Sheet1 / rustdesk：`設備名稱`, `ID`, `密碼`
- Sheet2 / anydesk：`設備名稱`, `ID`, `密碼`
- Sheet3 / vnc / tightvnc：`設備名稱`, `URL`, `密碼`, `埠`

程式內有容錯機制，可處理缺欄或非標準欄名，但建議以上述格式建立以獲得最佳體驗。

---

## 打包與部署建議

- 使用 PyInstaller 建立單一可執行檔（onefile）或 onedir：

	```powershell
	pip install pyinstaller
	pyinstaller --noconfirm --onefile --windowed Alldesk.py
	```

- 注意：onedir 與 onefile 在某些行為上不同（例如資源解析與 single-instance）；專案內有處理 onedir 時將外部執行檔複製至暫存以避免 single-instance 攔截的機制。

---

## 安全性與隱私

- `Alldesk.xlsx` 可能包含明文密碼；請務必妥善保管、不將含密的檔案公開或上傳到無信任的遠端儲存。
- 內部的 DES 實作僅供 TightVNC 密碼格式轉換使用，不應視為安全加密工具。

---

## 疑難排解

- 無法啟動外部客戶端：檢查 `exe/` 目錄或對應環境變數是否正確。
- Excel 無法自動開啟或選擇工作表：確認是否已安裝 Microsoft Excel 與 pywin32（若無 pywin32，程式會使用系統預設開啟）。
- 打包後行為異常：請確認 PyInstaller 的打包參數並測試 onefile 與 onedir 兩種模式。

---

## 授權

請參考專案中的 LICENSE 檔案。