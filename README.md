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
建議建立虛擬環境後安裝相依套件：

```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

必要套件：
- `openpyxl`（讀取 Excel）

選用套件：
- `pywin32`（若希望使用 COM automation 自動啟動 Excel 並指定工作表）

## 安裝步驟
1. 取得專案原始碼。
2. 建議建立並啟用虛擬環境，安裝 `requirements.txt`。
3. 確保 `exe/` 目錄中放置目標遠端工具的可執行檔（如 `rustdesk.exe`、`AnyDesk.exe`、`TightVNC.exe`）或使用環境變數覆寫路徑。

## 使用方式
在開發環境中執行：

```bash
python Alldesk.py
```

主要功能：
- 手動輸入 ID / Password / Port 並啟動對應遠端工具。
- 由 `Alldesk.xlsx` 載入客戶端清單並建立快速按鈕以便一鍵連線。
- 在啟動前自動生成/覆寫必要的應用程式設定檔（放在 `%APPDATA%` 下），以調整視窗或預載密碼等行為。

### Excel 工作表格式建議
- sheet1: `rustdesk`（或第一張表） — 欄位：tag, id, password
- sheet2: `anydesk`（或第二張表） — 欄位：tag, id, password
- sheet3: VNC（或第三張表） — 欄位：Item, URL, Password, Port

欄位名稱會以不區分大小寫的方式嘗試配對，程式也會容錯空值或不完整列。

## 截圖 / 示範
建議將 GUI 截圖放置於 `docs/screenshot.png`，在 README 中加入：

```markdown
![Alldesk GUI](docs/screenshot.png)
```

目前倉庫未包含示範截圖；若您提供截圖檔，我可替您插入並提交更新。

## 開發者備註
- 程式在嘗試使用 COM automation（當 `pywin32` 可用）以開啟 Excel 並選取工作表；若不可用則使用系統關聯或 `excel.exe`（若在登錄中可找到）啟動檔案。
- 所有對 Excel 的讀取都採用唯讀與 data-only 模式，以降低與使用者資料衝突的風險。
- 若要打包為單一可執行檔，請確認在打包設定中包含 `exe/` 與必要資源檔（如 `vnc.vnc` 範本）。

## 授權
請於此補上專案授權資訊（例如 MIT / Apache-2.0 / Proprietary）。

---

若要我代為執行 `git add` 與 `git commit`，或將截圖插入 README，我可以協助完成。
