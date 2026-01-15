# Alldesk

簡介
-
Alldesk 是一個以 GUI (tkinter) 提供的遠端支援啟動器，整合 RustDesk、AnyDesk 與 TightVNC 的快速啟動流程。設定由 `Alldesk.xlsx` 驅動（Excel 為單一來源），啟動前會自動產生或覆寫目標應用程式的設定檔以預設視圖/密碼等參數。

主要功能
-
- RustDesk：讀取 `rustdesk` 工作表，啟動前於 `%APPDATA%\RustDesk\config` 寫入 `RustDesk2.toml` 與 `peers\<id>.toml`，以設定 rendezvous/relay/key 與 peer 密碼。
- AnyDesk：讀取 `anydesk` 工作表，啟動前於 `%APPDATA%\AnyDesk\user.conf` 寫入 `ad.session.viewmode=<ID>:2`，影響視窗顯示模式。
- VNC (TightVNC)：讀取第 3 張工作表（index=2）以產生按鈕清單；啟動時會產生一份可寫的 `vnc.vnc`（保留綁定檔案內容並替換 host/埠/password），並強制 `[options]` 中包含 `viewonly=0`、`shared=1` 以允許滑鼠/遠端控制。

安裝與相依性
-
建議使用虛擬環境 (venv)。安裝相依套件：

```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

`requirements.txt` 目前列出：`pandas`, `openpyxl`, `pycryptodome`。

配置
-
- 將 `Alldesk.xlsx` 放在專案根目錄，三張 sheet 分別為：`rustdesk`、`anydesk`、第 3 張為 VNC。欄位名稱支援中 / 英對應（例如 `Item` / `設備名稱`，`Password` / `密碼`，`Port` / `埠`）。
- 可透過環境變數覆寫可執行檔路徑：`RUSTDESK_APP`、`ANYDESK_APP`。

使用方式
-
1. 啟動

```bash
python Alldesk.py
```

2. 主要操作流程

- 啟動後畫面含三個主要分頁：`RustDesk`、`AnyDesk`、`TightVNC`（VNC）。每個分頁上方有輸入欄位，可手動填入 `連接ID` 與 `密碼` 並按 `連接`。下方則會依 `Alldesk.xlsx` 的資料動態建立多個快速連線按鈕。
- 若想以 Excel 檢視/編輯設定，請按右方 `EXCEL` 分頁（點選後會自動開啟 `Alldesk.xlsx`，並嘗試選取對應的工作表）：
	- 若你在 `RustDesk` 分頁點選 `EXCEL`，會開啟 `rustdesk` 工作表（sheet 1）。
	- 若你在 `AnyDesk` 分頁點選 `EXCEL`，會開啟 `anydesk` 工作表（sheet 2）。
	- 若你在 `TightVNC` 分頁點選 `EXCEL`，會開啟第 3 張工作表（VNC 清單）。

3. 背後行為（簡要）

- RustDesk：啟動前程式會在 `%APPDATA%\RustDesk\config` 下寫入 `RustDesk2.toml` 與 `peers\<id>.toml`，以設定 rendezvous/relay/key 與 peer 密碼，然後以 `rustdesk.exe --connect <id> --password <pwd>` 啟動連線。
- AnyDesk：啟動前程式會在 `%APPDATA%\AnyDesk\user.conf` 寫入 `ad.session.viewmode=<ID>:2`，然後啟動 AnyDesk 並嘗試以命令列或管道傳入密碼。
- TightVNC：啟動時會從專案內的 `vnc.vnc` 範本產生可寫的 `vnc.vnc`，替換 `host` / `port` / `password` 並確保 `[options]` 包含 `viewonly=0`、`shared=1`，接著以 TightVNC 啟動並指定 `-optionsfile=` 參數。

4. 注意事項

- 如果你使用打包後的 `Alldesk.exe`，請確保 `Alldesk.xlsx` 與 `Alldesk.exe` 放在同一資料夾；程式會自動以 exe 所在目錄尋找該檔案。
- 若想要 Excel 在開啟後自動選取特定工作表，程式使用 COM automation（需在環境中包含 `pywin32` 或打包後的 exe 包含 pywin32）。若無法自動切換，請確認系統已安裝 Microsoft Excel 且允許 COM 控制。


偵錯與檢查項目
-
- 若 VNC 連線後無法控制滑鼠：程式會在啟動前在 `vnc.vnc` 的 `[options]` 區段強制寫入 `viewonly=0` 與 `shared=1`；你可以檢查由程式產生的 `vnc.vnc` 檔案（位於腳本資料夾或暫存目錄，視是否打包而定）。
- 若 AnyDesk / RustDesk 未遵循預期的設定行為，請確認目標主機上的 `%APPDATA%` 是否具有寫入權限，以及 `Alldesk.xlsx` 中的資料是否正確。

常見命令
-
- 執行語法檢查：`python -m py_compile Alldesk.py`
- 檢查最近變更：`git log -n 10 --oneline`

貢獻
-
若要新增功能或修正問題，請建立分支並發送 PR。測試 VNC / AnyDesk / RustDesk 的行為需在 Windows 環境並安裝相應客戶端。

作者註記
-
此專案在最近一次修改中：
- 將程式註解中文化並統一欄位命名（`Port` → `埠`）。
- 統一 VNC 欄位佈局並強制寫入 `vnc.vnc` 的 options。README 與 CHANGELOG 已同步更新。
