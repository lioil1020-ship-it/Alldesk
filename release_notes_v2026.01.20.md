Alldesk 主要發布說明

本次發布包含：
- Alldesk.exe：單可執行檔
- Alldesk.xlsx：excel資料庫
- Alldesk.zip：免安裝包（請將目錄解壓後執行 Alldesk.exe）。

重要改進與修正：
- 停用 UPX 壓縮以提升啟動與除錯可靠性（已禁用 UPX）。
- 移除所有除錯日誌輸出以降低磁碟寫入與洩漏風險。
- 加入「臨時複製 RustDesk.exe 至暫存」機制：在 onedir 環境下會先複製 RustDesk 可執行檔為隨機名稱，確保每次啟動為新執行個體，避免 RustDesk 的 single-instance 攔截 CLI 參數（如密碼）。
- 新增應用圖示與建立為視窗應用（無終端視窗）。
- 調整資源路徑解析以支援打包後的 _internal 與 exe 路徑。

使用建議與注意事項：
- 若要測試自動帶入密碼的行為，建議使用 onefile（Alldesk.exe）或 onedir（解壓 Alldesk.zip）兩者皆可，不過 onedir 會使用臨時複製以確保一致性。
- 本次發布包含 Alldesk.xlsx（用於定義遠端連線），請妥善保管並避免上傳或公開含密碼的檔案。

聯絡與回報：
- 如遇到打包後行為不一致、UI 自動化失敗或其他問題，請於 repository 提交 Issue 並附上簡要重現步驟與環境（Windows 版本、Alldesk 版本）。

