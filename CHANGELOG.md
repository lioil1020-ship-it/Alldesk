# Changelog

Recent changes (most recent first):

- 2026-01-16 55effbf: UI: 放大並加粗分頁標籤；統一 VNC 佈局並強制 vnc options
- 2026-01-16 cd7de83: 中文化註解、統一 VNC 欄位佈局、移除編輯埠按鈕
- 2026-01-16 22704b5: Auto: update Alldesk.py (rename/init refactor, AnyDesk/RustDesk config writers, translate comments) and update requirements.txt (remove unused deps)

Notes:
- VNC: 已加入寫入 `vnc.vnc` 時的強制 `[options]`，確保 `viewonly=0`、`shared=1` 以允許滑鼠控制。
- UI: Notebook 標籤已放大並改為粗體，使三個分頁標籤更醒目。
- 註解: 多處英文註解已翻成中文。
