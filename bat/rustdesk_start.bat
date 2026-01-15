@echo off
setlocal

:: 1. 設定參數
set "TARGET_ID=162136755"
set "PASSWORD=Shen+8888"
set "SVR_ADDR=everdura.ddnsfree.com"
set "SVR_KEY=kCC8dq5x8uvEI+fpbIsTpYhCMaMbAxpYmGv6XtR7NsY="
set "CONFIG_PATH=%AppData%\RustDesk\config"
set "PEER_FILE=%CONFIG_PATH%\peers\%TARGET_ID%.toml"
set "EXE_PATH=%~dp0rustdesk.exe"

:: 2. 徹底關閉並清理衝突檔 (解除唯讀並刪除 ThreadId 檔案)
taskkill /f /im rustdesk.exe >nul 2>&1
if exist "%PEER_FILE%" attrib -r "%PEER_FILE%"
del /f /q "%CONFIG_PATH%\peers\*ThreadId*" >nul 2>&1

:: 3. 寫入完整且成功的 162136755.toml 模板
(
    echo password = []
    echo size = [0, 0, 0, 0]
    echo size_ft = [0, 0, 0, 0]
    echo size_pf = [0, 0, 0, 0]
    echo view_style = 'adaptive'
    echo scroll_style = 'scrollauto'
    echo edge_scroll_edge_thickness = 100
    echo image_quality = 'balanced'
    echo custom_image_quality = [50]
    echo show_remote_cursor = false
    echo lock_after_session_end = false
    echo terminal-persistent = false
    echo privacy_mode = false
    echo allow_swap_key = false
    echo port_forwards = []
    echo direct_failures = 1
    echo disable_audio = false
    echo disable_clipboard = false
    echo enable-file-copy-paste = true
    echo show_quality_monitor = false
    echo follow_remote_cursor = false
    echo follow_remote_window = false
    echo keyboard_mode = 'map'
    echo view_only = false
    echo show_my_cursor = false
    echo sync-init-clipboard = false
    echo trackpad-speed = 100
    echo.
    echo [options]
    echo swap-left-right-mouse = ''
    echo codec-preference = 'auto'
    echo collapse_toolbar = ''
    echo zoom-cursor = ''
    echo i444 = ''
    echo custom-fps = '30'
    echo.
    echo [ui_flutter]
    echo wm_RemoteDesktop = '{"width":1300.0,"height":740.0,"offsetWidth":1270.0,"offsetHeight":710.0,"isMaximized":true,"isFullscreen":true}'
    echo.
    echo [info]
    echo username = '104'
    echo hostname = 'checkin-pc'
    echo platform = 'Windows'
) > "%PEER_FILE%"

:: 4. 寫入主設定檔 (解決伺服器識別與預載密碼)
(
    echo rendezvous_server = '%SVR_ADDR%:21116'
    echo nat_type = 1
    echo [options]
    echo custom-rendezvous-server = '%SVR_ADDR%'
    echo relay-server = '%SVR_ADDR%'
    echo key = '%SVR_KEY%'
    echo [peer_settings.%TARGET_ID%]
    echo password = '%PASSWORD%'
) > "%CONFIG_PATH%\RustDesk2.toml"

:: 5. 直接連線 (移除所有 timeout)
start "" "%EXE_PATH%" --connect %TARGET_ID% --password "%PASSWORD%"

exit