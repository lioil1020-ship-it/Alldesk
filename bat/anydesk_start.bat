@echo off
setlocal

:: 1. 設定
set "TARGET_ID=1228120787"
set "PASSWORD=Shen+8888"
set "CONF_FILE=%AppData%\AnyDesk\user.conf"

:: 2. 強制關閉現有的 AnyDesk
taskkill /f /im AnyDesk.exe >nul 2>&1
timeout /t 1 /nobreak >nul

:: 3. 初始化設定檔 (直接覆寫，解決亂碼並停用安裝提示)
if not exist "%AppData%\AnyDesk\" mkdir "%AppData%\AppData\AnyDesk"

:: 寫入這四行關鍵參數，能最有效地阻止安裝視窗
(
    echo ad.session.viewmode=%TARGET_ID%:2
) > "%CONF_FILE%"

:: 4. 執行連線 (移除 --silent 以確保視窗出現)
:: 使用 cmd /c 配合管道，可以減少 CMD 視窗殘留
echo %PASSWORD% | "%~dp0AnyDesk.exe" %TARGET_ID% --with-password

exit