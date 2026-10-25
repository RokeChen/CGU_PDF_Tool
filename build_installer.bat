@echo off
chcp 65001 >nul
setlocal
REM 編譯安裝程式：需要先執行 build_exe.bat 產生 dist\CGU_PDF_Tool.exe，並安裝 Inno Setup 6
REM （winget install --id JRSoftware.InnoSetup -e）
cd /d "%~dp0"
if not exist dist\CGU_PDF_Tool.exe (
  echo [錯誤] 找不到 dist\CGU_PDF_Tool.exe，請先執行 build_exe.bat
  goto :err
)
set "ISCC="
for %%P in ("%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" "%ProgramFiles%\Inno Setup 6\ISCC.exe") do (
  if not defined ISCC if exist %%P set "ISCC=%%~P"
)
if not defined ISCC (
  where ISCC >nul 2>&1 && set "ISCC=ISCC"
)
if not defined ISCC (
  echo [錯誤] 找不到 Inno Setup 6（ISCC.exe）。請執行：winget install --id JRSoftware.InnoSetup -e
  goto :err
)
echo 使用：%ISCC%
"%ISCC%" installer.iss || goto :err
echo.
echo 完成！安裝程式位於 dist 資料夾（CGU_PDF_Tool_Setup_版本.exe）
exit /b 0
:err
echo 建置安裝程式失敗。
exit /b 1
