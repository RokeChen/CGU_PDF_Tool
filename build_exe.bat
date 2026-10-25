@echo off
chcp 65001 >nul
setlocal
REM =============================================================
REM 在 Windows 上建置 CGU_PDF_Tool.exe
REM OCR 套件 rapidocr_onnxruntime 新版只支援 Python 3.6–3.12，
REM 因此優先使用 Python 3.12，其次 3.11；都沒有才用預設 Python（如 3.14，會安裝較舊的 rapidocr）。
REM 建議安裝 Python 3.12 x64：https://www.python.org/downloads/windows/
REM =============================================================
cd /d "%~dp0"
set "PYEXE="
for %%V in (3.12 3.11) do (
  if not defined PYEXE (
    py -%%V -c "import sys" >nul 2>&1 && set "PYEXE=py -%%V"
  )
)
if not defined PYEXE (
  echo [注意] 找不到 Python 3.12 / 3.11，改用預設 Python。若 OCR 安裝失敗，請安裝 Python 3.12 後再執行。
  set "PYEXE=python"
)
echo 使用：%PYEXE%
%PYEXE% --version
if not exist .venv (
  %PYEXE% -m venv .venv || goto :err
)
call .venv\Scripts\activate.bat || goto :err
python -m pip install --upgrade pip || goto :err
pip install -r requirements.txt pyinstaller || goto :err
pyinstaller --noconfirm --clean --onefile --windowed --name CGU_PDF_Tool --icon assets\app_icon.ico --version-file version_info.txt ^
  --add-data "assets\ocr;assets\ocr" ^
  --hidden-import docx2pdf --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes ^
  --collect-submodules win32com --collect-all pymupdf --collect-all rapidocr_onnxruntime --collect-all pdf2docx --collect-all docx --collect-all pyhanko --collect-all pyhanko_certvalidator ^
  --collect-binaries onnxruntime ^
  pdf_tool.py || goto :err
echo.
echo 建置完成，執行 OCR 自我測試…
dist\CGU_PDF_Tool.exe --selftest-ocr selftest_result.txt
type selftest_result.txt
echo.
echo 完成！執行檔位於 dist\CGU_PDF_Tool.exe
pause
exit /b 0
:err
echo 建置失敗，請檢查上方錯誤訊息。
pause
exit /b 1
