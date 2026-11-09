# -*- mode: python ; coding: utf-8 -*-
# macOS 打包設定：產生「CGU PDF Tool.app」（免安裝，可直接執行或拖到「應用程式」資料夾）
# 用法（需在 Mac 上執行）：python -m PyInstaller --noconfirm --clean CGU_PDF_Tool_mac.spec
import re
from PyInstaller.utils.hooks import collect_dynamic_libs, collect_all

# 版本號直接讀 pdf_tool.py 的 APP_VERSION，不需另外維護
with open('pdf_tool.py', encoding='utf-8') as f:
    VERSION = re.search(r'^APP_VERSION = "V([^"]+)"', f.read(), re.M).group(1)

datas = [('assets/ocr', 'assets/ocr')]
binaries = []
hiddenimports = ['docx2pdf']
binaries += collect_dynamic_libs('onnxruntime')
for _pkg in ('pymupdf', 'rapidocr_onnxruntime', 'pdf2docx', 'docx', 'pyhanko', 'pyhanko_certvalidator'):
    tmp_ret = collect_all(_pkg)
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

a = Analysis(
    ['pdf_tool.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['win32com', 'pythoncom', 'pywintypes'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='CGU PDF Tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,          # 依建置的 Mac 而定（GitHub macos-14 為 Apple Silicon arm64）
    codesign_identity=None,    # 無 Apple 開發者憑證：PyInstaller 會自動做 ad-hoc 簽章
    entitlements_file=None,
    icon='assets/app_icon.icns',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='CGU PDF Tool',
)
app = BUNDLE(
    coll,
    name='CGU PDF Tool.app',
    icon='assets/app_icon.icns',
    bundle_identifier='tw.edu.cgu.pdftool',
    version=VERSION,
    info_plist={
        'CFBundleName': 'CGU PDF Tool',
        'CFBundleDisplayName': '長庚大學PDF簡易編輯工具',
        'CFBundleShortVersionString': VERSION,
        'CFBundleVersion': VERSION,
        'NSHumanReadableCopyright': 'Copyright © 2026 CGU_RC（AGPL-3.0）',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '12.0',
        'NSAppleEventsUsageDescription': 'Word 轉 PDF 時需要控制 Microsoft Word。',
        # 讓 Finder 的「打開檔案的應用程式」可以選本程式開 PDF
        'CFBundleDocumentTypes': [{
            'CFBundleTypeName': 'PDF Document',
            'CFBundleTypeRole': 'Editor',
            'LSHandlerRank': 'Alternate',
            'LSItemContentTypes': ['com.adobe.pdf'],
        }],
    },
)
