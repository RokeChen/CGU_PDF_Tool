# -*- coding: utf-8 -*-
"""
長庚大學PDF簡易編輯工具（PyQt6 版）── 本機端 PDF 編輯 / 合併 / 組織 / 轉換 / 浮水印
=================================================================

執行前需安裝的第三方套件（請在命令提示字元 / PowerShell 執行）：

    pip install PyQt6 PyMuPDF pypdf img2pdf Pillow docx2pdf pywin32 rapidocr_onnxruntime

  * rapidocr_onnxruntime：離線 OCR（ONNX Runtime）；繁中模型放在 assets/ocr/
    ★ rapidocr_onnxruntime 新版僅支援 Python 3.6–3.12，建議用 Python 3.12 建置

  * PyQt6     ：圖形介面
  * PyMuPDF   ：PDF 繪製、縮圖、編輯、浮水印（import 名稱為 pymupdf / fitz）
  * pypdf     ：合併 PDF
  * img2pdf   ：圖片無損轉 PDF
  * Pillow    ：圖片前處理（處理 PNG 透明背景 RGBA 等）
  * Word 轉 PDF：優先使用 Microsoft Word，沒有 Word 時自動改用 LibreOffice（免費）
  * pywin32   ：docx2pdf 在 Windows 透過 COM 呼叫 Word 時需要（通常會自動安裝）

執行方式：  python pdf_tool.py

功能分頁：
  1. 格式轉換：拖曳 .docx / .jpg / .jpeg / .png 轉成 PDF
  2. 合併檔案：拖曳多個 PDF，調整順序後合併
  3. 組織頁面：縮圖網格、多選刪除、旋轉、拖曳排序後匯出
  4. 文字浮水印：半透明文字壓印在每頁中央（支援中文）
  5. 編輯 PDF：檢視/縮放/換頁，修改原有文字、新增文字/圖片/標記/註解、塗白、復原、存檔
"""

import os
import sys
import io
import traceback
import base64

# PyInstaller 以 console=False（無主控台）打包時 sys.stdout / sys.stderr 為 None，
# 任何第三方套件（如 docx2pdf 內的 tqdm 進度列）寫入 stdout 都會出現
# "'NoneType' object has no attribute 'write'"。這裡統一導向 devnull。
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

# PyMuPDF 新版建議 import pymupdf；舊版只有 fitz，兩者 API 相同
try:
    import pymupdf as fitz
except ImportError:  # pragma: no cover
    import fitz

from pypdf import PdfReader, PdfWriter
from PIL import Image
import img2pdf

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QPoint, QRect, QRectF, QSize, QPointF, QSettings, QByteArray, QObject, QUrl
from PyQt6.QtGui import (QPixmap, QImage, QPainter, QPen, QColor, QIcon,
                         QTransform, QAction, QActionGroup, QKeySequence,
                         QPainterPath, QPolygonF, QFont, QFontDatabase, QDesktopServices, QPageLayout, QShortcut)
from PyQt6.QtSvg import QSvgRenderer
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog, QAbstractPrintDialog
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QListWidgetItem, QLabel, QFileDialog, QMessageBox,
    QProgressBar, QLineEdit, QCheckBox, QSpinBox, QDoubleSpinBox, QSlider,
    QColorDialog, QFormLayout, QGroupBox, QScrollArea, QToolBar, QInputDialog,
    QAbstractItemView, QListView, QDialog, QDialogButtonBox, QPlainTextEdit,
    QSizePolicy, QStatusBar, QComboBox, QFrame, QStackedWidget, QToolButton,
    QButtonGroup, QGraphicsDropShadowEffect, QRadioButton, QMenu, QProgressDialog)


# ---- 版本與著作權資訊 ----
APP_VERSION = "V2.4.1"
APP_AUTHOR = "CGU_RC"
APP_YEAR = "2026"

APP_LICENSE = "AGPL-3.0"
APP_SOURCE_URL = "https://github.com/RokeChen/CGU_PDF_Tool"   # 原始碼網址，「關於」視窗會顯示連結
import datetime

# 長庚大學校徽（著作權屬長庚大學所有；透明背景 PNG，以 base64 內嵌，exe 不需外部檔案）
APP_LOGO_B64 = (
    "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"
)


# =====================================================================
# 共用工具函式
# =====================================================================

IMAGE_EXTS = (".jpg", ".jpeg", ".png")
WORD_EXTS = (".docx",)
PDF_EXTS = (".pdf",)


def find_cjk_font():
    """
    尋找系統上可用的中文（CJK）字型檔，用於浮水印與編輯時插入中文。
    依序嘗試 Windows 常見繁中字型、macOS、Linux Noto 字型。
    找不到時回傳 None，呼叫端會改用 PyMuPDF 內建的 'china-t'（繁體中文）字型。
    """
    candidates = [
        r"C:\Windows\Fonts\msjh.ttc",      # 微軟正黑體
        r"C:\Windows\Fonts\msjh.ttf",
        r"C:\Windows\Fonts\mingliu.ttc",   # 細明體
        r"C:\Windows\Fonts\kaiu.ttf",      # 標楷體
        "/System/Library/Fonts/PingFang.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def font_kwargs(fontfile=None):
    """
    產生 page.insert_text() 用的字型參數。
    有系統字型檔就嵌入該字型；否則用內建 CJK 字型 'china-t'。
    """
    fontfile = fontfile or find_cjk_font()
    if fontfile:
        return {"fontname": "cjkfont", "fontfile": fontfile}
    return {"fontname": "china-t"}


def is_ascii(text):
    """判斷文字是否全為 ASCII（可用標準 Helvetica 字型，保留較接近原樣）"""
    try:
        text.encode("ascii")
        return True
    except UnicodeEncodeError:
        return False


def unique_path(path):
    """若檔案已存在，自動在檔名後加上 (1)、(2)… 避免覆蓋"""
    base, ext = os.path.splitext(path)
    i, out = 1, path
    while os.path.exists(out):
        out = f"{base}({i}){ext}"
        i += 1
    return out


def fitz_pixmap_to_qpixmap(pix):
    """把 PyMuPDF 的 Pixmap 轉成 Qt 的 QPixmap 以便顯示"""
    fmt = QImage.Format.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
    img = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
    return QPixmap.fromImage(img.copy())  # copy() 確保資料脫離 pix 的記憶體


def thumb_icon(pix):
    """縮圖圖示：選取狀態也使用原圖，避免 Qt 預設把縮圖染成藍色"""
    icon = QIcon(pix)
    icon.addPixmap(pix, QIcon.Mode.Selected)
    return icon


def color_to_rgb01(qcolor):
    """QColor → PyMuPDF 使用的 (r, g, b) 0~1 浮點數"""
    return (qcolor.redF(), qcolor.greenF(), qcolor.blueF())


def int_color_to_rgb01(c):
    """get_text('dict') 傳回的整數 sRGB 色碼 → (r, g, b) 0~1"""
    return (((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255, (c & 255) / 255)


# =====================================================================
# 核心邏輯（與介面無關，可單獨測試）。progress 為回呼函式 progress(0~100)
# =====================================================================

def image_to_pdf_bytes_list(paths):
    """
    把圖片讀入並處理成 img2pdf 可接受的位元組。
    img2pdf 不接受含透明度（alpha）的圖片，因此 RGBA / LA / P(含透明) 的圖片
    會先用 Pillow 疊到白色背景上轉成 RGB，再存成 PNG 位元組。
    """
    result = []
    for p in paths:
        with Image.open(p) as im:
            has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
            if has_alpha or im.mode not in ("RGB", "L", "CMYK"):
                im = im.convert("RGBA")
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[-1])  # 以 alpha 通道當遮罩
                buf = io.BytesIO()
                bg.save(buf, format="PNG")
                result.append(buf.getvalue())
            else:
                with open(p, "rb") as f:   # 不需處理的圖直接用原檔（無損）
                    result.append(f.read())
    return result


def convert_images_to_pdf(paths, out_path):
    """把一張或多張圖片合成一個 PDF 檔（每張圖一頁）"""
    data = image_to_pdf_bytes_list(paths)
    with open(out_path, "wb") as f:
        f.write(img2pdf.convert(data))
    return out_path


def _docx_to_pdf_win32(docx_path, out_path):
    """直接以 COM 呼叫 Word 匯出 PDF（不經 docx2pdf，避免 stdout/tqdm 問題）。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError("缺少 pywin32，請執行：pip install pywin32")
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as e:
            raise RuntimeError(f"無法啟動 Microsoft Word（請確認已安裝並啟用授權）：{e}")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3      # msoAutomationSecurityForceDisable：不執行巨集
        except Exception:
            pass
        doc = word.Documents.Open(FileName=os.path.abspath(docx_path), ConfirmConversions=False,
                                  ReadOnly=True, AddToRecentFiles=False, Visible=False,
                                  OpenAndRepair=True, NoEncodingDialog=True)
        # 17 = wdExportFormatPDF
        doc.ExportAsFixedFormat(OutputFileName=os.path.abspath(out_path), ExportFormat=17,
                                OpenAfterExport=False, OptimizeFor=0, CreateBookmarks=1)
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"Word 轉檔失敗（請確認檔案未被 Word 開啟、未加密，且路徑可存取）：{e}")
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


LIBREOFFICE_URL = "https://zh-tw.libreoffice.org/download/download/"


class WordUnavailable(RuntimeError):
    """本機沒有可用的 Microsoft Word（未安裝或無法啟動）。"""


def find_soffice():
    """尋找 LibreOffice 的 soffice 執行檔；找不到回傳 None。"""
    import shutil
    cands = []
    if sys.platform == "win32":
        try:
            import winreg
            for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                for sub in (r"SOFTWARE\LibreOffice\UNO\InstallPath",
                            r"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath"):
                    try:
                        with winreg.OpenKey(hive, sub) as k:
                            d = winreg.QueryValueEx(k, "")[0]
                            cands.append(os.path.join(d, "soffice.exe"))
                    except OSError:
                        pass
        except ImportError:
            pass
        for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
            base = os.environ.get(env)
            if base:
                cands.append(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    elif sys.platform == "darwin":
        cands.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    for name in ("soffice", "libreoffice"):
        w = shutil.which(name)
        if w:
            cands.append(w)
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def _docx_to_pdf_libreoffice(docx_path, out_path, soffice=None):
    """以 LibreOffice（無畫面模式）把 Word 轉為 PDF。"""
    import subprocess
    import tempfile
    import shutil
    import pathlib
    soffice = soffice or find_soffice()
    if not soffice:
        raise RuntimeError("找不到 LibreOffice。")
    with tempfile.TemporaryDirectory(prefix="cgupdf_lo_") as tmp:
        # 使用獨立的設定檔目錄，避免使用者已開著 LibreOffice 時轉檔沒反應
        profile = pathlib.Path(tmp, "profile").as_uri()
        outdir = os.path.join(tmp, "out")
        os.makedirs(outdir)
        cmd = [soffice, f"-env:UserInstallation={profile}", "--headless", "--norestore",
               "--nolockcheck", "--convert-to", "pdf", "--outdir", outdir, os.path.abspath(docx_path)]
        kw = {}
        if sys.platform == "win32":
            kw["creationflags"] = 0x08000000      # CREATE_NO_WINDOW：不跳出黑色視窗
        try:
            r = subprocess.run(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE, timeout=300, **kw)
        except subprocess.TimeoutExpired:
            raise RuntimeError("LibreOffice 轉檔逾時（超過 5 分鐘）。")
        except Exception as e:
            raise RuntimeError(f"無法執行 LibreOffice：{e}")
        produced = os.path.join(outdir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        if not os.path.exists(produced):
            pdfs = [f for f in os.listdir(outdir) if f.lower().endswith(".pdf")]
            if not pdfs:
                msg = (r.stderr or r.stdout or b"").decode("utf-8", "replace").strip()
                raise RuntimeError("LibreOffice 轉檔失敗" + (f"：{msg[:300]}" if msg else "。"))
            produced = os.path.join(outdir, pdfs[0])
        if os.path.exists(out_path):
            os.remove(out_path)
        shutil.move(produced, out_path)


def convert_docx_to_pdf(docx_path, out_path):
    """
    把 Word 靜默轉為 PDF，依序嘗試：
      1. Microsoft Word（Windows：COM；macOS：docx2pdf）
      2. LibreOffice（免費，沒有 Word 時的備援）
    兩者都沒有時丟出 RuntimeError，訊息為使用者看得懂的中文。
    """
    import zipfile
    if not zipfile.is_zipfile(docx_path):
        raise RuntimeError("不是有效的 Word（.docx）檔，檔案可能已損毀，或其實是舊版 .doc 改了副檔名。")
    word_err = None
    try:
        if sys.platform == "win32":
            if not word_available():
                raise WordUnavailable("未安裝 Microsoft Word")
            _docx_to_pdf_win32(docx_path, out_path)
        elif sys.platform == "darwin" and os.path.isdir("/Applications/Microsoft Word.app"):
            try:
                from docx2pdf import convert as docx_convert
            except ImportError:
                raise WordUnavailable("未安裝 docx2pdf")
            try:
                docx_convert(docx_path, out_path)
            except Exception as e:
                raise RuntimeError(f"Word 轉檔失敗（請確認檔案未被開啟）：{e}")
        else:
            raise WordUnavailable("未安裝 Microsoft Word")
        if os.path.exists(out_path):
            return out_path
        word_err = RuntimeError("Word 轉檔沒有產生輸出檔。")
    except WordUnavailable as e:
        word_err = e
    except RuntimeError as e:
        word_err = e

    soffice = find_soffice()
    if soffice:
        try:
            _docx_to_pdf_libreoffice(docx_path, out_path, soffice)
        except RuntimeError as e:
            if isinstance(word_err, WordUnavailable):
                raise
            raise RuntimeError(f"{word_err}\n改用 LibreOffice 也失敗：{e}")
        if os.path.exists(out_path):
            return out_path
        raise RuntimeError("LibreOffice 轉檔沒有產生輸出檔。")

    if isinstance(word_err, WordUnavailable):
        raise RuntimeError("這台電腦沒有 Microsoft Word，也沒有 LibreOffice，無法轉換 Word 檔。\n"
                           f"請安裝免費的 LibreOffice 後再試：{LIBREOFFICE_URL}")
    raise word_err


def run_conversion(files, out_dir, combine_images, progress=None):
    """
    格式轉換主流程。
    files：檔案清單；out_dir：輸出資料夾（None＝與各原檔相同資料夾；合併圖片時存在第一張圖片的資料夾）；
    combine_images：是否把所有圖片合成一個 PDF。
    回傳 (成功輸出的檔案清單, 錯誤訊息清單)。單一檔案失敗不會中斷其他檔案。
    """
    outputs, errors = [], []
    images = [f for f in files if f.lower().endswith(IMAGE_EXTS)]
    words = [f for f in files if f.lower().endswith(WORD_EXTS)]
    jobs = []
    if combine_images and images:
        jobs.append(("images", images))
    else:
        jobs += [("image", [p]) for p in images]
    jobs += [("word", [p]) for p in words]
    total = max(len(jobs), 1)
    for i, (kind, paths) in enumerate(jobs):
        try:
            folder = out_dir or os.path.dirname(os.path.abspath(paths[0]))
            if kind == "images":
                out = unique_path(os.path.join(folder, "合併圖片.pdf"))
                convert_images_to_pdf(paths, out)
            elif kind == "image":
                name = os.path.splitext(os.path.basename(paths[0]))[0] + ".pdf"
                out = unique_path(os.path.join(folder, name))
                convert_images_to_pdf(paths, out)
            else:
                name = os.path.splitext(os.path.basename(paths[0]))[0] + ".pdf"
                out = unique_path(os.path.join(folder, name))
                convert_docx_to_pdf(paths[0], out)
            outputs.append(out)
        except Exception as e:
            errors.append(f"{os.path.basename(paths[0])}：{e}")
        if progress:
            progress(int((i + 1) * 100 / total))
    return outputs, errors


# =====================================================================
# PDF → Word（.docx）
#   主要：Microsoft Word（COM 自動化，Word 內建的 PDF 重排轉換）
#   備援：pdf2docx（MIT 授權，純 Python，無 Word 也能用）
#   掃描檔：先以 RapidOCR 辨識，重建為「文字在原位置」的 PDF 再轉換，讓 Word 內容可編輯
# =====================================================================

def pdf_needs_ocr(path):
    """回傳沒有文字層（疑似掃描）的頁碼清單"""
    with fitz.open(path) as d:
        return [i for i, pg in enumerate(d) if not pg.get_text("text").strip()]


def build_ocr_text_pdf(src_path, out_path, pages, progress=None):
    """
    把掃描頁改成「可見文字放在原位置」的頁面（不含原影像），其他頁原樣保留。
    這樣 Word / pdf2docx 會產生可編輯的文字（版面可能與原檔不同）。回傳辨識行數。
    """
    src = fitz.open(src_path)
    out = fitz.open()
    fpath = ocr_rebuild_font()
    f = fitz.Font(fontfile=fpath) if fpath else fitz.Font("cjk")
    n = 0
    for k, i in enumerate(range(len(src))):
        if i in pages:
            pg = src[i]
            res = ocr_page_results(pg)
            np_ = out.new_page(width=pg.rect.width, height=pg.rect.height)
            if fpath:
                np_.insert_font(fontname="OCRV", fontfile=fpath)
            else:
                np_.insert_font(fontname="OCRV", fontbuffer=f.buffer)
            for r, text, _ in sorted(res, key=lambda t: (round(t[0].y0 / 4), t[0].x0)):
                size = max(4, r.height * 0.8)
                w = f.text_length(text, fontsize=size)
                if w > r.width * 1.05 and w > 0:
                    size = max(4, size * r.width / w)
                np_.insert_text(fitz.Point(r.x0, r.y1 - r.height * 0.2), text, fontsize=size, fontname="OCRV")
                n += 1
        else:
            out.insert_pdf(src, from_page=i, to_page=i)
        if progress:
            progress(k + 1, len(src))
    out.save(out_path, garbage=3, deflate=True)
    out.close()
    src.close()
    return n


def ocr_rebuild_font():
    """
    掃描頁重建用的中文字型：Word 讀取內建 Droid Sans Fallback 時會漏字，
    因此優先使用系統的微軟正黑體 / 標楷體（實測 Word 可完整轉出），找不到才用內建字型。
    """
    for name in ("msjh.ttc", "kaiu.ttf", "msjhbd.ttc"):
        for d in (os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
                  os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")):
            p = os.path.join(d, name)
            if os.path.exists(p):
                return p
    for p in ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",):
        if os.path.exists(p):
            return p
    return None


def _norm_chars(t):
    return [c for c in t if not c.isspace()]


def docx_text(path):
    """讀出 .docx 的全部文字（段落 + 表格 + 文字方塊）"""
    import zipfile
    import re
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "replace")
    return "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", xml))


def text_coverage(pdf_path, docx_path):
    """docx 保留了 PDF 文字的比例（0–1），用來偵測 Word 轉換漏字"""
    with fitz.open(pdf_path) as d:
        src = _norm_chars("".join(pg.get_text("text") for pg in d))
    if not src:
        return 1.0
    got = docx_text(docx_path)
    from collections import Counter
    a, b = Counter(src), Counter(_norm_chars(got))
    return sum(min(v, b[k]) for k, v in a.items()) / len(src)


def word_available():
    """是否可用 Microsoft Word（COM）"""
    if sys.platform != "win32":
        return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"))
        return True
    except Exception:
        return False


def pdf_to_docx_word(pdf_path, out_path):
    """
    以 Microsoft Word 開啟 PDF（Word 會自動轉換重排）→ 另存 .docx。
    暫時設定 DisableConvertPdfWarning=1 避免「Word 會將 PDF 轉換為可編輯文件」提示，完成後還原。
    """
    import pythoncom
    import win32com.client
    import winreg
    pythoncom.CoInitialize()
    word = None
    key = old = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            ver = str(word.Version)
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"Software\Microsoft\Office\{ver}\Word\Options")
            try:
                old = winreg.QueryValueEx(key, "DisableConvertPdfWarning")[0]
            except OSError:
                old = None
            winreg.SetValueEx(key, "DisableConvertPdfWarning", 0, winreg.REG_DWORD, 1)
        except Exception:
            key = None
        doc = word.Documents.Open(FileName=os.path.abspath(pdf_path), ConfirmConversions=False, ReadOnly=True,
                                  AddToRecentFiles=False, Visible=False)
        try:
            doc.SaveAs2(FileName=os.path.abspath(out_path), FileFormat=16)   # 16 = wdFormatDocumentDefault (.docx)
        finally:
            doc.Close(False)
    except Exception as e:
        raise RuntimeError(f"Word 轉換失敗：{e}")
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        if key is not None:
            try:
                if old is None:
                    winreg.DeleteValue(key, "DisableConvertPdfWarning")
                else:
                    winreg.SetValueEx(key, "DisableConvertPdfWarning", 0, winreg.REG_DWORD, old)
                winreg.CloseKey(key)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def pdf_to_docx_pdf2docx(pdf_path, out_path):
    """備援引擎：pdf2docx（MIT），不需 Word"""
    try:
        from pdf2docx import Converter
    except ImportError:
        raise RuntimeError("找不到 pdf2docx，請執行：pip install pdf2docx")
    import logging
    logging.getLogger("pdf2docx").setLevel(logging.ERROR)
    cv = Converter(pdf_path)
    try:
        cv.convert(out_path, multi_processing=False)
    finally:
        cv.close()


def run_pdf_to_word(files, out_dir=None, ocr_scanned=True, engine="auto", progress=None):
    """
    批次 PDF → Word。out_dir=None 表示輸出到各 PDF 所在資料夾。
    engine：auto（有 Word 用 Word，否則 pdf2docx）/ word / pdf2docx
    回傳 dict：outputs [(docx, 引擎, 是否 OCR)], errors [訊息]
    """
    import tempfile
    outputs, errors = [], []
    use_word = engine == "word" or (engine == "auto" and word_available())
    total = max(len(files), 1)
    for i, f in enumerate(files):
        tmp = None
        try:
            folder = out_dir or os.path.dirname(os.path.abspath(f))
            out = unique_path(os.path.join(folder, os.path.splitext(os.path.basename(f))[0] + ".docx"))
            src, did_ocr = f, False
            if ocr_scanned:
                pages = pdf_needs_ocr(f)
                if pages:
                    fd, tmp = tempfile.mkstemp(suffix=".pdf", prefix="cgu_ocr_")
                    os.close(fd)
                    build_ocr_text_pdf(f, tmp, pages, progress=lambda a, b: progress and progress(
                        int((i + 0.6 * a / b) * 100 / total)))
                    src, did_ocr = tmp, True
            used = "Microsoft Word" if use_word else "pdf2docx"
            try:
                if use_word:
                    pdf_to_docx_word(src, out)
                    cov = text_coverage(src, out)
                    if cov < 0.9 and engine == "auto":      # Word 漏字（常見於特殊內嵌字型）→ 改用 pdf2docx
                        os.remove(out)
                        pdf_to_docx_pdf2docx(src, out)
                        used = f"pdf2docx（Word 只轉出 {cov:.0%} 文字，已改用 pdf2docx）"
                else:
                    pdf_to_docx_pdf2docx(src, out)
            except Exception as e:
                if use_word and engine == "auto":      # Word 失敗時自動改用 pdf2docx
                    pdf_to_docx_pdf2docx(src, out)
                    used = f"pdf2docx（Word 失敗：{e}）"
                else:
                    raise
            if not os.path.exists(out) or os.path.getsize(out) == 0:
                raise RuntimeError("沒有產生輸出檔")
            outputs.append((out, used, did_ocr))
        except Exception as e:
            errors.append(f"{os.path.basename(f)}：{e}")
        finally:
            if tmp and os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
        if progress:
            progress(int((i + 1) * 100 / total))
    return {"outputs": outputs, "errors": errors}


def merge_pdfs(paths, out_path, progress=None):
    """用 pypdf 依清單順序合併多個 PDF"""
    if len(paths) < 1:
        raise ValueError("沒有要合併的檔案")
    writer = PdfWriter()
    for i, p in enumerate(paths):
        reader = PdfReader(p)
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                raise RuntimeError(f"{os.path.basename(p)} 有密碼保護，無法合併")
        for page in reader.pages:
            writer.add_page(page)
        if progress:
            progress(int((i + 1) * 90 / len(paths)))
    with open(out_path, "wb") as f:
        writer.write(f)
    if progress:
        progress(100)
    return out_path


def open_as_pdf(path):
    """開啟 PDF；圖片（JPG / PNG）轉成單頁 PDF 後開啟。有密碼的 PDF 丟出 RuntimeError"""
    if path.lower().endswith(IMAGE_EXTS):
        with fitz.open(path) as img:
            return fitz.open("pdf", img.convert_to_pdf())
    doc = fitz.open(path)
    if doc.needs_pass:
        doc.close()
        raise RuntimeError(f"{os.path.basename(path)} 有密碼保護，請先解除密碼")
    return doc


def export_organized(src_path, page_specs, out_path, progress=None):
    """
    依新的頁面排序與旋轉匯出 PDF（可混合多個來源檔）。
    page_specs 每一項可以是：
      (原始頁碼, 旋轉角度)             → 取自 src_path
      (來源檔路徑, 原始頁碼, 旋轉角度) → 取自指定的 PDF / 圖片
      (None, (寬, 高), 旋轉角度)       → 空白頁
    順序即輸出順序；清單中沒有的頁面視為已刪除。
    """
    if not page_specs:
        raise ValueError("所有頁面都被刪除了，無法匯出")
    docs = {}
    out = fitz.open()
    n = len(page_specs)
    try:
        for i, spec in enumerate(page_specs):
            src, idx, rot = (src_path, spec[0], spec[1]) if len(spec) == 2 else spec
            if src is None:                                   # 空白頁
                w, h = idx
                page = out.new_page(width=w, height=h)
            else:
                if src not in docs:
                    docs[src] = open_as_pdf(src)
                out.insert_pdf(docs[src], from_page=idx, to_page=idx)
                page = out[-1]
            page.set_rotation((page.rotation + rot) % 360)
            if progress:
                progress(int((i + 1) * 95 / n))
        out.save(out_path, garbage=3, deflate=True)
    finally:
        out.close()
        for d in docs.values():
            d.close()
    if progress:
        progress(100)
    return out_path


def add_text_watermark(src_path, out_path, text, opacity=0.3, fontsize=60,
                       angle=45, color=(1, 0, 0), fontfile=None, progress=None):
    """
    在每一頁中央壓印半透明文字浮水印。
    * 中文：使用系統 CJK 字型檔（或內建 china-t）
    * 任意角度：用 morph（以頁面中心為軸的旋轉矩陣）
    * 半透明：fill_opacity
    """
    if not text.strip():
        raise ValueError("請輸入浮水印文字")
    doc = fitz.open(src_path)
    kw = font_kwargs(fontfile)
    font = fitz.Font(fontfile=kw["fontfile"]) if "fontfile" in kw else fitz.Font("china-t")
    n = len(doc)
    for i, page in enumerate(doc):
        # 使用未旋轉座標系的頁面中心；若頁面本身有旋轉，補償角度使畫面上看起來一致
        rect = page.cropbox
        center = fitz.Point((rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2)
        width = font.text_length(text, fontsize=fontsize)
        # 文字基線起點：讓文字水平置中、垂直大致置中
        start = fitz.Point(center.x - width / 2, center.y + fontsize * 0.35)
        mat = fitz.Matrix(angle - page.rotation)
        page.insert_text(start, text, fontsize=fontsize, color=color,
                         fill_opacity=opacity, morph=(center, mat), overlay=True, **kw)
        if progress:
            progress(int((i + 1) * 95 / n))
    doc.save(out_path, garbage=3, deflate=True)
    doc.close()
    if progress:
        progress(100)
    return out_path


# ---------------- 編輯功能的核心函式（操作已開啟的 fitz.Document） ----------------

# ---------------- 數位簽章（PKCS#12 / .pfx 憑證，使用 pyHanko）----------------
def pdf_signature_count(doc):
    """文件中已有的數位簽章欄位數"""
    n = 0
    try:
        for page in doc:
            for w_ in page.widgets() or []:
                if w_.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE:
                    n += 1
    except Exception:
        pass
    return n


def load_pfx(pfx_path, password):
    """讀取 .pfx / .p12 憑證，回傳 (signer, 憑證資訊 dict)；密碼錯誤或格式不符時丟出 RuntimeError"""
    try:
        from pyhanko.sign import signers
    except ImportError:
        raise RuntimeError("缺少數位簽章元件 pyhanko，請執行：pip install pyhanko")
    import logging
    logging.getLogger("pyhanko").setLevel(logging.CRITICAL)
    signer = signers.SimpleSigner.load_pkcs12(pfx_file=pfx_path, passphrase=(password or "").encode("utf-8"))
    if signer is None:
        raise RuntimeError("無法讀取憑證：密碼錯誤，或檔案不是有效的 .pfx / .p12 憑證。")
    cert = signer.signing_cert
    subj = cert.subject.native
    info = {
        "name": subj.get("common_name") or subj.get("organization_name") or "（未命名）",
        "issuer": cert.issuer.native.get("common_name") or cert.issuer.native.get("organization_name") or "",
        "not_after": cert["tbs_certificate"]["validity"]["not_after"].native,
        "not_before": cert["tbs_certificate"]["validity"]["not_before"].native,
    }
    return signer, info


def sign_pdf_bytes(data, signer, reason="", location="", field_name=None):
    """以增量更新方式簽署 PDF（不改動原內容，保留既有簽章），回傳簽署後的位元組"""
    import io as _io
    from pyhanko.sign import signers
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    w = IncrementalPdfFileWriter(_io.BytesIO(data), strict=False)
    meta = signers.PdfSignatureMetadata(field_name=field_name or "Signature1", reason=reason or None,
                                        location=location or None, md_algorithm="sha256")
    out = signers.sign_pdf(w, meta, signer=signer)
    return out.getvalue()


class DigitalSignDialog(QDialog):
    """數位簽章設定：選憑證、輸入密碼、簽章原因 / 地點、輸出檔名"""

    def __init__(self, parent, out_default):
        super().__init__(parent)
        self.setWindowTitle("數位簽章（.pfx 憑證）")
        self.setModal(True)
        self.setMinimumWidth(520)
        lay = QVBoxLayout(self)
        hint = QLabel("以憑證檔簽署後，Acrobat 等閱讀器會顯示簽署人與「簽章有效」；"
                      "簽署後若文件被修改，簽章會顯示無效。原檔不會被更動。")
        hint.setWordWrap(True)
        hint.setObjectName("PanelHint")
        lay.addWidget(hint)
        form = QFormLayout()
        st = settings()
        prow = QHBoxLayout()
        self.ed_pfx = QLineEdit(st.value("dsign/pfx", ""))
        self.ed_pfx.setPlaceholderText("選擇 .pfx 或 .p12 憑證檔")
        b = QPushButton("瀏覽…")
        b.clicked.connect(self._pick_pfx)
        prow.addWidget(self.ed_pfx, 1)
        prow.addWidget(b)
        form.addRow("憑證檔", prow)
        self.ed_pw = QLineEdit()
        self.ed_pw.setEchoMode(QLineEdit.EchoMode.Password)
        self.ed_pw.setPlaceholderText("憑證密碼（不會被儲存）")
        form.addRow("密碼", self.ed_pw)
        self.cmb_reason = QComboBox()
        self.cmb_reason.setEditable(True)
        self.cmb_reason.addItems(["核准", "同意", "已審閱", "確認內容無誤"])
        self.cmb_reason.setCurrentText(st.value("dsign/reason", "核准"))
        form.addRow("簽章原因", self.cmb_reason)
        self.ed_loc = QLineEdit(st.value("dsign/location", "長庚大學"))
        form.addRow("簽署地點", self.ed_loc)
        orow = QHBoxLayout()
        self.ed_out = QLineEdit(out_default)
        b2 = QPushButton("瀏覽…")
        b2.clicked.connect(self._pick_out)
        orow.addWidget(self.ed_out, 1)
        orow.addWidget(b2)
        form.addRow("另存為", orow)
        lay.addLayout(form)
        self.lbl_cert = QLabel("")
        self.lbl_cert.setWordWrap(True)
        lay.addWidget(self.lbl_cert)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("簽署")
        bb.button(QDialogButtonBox.StandardButton.Ok).setProperty("primary", True)
        bb.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        bb.accepted.connect(self._accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.signer = self.cert_info = None
        (self.ed_pw if self.ed_pfx.text() else self.ed_pfx).setFocus()

    def _pick_pfx(self):
        p, _ = QFileDialog.getOpenFileName(self, "選擇憑證檔", os.path.dirname(self.ed_pfx.text()),
                                           "憑證檔 (*.pfx *.p12)")
        if p:
            self.ed_pfx.setText(p)
            self.ed_pw.setFocus()

    def _pick_out(self):
        p, _ = QFileDialog.getSaveFileName(self, "簽署後另存為", self.ed_out.text(), "PDF (*.pdf)")
        if p:
            self.ed_out.setText(p)

    def _accept(self):
        pfx = self.ed_pfx.text().strip()
        if not os.path.isfile(pfx):
            return QMessageBox.warning(self, "數位簽章", "請選擇憑證檔（.pfx / .p12）。")
        if not self.ed_out.text().strip():
            return QMessageBox.warning(self, "數位簽章", "請指定輸出檔名。")
        try:
            self.signer, self.cert_info = load_pfx(pfx, self.ed_pw.text())
        except Exception as e:
            return QMessageBox.warning(self, "數位簽章", str(e))
        now = datetime.datetime.now(datetime.timezone.utc)
        na = self.cert_info["not_after"]
        if na and na < now:
            return QMessageBox.warning(self, "數位簽章",
                                       f"此憑證已於 {na.astimezone().strftime('%Y/%m/%d')} 到期，無法用來簽署。")
        st = settings()
        st.setValue("dsign/pfx", pfx)
        st.setValue("dsign/reason", self.cmb_reason.currentText())
        st.setValue("dsign/location", self.ed_loc.text())
        self.accept()


# ---------------- 打勾（核取方塊）----------------
CHECK_TAG = "CGU_CHECK"          # 打勾註解的識別標記，再點一次可取消
BOX_GLYPHS = set("□☐▢◻⬜❏❐❑❒▫")


def find_boxes(page, area=None):
    """
    找出頁面上的框框（未旋轉頁面座標），回傳 [Rect]：
      * 文字字元 □ ☐ 等（Word / 報表轉出的 PDF 常見）
      * 向量繪製的小正方形（表單、PDF 產生器畫的框）
    area 給定時只回傳與該範圍相交的框框。
    """
    boxes = []
    try:
        raw = page.get_text("rawdict")
        for b in raw["blocks"]:
            for ln in b.get("lines", []):
                for sp in ln["spans"]:
                    for ch in sp.get("chars", []):
                        if ch["c"] in BOX_GLYPHS:
                            r = fitz.Rect(ch["bbox"])
                            side = min(r.width, r.height)
                            c = (r.tl + r.br) / 2
                            # 字形外框通常比看到的方框大一些，縮成置中的正方形
                            side *= 0.82
                            boxes.append(fitz.Rect(c.x - side / 2, c.y - side / 2, c.x + side / 2, c.y + side / 2))
    except Exception:
        pass
    try:
        for d in page.get_drawings():
            r = fitz.Rect(d["rect"])
            if 4 <= r.width <= 30 and 4 <= r.height <= 30 and 0.75 <= r.width / max(r.height, 0.1) <= 1.33 \
                    and d.get("color") is not None:
                if not any(abs(r.x0 - x.x0) < 1.5 and abs(r.y0 - x.y0) < 1.5 for x in boxes):
                    boxes.append(r)
    except Exception:
        pass
    if area is not None:
        boxes = [b for b in boxes if b.intersects(area)]
    return boxes


def box_at(page, point, tol=3):
    """點選位置附近的框框（容許些微誤差）"""
    best = None
    for b in find_boxes(page):
        if (b + (-tol, -tol, tol, tol)).contains(point):
            if best is None or b.width < best.width:
                best = b
    return best


def check_annots(page, rect):
    """與指定框框位置重疊的打勾註解"""
    c = (rect.tl + rect.br) / 2
    return [a for a in (page.annots() or []) if a.info.get("title") == CHECK_TAG
            and (a.rect + (-2, -2, 2, 2)).contains(c)]


def add_check_mark(page, box, symbol="check", color=(0, 0, 0)):
    """在框框（未旋轉座標）畫上 ✓ / ✗ / ●；旋轉頁面也保持正向"""
    vis = (fitz.Rect(box) * page.rotation_matrix).normalize()
    x0, y0, w, h = vis.x0, vis.y0, vis.width, vis.height
    lw = max(1.0, min(w, h) * 0.13)
    to_page = lambda pts: [tuple(fitz.Point(x, y) * page.derotation_matrix) for x, y in pts]
    if symbol == "dot":
        r = fitz.Rect(x0 + w * 0.22, y0 + h * 0.22, x0 + w * 0.78, y0 + h * 0.78)
        a = page.add_circle_annot((r * page.derotation_matrix).normalize())
        a.set_colors(stroke=color, fill=color)
        a.set_border(width=0)
    elif symbol == "cross":
        a = page.add_ink_annot([to_page([(x0 + w * 0.2, y0 + h * 0.2), (x0 + w * 0.8, y0 + h * 0.8)]),
                                to_page([(x0 + w * 0.8, y0 + h * 0.2), (x0 + w * 0.2, y0 + h * 0.8)])])
        a.set_colors(stroke=color)
        a.set_border(width=lw)
    else:   # ✓：像手寫一樣稍微超出框框右上方
        a = page.add_ink_annot([to_page([(x0 + w * 0.15, y0 + h * 0.52), (x0 + w * 0.42, y0 + h * 0.82),
                                          (x0 + w * 1.02, y0 - h * 0.12)])])
        a.set_colors(stroke=color)
        a.set_border(width=lw)
    a.set_info(title=CHECK_TAG, content="打勾")
    a.update()
    return a


def toggle_check(page, box, symbol="check", color=(0, 0, 0)):
    """框框已勾 → 取消；未勾 → 打勾。回傳 True 表示打勾、False 表示取消"""
    existing = check_annots(page, box)
    if existing:
        for a in existing:
            page.delete_annot(a)
        return False
    add_check_mark(page, box, symbol, color)
    return True


def check_all_in(page, area, symbol="check", color=(0, 0, 0)):
    """範圍內所有尚未勾選的框框都打勾，回傳勾選數量"""
    n = 0
    for b in find_boxes(page, area):
        if not check_annots(page, b):
            add_check_mark(page, b, symbol, color)
            n += 1
    return n


# ---------------- 簽名 / 印章 ----------------
def signature_dir():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    d = os.path.join(base, "CGU_PDFTool", "signatures")
    os.makedirs(d, exist_ok=True)
    return d


def list_signatures():
    d = signature_dir()
    files = [os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith(".png")]
    return sorted(files, key=os.path.getmtime)


def _new_signature_path():
    import time
    return os.path.join(signature_dir(), time.strftime("sig_%Y%m%d_%H%M%S") + f"_{int(time.time()*1000)%1000:03d}.png")


def trim_transparent(im, pad=4):
    """裁掉四周透明的空白"""
    bbox = im.getchannel("A").point(lambda a: 255 if a > 40 else 0).getbbox()
    if not bbox:
        return im
    x0, y0, x1, y1 = bbox
    return im.crop((max(0, x0 - pad), max(0, y0 - pad), min(im.width, x1 + pad), min(im.height, y1 + pad)))


def remove_white_background(im, lo=55, hi=120):
    """
    去除白底：越接近白 / 淺灰（紙張）越透明，墨水與印泥保留。
    以 255 - min(R,G,B) 判斷「有多不白」：紅色印泥、黑 / 藍墨水都很高，紙張很低。
    """
    import numpy as np
    a = np.asarray(im.convert("RGBA")).astype(np.float32)
    ink = 255 - a[..., :3].min(axis=2)
    alpha = np.clip((ink - lo) / (hi - lo), 0, 1) * (a[..., 3] / 255.0)
    out = a.copy()
    out[..., 3] = alpha * 255
    return Image.fromarray(out.astype("uint8"), "RGBA")


def import_signature_image(path, remove_bg=True):
    """匯入簽名 / 印章圖檔：去白底、裁邊、縮到最長邊 1200px，存進簽名庫並回傳路徑"""
    with Image.open(path) as im:
        im.load()
        try:
            from PIL import ImageOps
            im = ImageOps.exif_transpose(im)     # 手機照片依 EXIF 轉正
        except Exception:
            pass
        has_alpha = im.mode in ("RGBA", "LA") and im.getchannel("A").getextrema()[0] < 250
        im = im.convert("RGBA")
        if remove_bg and not has_alpha:
            im = remove_white_background(im)
        im = trim_transparent(im)
        im.thumbnail((1200, 1200))
        out = _new_signature_path()
        im.save(out)
    return out


def default_sign_width(path):
    """預設蓋章寬度：接近正方形視為印章（約 2.5 公分），長條視為簽名（約 4.5 公分）"""
    try:
        with Image.open(path) as im:
            w, h = im.size
        return 72 if 0.6 <= w / max(h, 1) <= 1.6 else 128
    except Exception:
        return 100


SIGN_MARK = "CGUSign"      # 寫在圖片物件上的標記：之後可重新選取、移動


def _unique_png(png_bytes, salt):
    """
    在第一列透明像素的透明度（0~3，肉眼看不出）寫入隨機值，讓每次蓋章都是獨立圖片物件
    （PyMuPDF 會把內容相同的圖片合併成同一個，之後就無法個別移動）。
    """
    with Image.open(io.BytesIO(png_bytes)) as im:
        im = im.convert("RGBA")
        px = im.load()
        bit = 0
        for i in range(im.width):
            if bit >= 48:
                break
            r, g, b, a = px[i, 0]
            if a < 4:
                px[i, 0] = (r, g, b, (salt >> bit) & 3)
                bit += 2
        out = io.BytesIO()
        im.save(out, "PNG")
        return out.getvalue()


def place_signature(page, visual_rect, png):
    """
    把簽名 / 印章蓋到頁面上，回傳圖片 xref。
    visual_rect 為畫面上看到的方向（旋轉頁面也保持正向）；png 為圖檔路徑或 PNG 位元組。
    """
    import random
    if isinstance(png, str):
        with open(png, "rb") as f:
            png = f.read()
    rect = (fitz.Rect(visual_rect) * page.derotation_matrix).normalize()
    xref = page.insert_image(rect, stream=_unique_png(png, random.getrandbits(48)), keep_proportion=True,
                             overlay=True, rotate=(-page.rotation) % 360)
    page.parent.xref_set_key(xref, SIGN_MARK, "true")
    return xref


def signature_at(page, point):
    """點選位置（未旋轉頁面座標）上的已蓋簽名 / 印章：回傳 (xref, 外框 Rect) 或 None"""
    doc = page.parent
    hits = []
    for info in page.get_image_info(xrefs=True):
        x = info.get("xref")
        if x and fitz.Rect(info["bbox"]).contains(point) and doc.xref_get_key(x, SIGN_MARK)[1] == "true":
            hits.append((x, fitz.Rect(info["bbox"])))
    return hits[-1] if hits else None            # 最上層（最後畫的）優先


def signature_png(doc, xref):
    """取回已蓋簽名的 PNG（含透明度）"""
    pix = fitz.Pixmap(doc, xref)
    smask = doc.xref_get_key(xref, "SMask")
    if smask[0] == "xref":
        mask = fitz.Pixmap(doc, int(smask[1].split()[0]))
        pix = fitz.Pixmap(pix, mask)
    if pix.colorspace and pix.colorspace.n not in (1, 3):
        pix = fitz.Pixmap(fitz.csRGB, pix)
    return pix.tobytes("png")


def remove_signature(page, xref):
    """移除已蓋的簽名 / 印章（換成透明圖並取消標記）"""
    page.delete_image(xref)
    page.parent.xref_set_key(xref, SIGN_MARK, "null")


class SignOverlay(QWidget):
    """
    放在畫布上的可調整簽名框：拖曳中間移動、拖四角縮放（保持比例）；
    Enter / ✓ 確定、Esc / ✕ 取消、Delete / 🗑 刪除。
    """
    HANDLE = 9
    M = 6        # 外框與控制點留白

    def __init__(self, canvas, pixmap, rect, on_commit, on_cancel, on_delete=None):
        super().__init__(canvas)
        self.pix = pixmap
        self.aspect = pixmap.width() / max(pixmap.height(), 1)
        self.r = QRectF(rect)
        self.on_commit, self.on_cancel, self.on_delete = on_commit, on_cancel, on_delete
        self.drag = None
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.bar = QFrame(canvas)
        self.bar.setObjectName("MiniBar")
        bl = QHBoxLayout(self.bar)
        bl.setContentsMargins(4, 3, 4, 3)
        bl.setSpacing(2)
        for txt, tip, fn in (("✓ 確定", "確定位置（Enter）", self.commit), ("✕", "取消（Esc）", self.cancel)) + \
                ((("刪除", "刪除這個章（Delete）", self.delete),) if on_delete else ()):
            b = QToolButton()
            b.setText(txt)
            b.setToolTip(tip)
            b.clicked.connect(fn)
            bl.addWidget(b)
        self.bar.adjustSize()
        self.relayout()
        self.show()
        self.bar.show()
        self.setFocus()

    # ---- 版面 ----
    def relayout(self):
        m = self.M
        self.setGeometry(QRect(int(self.r.x()) - m, int(self.r.y()) - m, int(self.r.width()) + 2 * m,
                               int(self.r.height()) + 2 * m))
        bx = int(self.r.x())
        by = int(self.r.y()) - self.bar.height() - m - 2
        if by < 0:
            by = int(self.r.bottom()) + m + 2
        self.bar.move(max(0, bx), by)
        self.update()

    def local_rect(self):
        return QRectF(self.M, self.M, self.r.width(), self.r.height())

    def corners(self):
        lr = self.local_rect()
        return {"tl": lr.topLeft(), "tr": lr.topRight(), "bl": lr.bottomLeft(), "br": lr.bottomRight()}

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        lr = self.local_rect()
        p.drawPixmap(lr, self.pix, QRectF(self.pix.rect()))
        p.setPen(QPen(QColor("#0B6BCB"), 1.5, Qt.PenStyle.DashLine))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRect(lr)
        p.setPen(QPen(QColor("#0B6BCB"), 1.5))
        p.setBrush(QColor("#FFFFFF"))
        h = self.HANDLE
        for c in self.corners().values():
            p.drawRect(QRectF(c.x() - h / 2, c.y() - h / 2, h, h))
        p.end()

    # ---- 滑鼠 ----
    def _hit(self, pos):
        for k, c in self.corners().items():
            if abs(pos.x() - c.x()) <= self.HANDLE and abs(pos.y() - c.y()) <= self.HANDLE:
                return k
        return "move" if self.local_rect().contains(pos) else None

    def mousePressEvent(self, e):
        if e.button() != Qt.MouseButton.LeftButton:
            return
        hit = self._hit(e.position())
        if hit:
            self.drag = (hit, e.globalPosition(), QRectF(self.r))
        self.setFocus()

    def mouseMoveEvent(self, e):
        if not self.drag:
            hit = self._hit(e.position())
            cur = {"tl": Qt.CursorShape.SizeFDiagCursor, "br": Qt.CursorShape.SizeFDiagCursor,
                   "tr": Qt.CursorShape.SizeBDiagCursor, "bl": Qt.CursorShape.SizeBDiagCursor,
                   "move": Qt.CursorShape.SizeAllCursor}.get(hit, Qt.CursorShape.ArrowCursor)
            self.setCursor(cur)
            return
        kind, start, r0 = self.drag
        d = e.globalPosition() - start
        if kind == "move":
            self.r = r0.translated(d.x(), d.y())
        else:
            # 以對角為錨點縮放，寬度決定大小、高度依比例
            anchor = {"tl": r0.bottomRight(), "tr": r0.bottomLeft(), "bl": r0.topRight(), "br": r0.topLeft()}[kind]
            sx = -1 if kind in ("tl", "bl") else 1
            w = max(16.0, r0.width() + sx * d.x())
            h = w / self.aspect
            x = anchor.x() - w if sx < 0 else anchor.x()
            y = anchor.y() - h if kind in ("tl", "tr") else anchor.y()
            self.r = QRectF(x, y, w, h)
        self.relayout()

    def mouseReleaseEvent(self, e):
        self.drag = None

    def mouseDoubleClickEvent(self, e):
        self.commit()

    def keyPressEvent(self, e):
        k = e.key()
        if k in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.commit()
        elif k == Qt.Key.Key_Escape:
            self.cancel()
        elif k in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace) and self.on_delete:
            self.delete()
        elif k in (Qt.Key.Key_Left, Qt.Key.Key_Right, Qt.Key.Key_Up, Qt.Key.Key_Down):
            step = 10 if e.modifiers() & Qt.KeyboardModifier.ShiftModifier else 1
            dx = {Qt.Key.Key_Left: -step, Qt.Key.Key_Right: step}.get(k, 0)
            dy = {Qt.Key.Key_Up: -step, Qt.Key.Key_Down: step}.get(k, 0)
            self.r.translate(dx, dy)
            self.relayout()
        else:
            super().keyPressEvent(e)

    # ---- 結束 ----
    def _close(self):
        self.bar.deleteLater()
        self.deleteLater()

    def commit(self):
        r = QRectF(self.r)
        self._close()
        self.on_commit(r)

    def cancel(self):
        self._close()
        self.on_cancel()

    def delete(self):
        self._close()
        self.on_delete()


class SignaturePad(QWidget):
    """手寫簽名板：按住滑鼠左鍵書寫"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(520, 220)
        self.strokes = []
        self.color = QColor("#1A1A1A")
        self.pen_w = 3.0
        self.setCursor(Qt.CursorShape.CrossCursor)

    def clear(self):
        self.strokes = []
        self.update()

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.strokes.append([e.position()])
            self.update()

    def mouseMoveEvent(self, e):
        if self.strokes and e.buttons() & Qt.MouseButton.LeftButton:
            self.strokes[-1].append(e.position())
            self.update()

    def _draw(self, p, scale=1.0):
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        pen = QPen(self.color, self.pen_w * scale, Qt.PenStyle.SolidLine,
                   Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
        p.setPen(pen)
        for st in self.strokes:
            pts = [QPointF(q.x() * scale, q.y() * scale) for q in st]
            if len(pts) == 1:
                p.drawPoint(pts[0])
            else:
                path = QPainterPath(pts[0])
                for q in pts[1:]:
                    path.lineTo(q)
                p.drawPath(path)

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("#FFFFFF"))
        p.setPen(QPen(QColor("#C8C8C8"), 1, Qt.PenStyle.DashLine))
        y = int(self.height() * 0.72)
        p.drawLine(20, y, self.width() - 20, y)
        p.setPen(QColor("#B0B0B0"))
        if not self.strokes:
            p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "在此用滑鼠簽名")
        self._draw(p)
        p.end()

    def to_image(self, scale=3):
        """輸出透明背景、高解析度的簽名圖（PIL RGBA）"""
        img = QImage(self.width() * scale, self.height() * scale, QImage.Format.Format_RGBA8888)
        img.fill(Qt.GlobalColor.transparent)
        p = QPainter(img)
        self._draw(p, scale)
        p.end()
        ptr = img.constBits()
        ptr.setsize(img.sizeInBytes())
        im = Image.frombuffer("RGBA", (img.width(), img.height()), bytes(ptr), "raw", "RGBA",
                              img.bytesPerLine(), 1)
        return trim_transparent(im.copy(), pad=6)


class SignatureDrawDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("手寫簽名")
        self.setModal(True)
        lay = QVBoxLayout(self)
        self.pad = SignaturePad()
        lay.addWidget(self.pad)
        row = QHBoxLayout()
        row.addWidget(QLabel("筆色："))
        self.cmb_color = QComboBox()
        for name, c in (("黑色", "#1A1A1A"), ("藍色", "#1B3FA0"), ("紅色", "#B7282E")):
            self.cmb_color.addItem(name, c)
        self.cmb_color.currentIndexChanged.connect(self._style)
        row.addWidget(self.cmb_color)
        row.addWidget(QLabel("粗細："))
        self.cmb_w = QComboBox()
        for name, w_ in (("細", 2.0), ("中", 3.0), ("粗", 4.5)):
            self.cmb_w.addItem(name, w_)
        self.cmb_w.setCurrentIndex(1)
        self.cmb_w.currentIndexChanged.connect(self._style)
        row.addWidget(self.cmb_w)
        row.addStretch()
        btn_clear = QPushButton("清除重寫")
        btn_clear.clicked.connect(self.pad.clear)
        row.addWidget(btn_clear)
        lay.addLayout(row)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("儲存簽名")
        bb.button(QDialogButtonBox.StandardButton.Ok).setProperty("primary", True)
        bb.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        bb.accepted.connect(self._accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _style(self):
        self.pad.color = QColor(self.cmb_color.currentData())
        self.pad.pen_w = self.cmb_w.currentData()
        self.pad.update()

    def _accept(self):
        if not self.pad.strokes:
            QMessageBox.warning(self, "手寫簽名", "請先簽名。")
            return
        self.accept()

    def save(self):
        out = _new_signature_path()
        self.pad.to_image().save(out)
        return out


# ---------------- 超連結 ----------------
def find_line_at(page, point):
    """找出包含指定座標點的整行文字，回傳 (外框 Rect, 文字) 或 None"""
    d = page.get_text("dict")
    for block in d["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            if fitz.Rect(line["bbox"]).contains(point):
                text = "".join(sp["text"] for sp in line["spans"]).strip()
                if text:
                    return fitz.Rect(line["bbox"]), text
    return None


def guess_page_number(text, page_count):
    """目錄常見寫法「第一章 ........ 12」：取行尾數字當預設頁碼"""
    import re
    m = re.search(r"(\d+)\s*$", text or "")
    if m and 1 <= int(m.group(1)) <= page_count:
        return int(m.group(1))
    return None


def normalize_url(url):
    """補上網址開頭：www.xxx → https://www.xxx；含 @ 的視為 Email"""
    url = url.strip()
    if not url:
        return url
    low = url.lower()
    if low.startswith(("http://", "https://", "mailto:", "ftp://", "file:")):
        return url
    if "@" in url and "/" not in url:
        return "mailto:" + url
    return "https://" + url


def read_links(page):
    """
    直接從 PDF 註解資料讀出本頁連結（不經 page.get_links()，避免 PyMuPDF 頁面快取讀到舊資料）。
    回傳 [{"xref", "rect"（未旋轉的頁面座標）, "kind": "uri" / "page" / "other", "uri", "page"}]
    """
    import re
    doc = page.parent
    out, page_xrefs = [], None
    for xref, atype, _ in page.annot_xrefs():
        if atype != fitz.PDF_ANNOT_LINK:
            continue
        t, v = doc.xref_get_key(xref, "Rect")
        nums = [float(x) for x in re.findall(r"-?[\d.]+", v)][:4] if t == "array" else []
        if len(nums) < 4:
            continue
        info = {"xref": xref, "rect": (fitz.Rect(nums) * page.transformation_matrix).normalize(),
                "kind": "other", "uri": None, "page": None}
        act = doc.xref_get_key(xref, "A/S")[1]
        if act == "/URI":
            u = doc.xref_get_key(xref, "A/URI")
            info["kind"], info["uri"] = "uri", (u[1] if u[0] == "string" else "")
        else:
            dt, dv = doc.xref_get_key(xref, "A/D") if act == "/GoTo" else doc.xref_get_key(xref, "Dest")
            m = re.match(r"\[\s*(\d+)\s+0\s+R", dv or "") if dt == "array" else None
            if m:
                if page_xrefs is None:
                    page_xrefs = {doc.page_xref(i): i for i in range(len(doc))}
                pno = page_xrefs.get(int(m.group(1)))
                if pno is not None:
                    info["kind"], info["page"] = "page", pno
        out.append(info)
    return out


def link_at(page, point):
    """回傳包含指定座標點（未旋轉頁面座標）的連結，或 None；重疊時取最小的那個"""
    hits = [ln for ln in read_links(page) if ln["rect"].contains(point)]
    return min(hits, key=lambda ln: ln["rect"].width * ln["rect"].height) if hits else None


LINK_BG_TAG = "CGU_LINK_BG"            # 連結底色註解的識別標記
LINK_BG_COLOR = (0.84, 0.91, 1.0)      # 淺藍色（以「色彩增值」疊在文字上，文字仍保持原色）


def link_bg_annots(page, rect):
    """找出與連結範圍對應的底色註解"""
    r = fitz.Rect(rect)
    out = []
    for a in page.annots() or []:
        if a.info.get("title") == LINK_BG_TAG and abs(a.rect.x0 - r.x0) < 2 and abs(a.rect.y0 - r.y0) < 2 \
                and abs(a.rect.x1 - r.x1) < 2 and abs(a.rect.y1 - r.y1) < 2:
            out.append(a)
    return out


def has_link_bg(page, rect):
    return bool(link_bg_annots(page, rect))


def add_link(page, rect, target, bg=True):
    """
    target：("uri", 網址) 或 ("page", 頁碼 0 起算)；rect 為未旋轉的頁面座標。
    bg=True 時先加一層淺藍底色（方框註解，無框線），再加連結 → 連結在最上層，點擊不受影響。
    """
    rect = fitz.Rect(rect)
    if bg and not has_link_bg(page, rect):
        a = page.add_rect_annot(rect)
        a.set_colors(stroke=LINK_BG_COLOR, fill=LINK_BG_COLOR)
        a.set_border(width=0)
        a.set_blendmode(fitz.PDF_BM_Multiply)
        a.set_info(title=LINK_BG_TAG, content="超連結提示底色")
        a.set_flags(fitz.PDF_ANNOT_IS_PRINT | fitz.PDF_ANNOT_IS_LOCKED | fitz.PDF_ANNOT_IS_READ_ONLY)
        a.update()
    kind, val = target
    if kind == "uri":
        page.insert_link({"kind": fitz.LINK_URI, "from": rect, "uri": val})
    else:
        page.insert_link({"kind": fitz.LINK_GOTO, "from": rect, "page": int(val),
                          "to": fitz.Point(0, 0), "zoom": 0})


def remove_link(page, link, keep_bg=False):
    """依位置與種類找回連結再刪除（復原後 xref 可能改變，不直接用舊 xref）；預設連同底色一起刪"""
    for ln in read_links(page):
        if ln["rect"] == link["rect"] and ln["kind"] == link["kind"] \
                and ln["uri"] == link["uri"] and ln["page"] == link["page"]:
            page.delete_link({"xref": ln["xref"], "kind": fitz.LINK_NONE, "from": ln["rect"]})
            break
    if not keep_bg:
        for a in link_bg_annots(page, link["rect"]):
            page.delete_annot(a)


def replace_link(page, old, rect, target, bg=True):
    remove_link(page, old)
    add_link(page, rect, target, bg)


class LinkDialog(QDialog):
    """超連結設定：連到網址 / 連到本檔第幾頁；編輯既有連結時可刪除"""
    DELETE = 2

    def __init__(self, parent, page_count, current_page, text=None, link=None, bg=True):
        super().__init__(parent)
        self.setWindowTitle("編輯超連結" if link else "新增超連結")
        self.setModal(True)
        self.setMinimumWidth(420)
        lay = QVBoxLayout(self)
        if text:
            t = QLabel(f"連結文字：{text if len(text) <= 40 else text[:40] + '…'}")
            t.setObjectName("PanelHint")
            t.setWordWrap(True)
            lay.addWidget(t)
        self.rb_page = QRadioButton("跳到本檔第幾頁")
        self.rb_uri = QRadioButton("開啟網址（網站或 Email）")
        self.spin_page = QSpinBox()
        self.spin_page.setRange(1, max(1, page_count))
        self.spin_page.setSuffix(f"  / {page_count} 頁")
        self.edit_uri = QLineEdit()
        self.edit_uri.setPlaceholderText("例如 https://www.cgu.edu.tw")
        prow = QHBoxLayout(); prow.addWidget(self.rb_page); prow.addWidget(self.spin_page, 1)
        urow = QHBoxLayout(); urow.addWidget(self.rb_uri); urow.addWidget(self.edit_uri, 1)
        lay.addLayout(prow)
        lay.addLayout(urow)
        # 預設值
        guess = guess_page_number(text, page_count)
        if link and link["kind"] == "uri":
            self.rb_uri.setChecked(True)
            self.edit_uri.setText(link["uri"] or "")
            self.spin_page.setValue(current_page + 1)
        elif link and link["kind"] == "page":
            self.rb_page.setChecked(True)
            self.spin_page.setValue(link["page"] + 1)
        else:
            self.rb_page.setChecked(True)
            self.spin_page.setValue(guess or min(current_page + 2, page_count))
        self.chk_bg = QCheckBox("加上淺藍底色，提醒讀者這裡可以點")
        self.chk_bg.setChecked(bg)
        lay.addWidget(self.chk_bg)
        self.edit_uri.textEdited.connect(lambda _: self.rb_uri.setChecked(True))
        self.spin_page.valueChanged.connect(lambda _: self.rb_page.setChecked(True))
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("確定")
        bb.button(QDialogButtonBox.StandardButton.Ok).setProperty("primary", True)
        bb.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        if link:
            btn_del = bb.addButton("刪除連結", QDialogButtonBox.ButtonRole.DestructiveRole)
            btn_del.setProperty("variant", "outline-danger")
            btn_del.clicked.connect(lambda: self.done(self.DELETE))
        bb.accepted.connect(self._accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _accept(self):
        if self.rb_uri.isChecked() and not normalize_url(self.edit_uri.text()):
            QMessageBox.warning(self, "超連結", "請輸入網址。")
            return
        self.accept()

    def target(self):
        if self.rb_uri.isChecked():
            return ("uri", normalize_url(self.edit_uri.text()))
        return ("page", self.spin_page.value() - 1)


def find_span_at(page, point):
    """在頁面上找出包含指定座標點的文字片段（span），回傳 span dict 或 None"""
    d = page.get_text("dict")
    for block in d["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                if fitz.Rect(span["bbox"]).contains(point) and span["text"].strip():
                    span = dict(span)
                    span["dir"] = line.get("dir", (1, 0))
                    return span
    return None


def add_text_box(page, point, text, fontsize=14, color=(0, 0, 0), fontfile=None):
    """在指定位置（左上角）新增文字，支援中文與多行"""
    kw = {"fontname": "helv"} if is_ascii(text) else font_kwargs(fontfile)
    page.insert_text(fitz.Point(point.x, point.y + fontsize), text,
                     fontsize=fontsize, color=color, **kw)


def insert_image_at(page, point, image_path, width=200):
    """在指定位置插入圖片，寬度預設 200pt，高度依比例"""
    with Image.open(image_path) as im:
        w, h = im.size
    height = width * h / w
    rect = fitz.Rect(point.x, point.y, point.x + width, point.y + height)
    page.insert_image(rect, filename=image_path, keep_proportion=True)
    return rect


def highlight_area(page, rect, color=(1, 1, 0)):
    """螢光標記：範圍內有文字就標記那些字，否則整個框標記"""
    words = [fitz.Rect(w[:4]) for w in page.get_text("words") if fitz.Rect(w[:4]).intersects(rect)]
    annot = page.add_highlight_annot(words if words else rect)
    annot.set_colors(stroke=color)
    annot.update()
    return annot


def whiteout_area(page, rect):
    """塗白 / 塗銷：永久刪除範圍內的文字與圖片，並以白色填滿"""
    page.add_redact_annot(rect, fill=(1, 1, 1))
    page.apply_redactions()


def delete_annot_at(page, point):
    """刪除包含指定座標點的註解（便利貼、標記、方框等），回傳是否有刪除"""
    for annot in page.annots() or []:
        if annot.rect.contains(point):
            page.delete_annot(annot)
            return True
    return False


# =====================================================================
# 字型比對：沿用原嵌入字型 → 對應系統字型 → 內建字型
# =====================================================================

def resource_path(rel):
    """取得資源檔路徑；打包成 exe 後資源會被解壓到 sys._MEIPASS"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def font_dirs():
    """系統字型資料夾（Windows 為主，Linux/macOS 供測試或跨平台使用）"""
    dirs = [os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
            "/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts"),
            "/System/Library/Fonts", "/Library/Fonts"]
    return [d for d in dirs if d and os.path.isdir(d)]


_FONT_INDEX = None


def font_index():
    """建立「小寫檔名 → 完整路徑」索引（只掃描一次）"""
    global _FONT_INDEX
    if _FONT_INDEX is None:
        _FONT_INDEX = {}
        for d in font_dirs():
            for root, _, files in os.walk(d):
                for f in files:
                    if f.lower().endswith((".ttf", ".ttc", ".otf")):
                        _FONT_INDEX.setdefault(f.lower(), os.path.join(root, f))
    return _FONT_INDEX


# 字型名稱關鍵字 → 候選系統字型檔（依序嘗試；含 Linux 對應字型供測試）
# 每項：(關鍵字, 一般, 粗體, 斜體, 粗斜體)
FONT_MAP = [
    (("標楷", "楷", "dfkaishu", "dfkai", "kaiu", "biaukai", "kaiti", "kai"), ["kaiu.ttf", "simkai.ttf"], None, None, None),
    (("新細明", "細明", "pmingliu", "mingliu", "ming", "mincho", "song", "宋"),
     ["mingliu.ttc", "mingliub.ttc", "notoserifcjk-regular.ttc"], ["notoserifcjk-bold.ttc"], None, None),
    (("微軟正黑", "正黑", "jhenghei", "msjh"), ["msjh.ttc", "msjh.ttf", "notosanscjk-regular.ttc"],
     ["msjhbd.ttc", "msjhbd.ttf", "notosanscjk-bold.ttc"], None, None),
    (("微軟雅黑", "yahei", "msyh"), ["msyh.ttc", "msjh.ttc", "notosanscjk-regular.ttc"], ["msyhbd.ttc"], None, None),
    (("calibri",), ["calibri.ttf", "carlito-regular.ttf"], ["calibrib.ttf", "carlito-bold.ttf"],
     ["calibrii.ttf", "carlito-italic.ttf"], ["calibriz.ttf", "carlito-bolditalic.ttf"]),
    (("cambria",), ["cambria.ttc", "caladea-regular.ttf"], ["cambriab.ttf", "caladea-bold.ttf"],
     ["cambriai.ttf", "caladea-italic.ttf"], ["cambriaz.ttf", "caladea-bolditalic.ttf"]),
    (("times", "roman", "tiro"), ["times.ttf", "liberationserif-regular.ttf", "dejavuserif.ttf"],
     ["timesbd.ttf", "liberationserif-bold.ttf"], ["timesi.ttf", "liberationserif-italic.ttf"],
     ["timesbi.ttf", "liberationserif-bolditalic.ttf"]),
    (("courier", "mono", "consolas"), ["cour.ttf", "consola.ttf", "liberationmono-regular.ttf", "dejavusansmono.ttf"],
     ["courbd.ttf", "liberationmono-bold.ttf"], ["couri.ttf", "liberationmono-italic.ttf"],
     ["courbi.ttf", "liberationmono-bolditalic.ttf"]),
    (("arial", "helvetica", "helv", "sans"), ["arial.ttf", "liberationsans-regular.ttf", "dejavusans.ttf"],
     ["arialbd.ttf", "liberationsans-bold.ttf", "dejavusans-bold.ttf"],
     ["ariali.ttf", "liberationsans-italic.ttf"], ["arialbi.ttf", "liberationsans-bolditalic.ttf"]),
    (("georgia",), ["georgia.ttf", "liberationserif-regular.ttf"], ["georgiab.ttf"], ["georgiai.ttf"], ["georgiaz.ttf"]),
    (("verdana", "tahoma", "segoe"), ["verdana.ttf", "tahoma.ttf", "segoeui.ttf", "dejavusans.ttf"],
     ["verdanab.ttf", "tahomabd.ttf", "segoeuib.ttf"], None, None),
]
# 中文（CJK）通用備援：微軟正黑 → 新細明 → 標楷 → Noto CJK
CJK_FALLBACK = ["msjh.ttc", "msjh.ttf", "mingliu.ttc", "kaiu.ttf", "notosanscjk-regular.ttc",
                "notoserifcjk-regular.ttc", "droidsansfallbackfull.ttf"]


def strip_subset(name):
    """
    去掉子集字型前綴，例如 'ABCDEF+Calibri' → 'Calibri'；
    並解碼 PDF 名稱中的 #XX 十六進位（中文字型名常以 Big5/GBK 編碼，如 #BC#D0#B7#A2#C5#E9 = 標楷體）
    """
    name = name or ""
    if len(name) > 7 and name[6] == "+" and name[:6].isupper():
        name = name[7:]
    if "#" in name:
        try:
            raw = bytearray()
            i = 0
            while i < len(name):
                if name[i] == "#" and i + 2 < len(name) + 1 and all(c in "0123456789abcdefABCDEF" for c in name[i + 1:i + 3]):
                    raw.append(int(name[i + 1:i + 3], 16))
                    i += 3
                else:
                    raw += name[i].encode("utf-8")
                    i += 1
            for enc in ("utf-8", "big5", "gbk"):
                try:
                    name = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
        except Exception:
            pass
    return name


# 編輯對話框 / 屬性面板可選的字型（顯示名稱, 候選檔名）；只列出系統中實際存在的
UI_FONTS = [
    ("標楷體", ["kaiu.ttf"]),
    ("新細明體", ["mingliu.ttc", "pmingliu.ttf"]),
    ("細明體", ["mingliu.ttc", "mingliu.ttf"]),
    ("微軟正黑體", ["msjh.ttc", "msjh.ttf"]),
    ("微軟雅黑", ["msyh.ttc", "msyh.ttf"]),
    ("Times New Roman", ["times.ttf"]),
    ("Arial", ["arial.ttf"]),
    ("Calibri", ["calibri.ttf"]),
    ("Cambria", ["cambria.ttc"]),
    ("Courier New", ["cour.ttf"]),
    # 非 Windows（測試 / Linux）對應字型
    ("Noto Serif CJK（明體類）", ["notoserifcjk-regular.ttc"]),
    ("Noto Sans CJK（黑體類）", ["notosanscjk-regular.ttc"]),
    ("AR PL UKai（楷體類）", ["ukai.ttc"]),
    ("Liberation Serif", ["liberationserif-regular.ttf"]),
    ("Liberation Sans", ["liberationsans-regular.ttf"]),
]


def list_ui_fonts():
    """回傳 [(顯示名稱, 字型檔路徑)]，只含系統中找得到的字型"""
    idx = font_index()
    out = []
    for label, files in UI_FONTS:
        for f in files:
            p = idx.get(f.lower())
            if p:
                out.append((label, p))
                break
    return out


def default_ocr_font_path():
    """OCR 替換字型的預設值：有標楷體就用標楷體，否則第一個可用的中文字型"""
    fonts = dict(list_ui_fonts())
    for k in ("標楷體", "新細明體", "AR PL UKai（楷體類）", "Noto Serif CJK（明體類）", "微軟正黑體"):
        if k in fonts:
            return fonts[k]
    return None


def font_info_from_path(path, label=None):
    """把字型檔包成 choose_font() 相同格式的 dict"""
    f = fitz.Font(fontfile=path)
    return {"kind": "user", "label": f"{label or os.path.basename(path)}（{os.path.basename(path)}）",
            "insert": {"fontname": "U" + str(abs(hash(path)) % 100000), "fontfile": path}, "font": f}


def has_cjk(text):
    return any(ord(c) > 0x2E80 for c in text)


def font_covers(font, text):
    """確認字型含有文字中所有字元（空白除外）"""
    try:
        return all(font.has_glyph(ord(c)) for c in text if not c.isspace())
    except Exception:
        return False


def load_font_file(path):
    try:
        return fitz.Font(fontfile=path)
    except Exception:
        return None


def system_font_candidates(fontname, flags=0, text=""):
    """依原字型名稱與粗/斜/襯線/等寬旗標，列出可能對應的系統字型檔路徑"""
    idx = font_index()
    low = strip_subset(fontname).lower().replace(" ", "")
    bold = bool(flags & 16) or "bold" in low or "-bd" in low or "粗" in low
    italic = bool(flags & 2) or "italic" in low or "oblique" in low
    serif, mono = bool(flags & 4), bool(flags & 8)
    names = []
    for keys, reg, bd, it, bi in FONT_MAP:
        if any(k in low for k in keys):
            style = (bi if bold and italic else bd if bold else it if italic else None) or []
            names += style + reg
    if not names:   # 名稱認不出來：依旗標推測
        if mono:
            names += ["cour.ttf", "consola.ttf", "liberationmono-regular.ttf"]
        elif serif:
            names += (["timesbd.ttf"] if bold else []) + ["times.ttf", "liberationserif-regular.ttf"]
        else:
            names += (["arialbd.ttf", "liberationsans-bold.ttf"] if bold else []) + \
                     ["arial.ttf", "liberationsans-regular.ttf", "dejavusans.ttf"]
    if has_cjk(text):
        names += (["msjhbd.ttc", "notosanscjk-bold.ttc"] if bold else []) + CJK_FALLBACK
    seen, out = set(), []
    for n in names:
        p = idx.get(n.lower())
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return out


def find_page_font_xref(page, fontname):
    """依 span 的字型名稱找出頁面字型資源的 xref"""
    target = strip_subset(fontname)
    for xref, ext, ftype, basefont, *_ in page.get_fonts(full=True):
        b = strip_subset(basefont).replace(" ", "").lower()
        t = target.replace(" ", "").lower()
        # span 的字型名稱可能被截短，因此也接受「開頭相同」
        if t and (b == t or b.startswith(t) or t.startswith(b)):
            return xref
    return 0


def choose_font(page, span, text):
    """
    決定替換文字要用的字型，回傳 dict：
      kind  : 'embedded'（沿用原嵌入字型）/ 'system'（系統字型）/ 'builtin'（內建字型）
      label : 顯示給使用者看的說明
      insert: 傳給 page.insert_text 的字型參數
      font  : fitz.Font（用來量文字寬度）
    """
    fname = span.get("font", "")
    flags = span.get("flags", 0)
    # 1) 沿用原 PDF 內嵌的字型（子集字型可能缺字，逐字檢查）
    xref = find_page_font_xref(page, fname)
    reason = "原文件字型未內嵌"
    if xref:
        try:
            name, ext, ftype, buf = page.parent.extract_font(xref)
            if buf and ext not in ("n/a", "") and ftype != "Type3":
                reason = "原內嵌字型為子集，缺少需要的字或無字元對照表"
                f = fitz.Font(fontbuffer=buf)
                if font_covers(f, text):
                    return {"kind": "embedded", "label": f"原文件內嵌字型：{strip_subset(fname)}",
                            "insert": {"fontname": f"E{xref}", "fontbuffer": buf}, "font": f}
        except Exception:
            pass
    # 2) 對應的系統字型
    for path in system_font_candidates(fname, flags, text):
        f = load_font_file(path)
        if f and font_covers(f, text):
            return {"kind": "system", "label": f"系統字型：{os.path.basename(path)}（{f.name}）\n（{reason}）",
                    "insert": {"fontname": "S" + str(abs(hash(path)) % 100000), "fontfile": path}, "font": f}
    # 3) 內建字型
    if not has_cjk(text) and is_ascii(text):
        bold, italic = bool(flags & 16), bool(flags & 2)
        low = fname.lower()
        fam = "tiro" if ("times" in low or flags & 4) else "cour" if ("cour" in low or flags & 8) else "helv"
        code = {"helv": ["helv", "hebo", "heit", "hebi"], "tiro": ["tiro", "tibo", "tiit", "tibi"],
                "cour": ["cour", "cobo", "coit", "cobi"]}[fam][bold * 1 + italic * 2]
        return {"kind": "builtin", "label": f"內建字型：{code}", "insert": {"fontname": code},
                "font": fitz.Font(code)}
    f = fitz.Font("cjk")
    return {"kind": "builtin", "label": "內建中文字型：Droid Sans Fallback",
            "insert": {"fontname": "CJKB", "fontbuffer": f.buffer}, "font": f}


def is_invisible_span(page, span):
    """判斷 span 是否為 OCR 產生的隱形文字（texttrace type 3 = 不可見）"""
    bb = fitz.Rect(span["bbox"])
    try:
        for t in page.get_texttrace():
            if t.get("type") == 3 and fitz.Rect(t["bbox"]).intersects(bb):
                return True
    except Exception:
        pass
    return False


def sample_colors(page, rect, dpi=150):
    """在頁面影像上取樣：背景色（框外圍一圈的中位數）與文字色（框內最深的像素）"""
    r = fitz.Rect(rect) + (-3, -3, 3, 3)
    r &= page.rect
    pix = page.get_pixmap(clip=r, dpi=dpi, annots=False)
    n, w, h, s = pix.n, pix.width, pix.height, pix.samples
    ring, inner = [], []
    for y in range(h):
        for x in range(w):
            i = (y * w + x) * n
            px = (s[i], s[i + 1], s[i + 2]) if n >= 3 else (s[i],) * 3
            (ring if (x < 3 or y < 3 or x >= w - 3 or y >= h - 3) else inner).append(px)
    def med(lst):
        if not lst:
            return (1, 1, 1)
        return tuple(sorted(c[k] for c in lst)[len(lst) // 2] / 255 for k in range(3))
    bg = med(ring)
    inner.sort(key=sum)
    fg = med(inner[: max(1, len(inner) // 40)]) if inner else (0, 0, 0)   # 最深的 2.5% 像素 ≈ 筆畫顏色
    return bg, fg


def _ink_bbox_from_pixmap(pix, thresh, skip_lines=True):
    """在灰階 pixmap 中找出比 thresh 暗的像素範圍（略過貫穿整列/整欄的表格線）→ (x0,y0,x1,y1) 像素或 None"""
    w, h, s, n = pix.width, pix.height, pix.samples, pix.n
    rows = [0] * h
    cols = [0] * w
    dark = []
    for y in range(h):
        base = y * w * n
        for x in range(w):
            if s[base + x * n] < thresh:
                rows[y] += 1
                cols[x] += 1
                dark.append((x, y))
    if not dark:
        return None
    bad_r = {y for y in range(h) if skip_lines and rows[y] > 0.85 * w}
    bad_c = {x for x in range(w) if skip_lines and cols[x] > 0.85 * h}
    pts = [(x, y) for x, y in dark if y not in bad_r and x not in bad_c]
    if not pts:
        return None
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return min(xs), min(ys), max(xs) + 1, max(ys) + 1


def measure_ink(page, rect, bg=(1, 1, 1), fg=(0, 0, 0), dpi=300):
    """量測頁面影像在 rect 範圍內「字跡」的實際範圍（PDF 座標）；量不到回傳 None"""
    r = fitz.Rect(rect) & page.rect
    if r.is_empty:
        return None
    pix = page.get_pixmap(clip=r, dpi=dpi, colorspace=fitz.csGRAY, annots=False)
    lum = lambda c: (0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]) * 255
    thresh = (lum(bg) + lum(fg)) / 2
    bb = _ink_bbox_from_pixmap(pix, thresh)
    if not bb:
        return None
    k = 72 / dpi
    return fitz.Rect(r.x0 + bb[0] * k, r.y0 + bb[1] * k, r.x0 + bb[2] * k, r.y0 + bb[3] * k)


_INK_CACHE = {}


def font_ink_metrics(fi, text, ref=100):
    """
    以字型 fi 在字級 ref 畫出 text，量出字跡相對於基線原點的範圍：
    回傳 (left, top, right, bottom)（單位：pt，top 為負值 = 基線上方）
    """
    key = (fi["label"], text, ref)
    if key in _INK_CACHE:
        return _INK_CACHE[key]
    d = fitz.open()
    width = fi["font"].text_length(text, fontsize=ref) + ref * 2
    p = d.new_page(width=max(200, width), height=ref * 3)
    ins = dict(fi["insert"])
    if "fontbuffer" in ins:
        p.insert_font(fontname=ins["fontname"], fontbuffer=ins.pop("fontbuffer"))
    ox, oy = ref * 0.5, ref * 2
    p.insert_text((ox, oy), text, fontsize=ref, color=(0, 0, 0), **ins)
    pix = p.get_pixmap(dpi=72, colorspace=fitz.csGRAY)
    bb = _ink_bbox_from_pixmap(pix, 128, skip_lines=False)
    res = (bb[0] - ox, bb[1] - oy, bb[2] - ox, bb[3] - oy) if bb else None
    _INK_CACHE[key] = res
    return res


def ocr_fit(page, span, fi, st=None):
    """
    OCR 文字：讓替換字型的「字跡高度」與原影像相同，並對齊原字跡的左緣與底緣。
    回傳 (字級, 基線原點)；量測失敗時退回以框高估計。
    """
    st = st or span_style(page, span)
    bbox = fitz.Rect(span["bbox"])
    ink = st.get("ink")
    m = font_ink_metrics(fi, span["text"].strip() or "字") if ink else None
    if ink and m and (m[3] - m[1]) > 0:
        size = 100 * ink.height / (m[3] - m[1])
        origin = fitz.Point(ink.x0 - m[0] * size / 100, ink.y1 - m[3] * size / 100)
        return size, origin
    size = bbox.height * 0.95
    return size, fitz.Point(bbox.x0, bbox.y1 - bbox.height * 0.18)


def span_style(page, span):
    """取得 span 的顯示資訊：字型名稱、大小、顏色、是否 OCR"""
    ocr = is_invisible_span(page, span)
    if ocr:
        bg, fg = sample_colors(page, span["bbox"])
        ink = measure_ink(page, span["bbox"], bg, fg)
        h = ink.height if ink else fitz.Rect(span["bbox"]).height * 0.8
        # 粗估字級：中文字跡約佔字級 0.88，純英數約 0.72（精確值由 ocr_fit 依實際字型計算）
        est = h / (0.88 if has_cjk(span["text"]) else 0.72)
        return {"ocr": True, "font": "（掃描影像，OCR 辨識）", "size": est,
                "color": fg, "bg": bg, "ink": ink}
    return {"ocr": False, "font": strip_subset(span.get("font", "")), "size": span.get("size", 11),
            "color": int_color_to_rgb01(span.get("color", 0)), "bg": None}


# OCR 文字替換時的預設字型檔（由屬性面板設定；None = 自動）
OCR_DEFAULT_FONT = None


def resolve_font(page, span, choice, text, st=None):
    """依使用者選擇（字型檔路徑 / 'original' / None=自動）決定替換字型"""
    st = st or span_style(page, span)
    t = text or " "
    if choice and choice != "original" and os.path.exists(choice):
        return font_info_from_path(choice)
    if choice == "original" or not st["ocr"]:
        return choose_font(page, span, t)
    if OCR_DEFAULT_FONT and os.path.exists(OCR_DEFAULT_FONT):
        return font_info_from_path(OCR_DEFAULT_FONT, "OCR 預設字型")
    return choose_font(page, {"font": "PMingLiU" if has_cjk(t) else "Times New Roman", "flags": 4}, t)


def default_edit_size(page, span, fi, st=None):
    """編輯時的預設字級：原生文字 = 原字級；OCR 文字 = 依字跡高度校正後的字級"""
    st = st or span_style(page, span)
    if st["ocr"]:
        return ocr_fit(page, span, fi, st)[0]
    return span.get("size", 11)


# 字型檔的粗體 / 斜體版本（Windows 常見字型 + Noto；找不到就用模擬）
_STYLE_VARIANTS = {
    "msjh.ttc": ("msjhbd.ttc", None, None), "msjhl.ttc": ("msjhbd.ttc", None, None),
    "msyh.ttc": ("msyhbd.ttc", None, None),
    "arial.ttf": ("arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    "times.ttf": ("timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    "calibri.ttf": ("calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    "cour.ttf": ("courbd.ttf", "couri.ttf", "courbi.ttf"),
    "segoeui.ttf": ("segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"),
    "tahoma.ttf": ("tahomabd.ttf", None, None),
    "verdana.ttf": ("verdanab.ttf", "verdanai.ttf", "verdanaz.ttf"),
}


def span_style_flags(span):
    """原生文字 span 的 (粗體, 斜體)：PyMuPDF flags bold=16、italic=2，另參考字型名稱"""
    flags = span.get("flags", 0)
    name = (span.get("font") or "").lower()
    bold = bool(flags & 16) or any(k in name for k in ("bold", "black", "heavy", "semibold", ",bd"))
    italic = bool(flags & 2) or any(k in name for k in ("italic", "oblique"))
    return bold, italic


def styled_font(fi, bold, italic):
    """
    取得指定粗 / 斜體的字型 → (字型資訊, 需模擬粗體, 需模擬斜體)
    先看目前字型本身是否已是粗 / 斜體，其次找真正的粗 / 斜體字型檔或內建字型，最後才模擬。
    """
    f = fi["font"]
    has_b, has_i = bool(getattr(f, "is_bold", False)), bool(getattr(f, "is_italic", False))
    need_b, need_i = bold and not has_b, italic and not has_i
    if not (need_b or need_i):
        return fi, False, False
    ins = fi.get("insert", {})
    code = ins.get("fontname", "")
    if "fontfile" not in ins and "fontbuffer" not in ins and code[:2] in ("he", "ti", "co"):
        fam = {"he": ["helv", "hebo", "heit", "hebi"], "ti": ["tiro", "tibo", "tiit", "tibi"],
               "co": ["cour", "cobo", "coit", "cobi"]}[code[:2]]
        c2 = fam[(bold or has_b) * 1 + (italic or has_i) * 2]
        return {"kind": "builtin", "label": f"內建字型：{c2}", "insert": {"fontname": c2}, "font": fitz.Font(c2)}, False, False
    path = ins.get("fontfile")
    if path:
        d, base = os.path.split(path)
        cands = []
        var = _STYLE_VARIANTS.get(base.lower())
        want_b, want_i = bold or has_b, italic or has_i
        if var:
            idx = {(True, False): 0, (False, True): 1, (True, True): 2}.get((want_b, want_i))
            if idx is not None and var[idx]:
                cands.append(var[idx])
            if want_b and want_i and var[0]:
                cands.append(var[0])            # 只有粗體檔：粗體用真字型，斜體模擬
        for a in ("-Regular", "Regular", "-Light"):
            if a in base:
                suf = "-BoldItalic" if (want_b and want_i) else "-Bold" if want_b else "-Italic"
                cands.append(base.replace(a, suf))
                if want_b and want_i:
                    cands.append(base.replace(a, "-Bold"))
                break
        for c in cands:
            p2 = os.path.join(d, c)
            if not os.path.exists(p2):
                p2 = next((x for x in (os.path.join(d, n) for n in os.listdir(d) if n.lower() == c.lower())), None) if os.path.isdir(d) else None
            if p2 and os.path.exists(p2):
                f2 = load_font_file(p2)
                if f2:
                    fi2 = {"kind": fi.get("kind", "system"), "label": fi["label"].split("\n")[0] + f" → {os.path.basename(p2)}",
                           "insert": {"fontname": "S" + str(abs(hash(p2)) % 100000), "fontfile": p2}, "font": f2}
                    b2, i2 = bool(getattr(f2, "is_bold", False)) or "bd" in c.lower() or "bold" in c.lower(), \
                        bool(getattr(f2, "is_italic", False)) or "italic" in c.lower() or c.lower().endswith(("i.ttf", "z.ttf"))
                    return fi2, bold and not b2, italic and not i2
    return fi, need_b, need_i


def insert_styled_text(page, origin, text, fi, size, color, bold=False, italic=False, underline=False):
    """以粗體 / 斜體 / 底線寫入文字（盡量用真字型，否則模擬），回傳 (字型資訊, 是否模擬粗, 是否模擬斜)"""
    fi2, sim_b, sim_i = styled_font(fi, bold, italic)
    ins = dict(fi2["insert"])
    if "fontbuffer" in ins:   # insert_text 不接受 fontbuffer，需先註冊字型到頁面
        page.insert_font(fontname=ins["fontname"], fontbuffer=ins.pop("fontbuffer"))
    kw = {}
    if sim_b:     # 模擬粗體：填色 + 同色描邊
        kw.update(render_mode=2, fill=color, border_width=0.035)
    if sim_i:     # 模擬斜體：以插入點為中心水平錯切
        kw["morph"] = (fitz.Point(origin), fitz.Matrix(1, 0, 0.21, 1, 0, 0))
    page.insert_text(origin, text, fontsize=size, color=color, **ins, **kw)
    if underline:
        w = fi2["font"].text_length(text, fontsize=size)
        y = origin.y + size * 0.12
        page.draw_line(fitz.Point(origin.x, y), fitz.Point(origin.x + w, y), color=color, width=max(0.5, size * 0.06))
    return fi2, sim_b, sim_i


def replace_span_text(page, span, new_text, fontfile=None, shrink=False, size=None, color=None,
                      bold=None, italic=None, underline=False):
    """
    取代既有文字（原生文字或 OCR 文字），盡量保持原字型、大小、顏色、基線：
      原生文字：塗銷原字（保留背景圖片與線條）→ choose_font() 挑字型 → 原基線插入
      OCR 文字：以取樣背景色蓋掉影像上的原字 → 移除隱形文字 → 以取樣文字色插入新字
      fontfile：使用者指定的字型檔路徑；"original" = 沿用原嵌入字型；None = 自動
      size    ：使用者指定字級（None = 沿用原大小）
      shrink=True：新文字比原範圍寬時自動縮小字級（最多縮到 50%）；預設關閉，
                   只有在文字會超出頁面右緣時才自動縮小
      bold / italic：None = 沿用原樣式；True / False = 指定；underline：加底線
    回傳使用的字型資訊 dict。
    """
    user_size = size
    bbox = fitz.Rect(span["bbox"])
    st = span_style(page, span)
    # 先決定字型（塗銷後原字型可能已從頁面資源移除，必須在塗銷前擷取）
    fi = None
    if new_text:
        if fontfile and fontfile != "original" and os.path.exists(fontfile):
            fi = font_info_from_path(fontfile)            # 使用者指定字型
        elif st["ocr"]:
            if OCR_DEFAULT_FONT and os.path.exists(OCR_DEFAULT_FONT):
                fi = font_info_from_path(OCR_DEFAULT_FONT)
            else:
                fake = {"font": "PMingLiU" if has_cjk(new_text) else "Times New Roman", "flags": 4}
                fi = choose_font(page, fake, new_text)
        else:
            fi = choose_font(page, span, new_text)
    fit = ocr_fit(page, span, fi, st) if (st["ocr"] and fi) else None
    # 塗銷範圍略縮小，避免誤刪上下行的字
    red = fitz.Rect(bbox.x0, bbox.y0 + bbox.height * 0.15, bbox.x1, bbox.y1 - bbox.height * 0.15)
    page.add_redact_annot(red, fill=False)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                          graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    if st["ocr"]:
        page.draw_rect(bbox + (-1, -1, 1, 1), color=None, fill=st["bg"], overlay=True)
    if not new_text:
        return None
    if st["ocr"]:
        size, origin = fit
        if user_size:   # 使用者指定字級：維持字跡底緣對齊，依比例調整原點
            m = font_ink_metrics(fi, span["text"].strip() or "字")
            ink = st.get("ink")
            if m and ink:
                origin = fitz.Point(ink.x0 - m[0] * user_size / 100, ink.y1 - m[3] * user_size / 100)
            size = user_size
    else:
        size = user_size or span.get("size", 11)       # 原生文字：完全沿用原字級
        origin = fitz.Point(span.get("origin", (bbox.x0, bbox.y1 - size * 0.2)))
    width = fi["font"].text_length(new_text, fontsize=size)
    if shrink and width > bbox.width * 1.02 and width > 0:
        size = max(size * 0.5, size * bbox.width / width)
    elif width > 0 and origin.x + width > page.rect.width - 2:   # 會超出頁面才縮小
        size = max(size * 0.5, size * (page.rect.width - 2 - origin.x) / width)
    ob, oi = (False, False) if st["ocr"] else span_style_flags(span)
    bold = ob if bold is None else bold
    italic = oi if italic is None else italic
    fi2, sim_b, sim_i = insert_styled_text(page, origin, new_text, fi, size, color or st["color"],
                                           bold=bold, italic=italic, underline=underline)
    fi2 = dict(fi2)
    fi2["size"] = size
    fi2["sim"] = (sim_b, sim_i)
    return fi2


# =====================================================================
# OCR（RapidOCR + ONNX Runtime，離線執行；繁體中文辨識模型）
# =====================================================================

_OCR_ENGINES = {}


def get_ocr_engine(kind="default"):
    """
    建立（並快取）OCR 引擎：
      default：RapidOCR 內建模型（文字偵測 + 中英辨識，英文與空白較準）
      cht    ：內附的繁體中文辨識模型（中文行改用它，輸出繁體字）
    """
    if kind not in _OCR_ENGINES:
        from rapidocr_onnxruntime import RapidOCR
        if kind == "cht":
            rec = resource_path(os.path.join("assets", "ocr", "chinese_cht_rec.onnx"))
            keys = resource_path(os.path.join("assets", "ocr", "chinese_cht_dict.txt"))
            if not (os.path.exists(rec) and os.path.exists(keys)):
                return get_ocr_engine("default")
            _OCR_ENGINES[kind] = RapidOCR(rec_model_path=rec, rec_keys_path=keys)
        else:
            _OCR_ENGINES[kind] = RapidOCR()
    return _OCR_ENGINES[kind]


def ocr_page_results(page, dpi=200):
    """
    把頁面轉成影像做 OCR，回傳 [(fitz.Rect PDF座標, 文字, 信心值)]
    流程：內建模型偵測文字行並辨識 → 含中文的行再用繁體中文模型辨識一次（輸出繁體）
    """
    import numpy as np
    pix = page.get_pixmap(dpi=dpi, annots=False)
    img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)
    img = img[:, :, :3][:, :, ::-1].copy() if pix.n >= 3 else np.repeat(img[:, :, :1], 3, axis=2)
    res, _ = get_ocr_engine("default")(img)
    cht = get_ocr_engine("cht")
    out = []
    scale = 72 / dpi
    mat = page.derotation_matrix if page.rotation else fitz.Identity
    for box, text, score in res or []:
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        x0, y0 = max(0, int(min(xs)) - 2), max(0, int(min(ys)) - 2)
        x1, y1 = min(pix.width, int(max(xs)) + 2), min(pix.height, int(max(ys)) + 2)
        if has_cjk(text) and cht is not get_ocr_engine("default") and x1 > x0 and y1 > y0:
            try:
                # 只做「辨識」：把這一行的影像直接交給繁中模型
                rec_fn = getattr(cht, "text_recognizer", None) or cht.text_rec   # 1.2.x / 1.4.x 名稱不同
                r2 = rec_fn([img[y0:y1, x0:x1].copy()])[0]
                if r2 and str(r2[0][0]).strip():
                    text, score = r2[0][0], r2[0][1]
            except Exception:
                pass
        r = fitz.Rect(min(xs) * scale, min(ys) * scale, max(xs) * scale, max(ys) * scale) * mat
        if str(text).strip() and float(score) >= 0.3:   # 過濾信心太低的雜訊
            out.append((r, str(text).strip(), float(score)))
    return out


def add_ocr_text_layer(page, results):
    """在頁面加入「隱形文字層」（render_mode=3），可選取、可搜尋，外觀不變"""
    f = fitz.Font("cjk")
    page.insert_font(fontname="OCRF", fontbuffer=f.buffer)
    for r, text, _ in results:
        size = max(4, r.height * 0.85)
        w = f.text_length(text, fontsize=size)
        origin = fitz.Point(r.x0, r.y1 - r.height * 0.18)
        sx = r.width / w if w > 0 else 1
        page.insert_text(origin, text, fontsize=size, fontname="OCRF", render_mode=3,
                         morph=(origin, fitz.Matrix(sx, 1)))


def ocr_document_bytes(data, pages, dpi=200, progress=None):
    """
    背景執行緒用：開啟文件位元組 → 對指定頁 OCR 並加隱形文字層
    → 回傳 (新位元組, 辨識行數, 已辨識頁碼清單, 略過頁碼清單)
    已有可選取文字的頁面會略過（避免重複疊加）。
    """
    doc = fitz.open("pdf", data)
    total = 0
    done_pages, skipped = [], []
    for k, i in enumerate(pages):
        page = doc[i]
        if not page.get_text("text").strip():
            res = ocr_page_results(page, dpi)
            add_ocr_text_layer(page, res)
            total += len(res)
            done_pages.append(i)
        else:
            skipped.append(i)            # 已有文字的頁面略過
        if progress:
            progress(int((k + 1) * 100 / len(pages)))
    out = doc.tobytes(garbage=3, deflate=True)
    doc.close()
    return out, total, done_pages, skipped


def make_scanned_test_pdf(path, lines=("長庚大學 PDF 測試", "Hello OCR 12345", "繁體中文辨識")):
    """產生「只有影像、沒有文字」的測試 PDF（模擬掃描檔），供自我測試使用"""
    src = fitz.open()
    p = src.new_page(width=595, height=842)
    f = fitz.Font("cjk")
    p.insert_font(fontname="T", fontbuffer=f.buffer)
    for i, t in enumerate(lines):
        p.insert_text((60, 120 + i * 60), t, fontsize=26, fontname="T")
    pix = p.get_pixmap(dpi=150)
    out = fitz.open()
    op = out.new_page(width=595, height=842)
    op.insert_image(op.rect, stream=pix.tobytes("png"))
    out.save(path)
    return path


def ocr_selftest(out_txt=None):
    """命令列自我測試：python pdf_tool.py --selftest-ocr [輸出.txt]（exe 也適用）"""
    import tempfile
    lines = []
    try:
        pdf = make_scanned_test_pdf(os.path.join(tempfile.gettempdir(), "cgu_ocr_selftest.pdf"))
        data, n, *_ = ocr_document_bytes(open(pdf, "rb").read(), [0])
        text = fitz.open("pdf", data)[0].get_text()
        lines.append(f"OCR lines: {n}")
        lines.append(text)
        ok = "Hello" in text and "長庚" in text
        lines.append("SELFTEST " + ("PASS" if ok else "FAIL"))
    except Exception:
        lines.append("SELFTEST ERROR\n" + traceback.format_exc())
        ok = False
    msg = "\n".join(lines)
    if out_txt:
        with open(out_txt, "w", encoding="utf-8") as fh:
            fh.write(msg)
    else:
        print(msg)
    return ok


# =====================================================================
# 背景執行緒：避免轉檔 / 合併等耗時工作讓介面卡死
# =====================================================================

class Worker(QThread):
    """
    通用背景工作執行緒。
    傳入要執行的函式 func(*args, progress=回呼)，
    透過 signal 回報進度、結果或錯誤（signal 會自動排入主執行緒，安全更新 UI）。
    """
    progress = pyqtSignal(int)
    result = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func, self.args, self.kwargs = func, args, kwargs

    def run(self):
        try:
            res = self.func(*self.args, progress=self.progress.emit, **self.kwargs)
            self.result.emit(res)
        except Exception as e:
            traceback.print_exc()
            self.error.emit(str(e))


# =====================================================================
# 可拖曳檔案的元件
# =====================================================================

def local_files_from_event(event, exts):
    """從拖放事件取出符合副檔名的本機檔案路徑"""
    paths = []
    if event.mimeData().hasUrls():
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if p and p.lower().endswith(exts):
                paths.append(p)
    return paths


class DropListWidget(QListWidget):
    """支援從檔案總管拖入檔案的清單；filesDropped 送出檔案路徑清單"""
    filesDropped = pyqtSignal(list)

    def __init__(self, exts, parent=None):
        super().__init__(parent)
        self.exts = exts
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            super().dragEnterEvent(e)

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            super().dragMoveEvent(e)

    def dropEvent(self, e):
        if e.mimeData().hasUrls():
            paths = local_files_from_event(e, self.exts)
            if paths:
                self.filesDropped.emit(paths)
            e.acceptProposedAction()
        else:
            super().dropEvent(e)


class FileDropMixin:
    """讓整個分頁可接受拖入檔案；子類別實作 on_files_dropped(paths)"""
    drop_exts = PDF_EXTS

    def dragEnterEvent(self, e):
        if local_files_from_event(e, self.drop_exts):
            e.acceptProposedAction()

    def dropEvent(self, e):
        paths = local_files_from_event(e, self.drop_exts)
        if paths:
            self.on_files_dropped(paths)


# =====================================================================
# 各分頁的共用基底：進度條 + 執行背景工作
# =====================================================================

class BaseTab(QWidget):
    def __init__(self, main):
        super().__init__()
        self.main = main          # 主視窗，用來更新狀態列
        self.worker = None
        self.progress = QProgressBar()
        self.progress.setValue(0)

    def status(self, msg, ms=5000):
        self.main.statusBar().showMessage(msg, ms)

    def error(self, msg):
        QMessageBox.critical(self, "錯誤", msg)
        self.status("發生錯誤：" + msg.splitlines()[0])

    def run_worker(self, func, *args, on_done=None, busy_widgets=(), **kwargs):
        """啟動背景工作；完成前停用相關按鈕，避免重複點擊"""
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "請稍候", "目前有工作正在進行中")
            return
        for w in busy_widgets:
            w.setEnabled(False)
        self.progress.setValue(0)
        self.worker = Worker(func, *args, **kwargs)
        self.worker.progress.connect(self.progress.setValue)

        def finish():
            for w in busy_widgets:
                w.setEnabled(True)
        self.worker.result.connect(lambda r: (finish(), on_done and on_done(r)))
        self.worker.error.connect(lambda m: (finish(), self.error(m)))
        self.worker.start()
        self.status("處理中…", 0)


# =====================================================================
# 分頁 1：格式轉換
# =====================================================================

class PdfToWordPanel(QWidget):
    """格式轉換 →「PDF → Word」：批次清單、輸出資料夾、掃描檔 OCR、轉換引擎"""

    def __init__(self, tab):
        super().__init__()
        self.tab = tab
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 4, 0, 0)
        lay.addWidget(QLabel("將 PDF 拖曳到下方清單（或按「加入 PDF」），轉成可編輯的 Word 文件："))
        self.list = DropListWidget(PDF_EXTS)
        self.list.filesDropped.connect(self.add_files)
        lay.addWidget(self.list)
        row = QHBoxLayout()
        self.btn_add = bs(QPushButton("加入 PDF…"), "outline-primary")
        self.btn_remove = bs(QPushButton("移除選取"), "outline-danger")
        self.btn_clear = bs(QPushButton("清空"), "outline-danger")
        for b in (self.btn_add, self.btn_remove, self.btn_clear):
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        form = QFormLayout()
        orow = QHBoxLayout()
        self.chk_same = QCheckBox("與原 PDF 相同資料夾")
        self.chk_same.setChecked(True)
        self.out_dir = QLineEdit(os.path.expanduser("~"))
        self.out_dir.setEnabled(False)
        self.btn_browse = bs(QPushButton("瀏覽…"), "outline-secondary")
        self.btn_browse.setEnabled(False)
        orow.addWidget(self.chk_same)
        orow.addWidget(self.out_dir, 1)
        orow.addWidget(self.btn_browse)
        form.addRow("輸出資料夾：", orow)
        self.cmb_engine = QComboBox()
        has_word = word_available()
        self.cmb_engine.addItem("自動（有 Word 用 Word，否則 pdf2docx）", "auto")
        self.cmb_engine.addItem("Microsoft Word（版面較完整）" + ("" if has_word else "（未偵測到）"), "word")
        self.cmb_engine.addItem("pdf2docx（不需 Word）", "pdf2docx")
        form.addRow("轉換引擎：", self.cmb_engine)
        self.chk_ocr = QCheckBox("掃描檔先 OCR（沒有文字層的頁面先辨識，輸出可編輯文字；版面可能與原檔不同）")
        self.chk_ocr.setChecked(True)
        form.addRow("", self.chk_ocr)
        lay.addLayout(form)
        self.btn_go = bs(QPushButton("開始轉換為 Word"), "primary")
        lay.addWidget(self.btn_go, 0, Qt.AlignmentFlag.AlignRight)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        lay.addWidget(self.progress)
        self.btn_add.clicked.connect(self.pick_files)
        self.btn_remove.clicked.connect(lambda: [self.list.takeItem(self.list.row(i)) for i in self.list.selectedItems()])
        self.btn_clear.clicked.connect(self.list.clear)
        self.chk_same.toggled.connect(lambda on: (self.out_dir.setEnabled(not on), self.btn_browse.setEnabled(not on)))
        self.btn_browse.clicked.connect(self.pick_dir)
        self.btn_go.clicked.connect(self.start)
        self.worker = None
        self.last_result = None

    def files(self):
        return [self.list.item(i).text() for i in range(self.list.count())]

    def add_files(self, paths):
        existing = set(self.files())
        new = [p for p in paths if p not in existing and p.lower().endswith(PDF_EXTS)]
        for p in new:
            self.list.addItem(p)
        self.tab.status(f"已加入 {len(new)} 個 PDF")

    def pick_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "選擇 PDF", "", "PDF (*.pdf)")
        if paths:
            self.add_files(paths)

    def pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, "選擇輸出資料夾", self.out_dir.text())
        if d:
            self.out_dir.setText(d)

    def start(self):
        files = self.files()
        if not files:
            return self.tab.error("請先加入要轉換的 PDF")
        out_dir = None if self.chk_same.isChecked() else self.out_dir.text().strip()
        if out_dir is not None and not os.path.isdir(out_dir):
            return self.tab.error("輸出資料夾不存在")
        if self.worker and self.worker.isRunning():
            return
        self.btn_go.setEnabled(False)
        self.progress.setValue(0)
        self.worker = Worker(run_pdf_to_word, files, out_dir, self.chk_ocr.isChecked(), self.cmb_engine.currentData())
        self.worker.progress.connect(self.progress.setValue)
        self.worker.result.connect(self.done)
        self.worker.error.connect(lambda m: (self.btn_go.setEnabled(True), self.tab.error(m)))
        self.worker.start()
        self.tab.status("PDF → Word 轉換中…", 0)

    def done(self, res, show=True):
        self.btn_go.setEnabled(True)
        self.progress.setValue(100)
        self.last_result = res
        outs, errs = res["outputs"], res["errors"]
        lines = [f"成功轉換 {len(outs)} 個檔案" + (f"，失敗 {len(errs)} 個" if errs else "") + "。"]
        for o, eng, ocr in outs:
            lines.append(f"✓ {os.path.basename(o)}（{eng}{'，已先 OCR' if ocr else ''}）")
        if errs:
            lines.append("")
            lines += ["✗ " + e for e in errs]
        if any(o[2] for o in outs):
            lines.append("\n※ 掃描頁已先 OCR 轉為文字，版面可能與原檔不同，請檢查辨識結果。")
        self.tab.status(lines[0])
        if not show:
            return
        box = QMessageBox(QMessageBox.Icon.Warning if errs else QMessageBox.Icon.Information,
                          "PDF → Word 完成", "\n".join(lines), parent=self)
        btn_open = box.addButton("開啟資料夾", QMessageBox.ButtonRole.ActionRole) if outs else None
        box.addButton("確定", QMessageBox.ButtonRole.AcceptRole)
        box.exec()
        if btn_open is not None and box.clickedButton() is btn_open:
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(outs[0][0])))


class ConvertTab(BaseTab):
    def __init__(self, main):
        super().__init__(main)
        root = QVBoxLayout(self)
        # ---- 模式切換：Word/圖片 → PDF ｜ PDF → Word ----
        mrow = QHBoxLayout()
        mrow.addWidget(QLabel("轉換方式："))
        self.rb_to_pdf = QRadioButton("Word / 圖片 → PDF")
        self.rb_to_word = QRadioButton("PDF → Word（.docx）")
        self.rb_to_pdf.setChecked(True)
        self.mode_group = QButtonGroup(self)
        for i_, rb in enumerate((self.rb_to_pdf, self.rb_to_word)):
            self.mode_group.addButton(rb, i_)
            mrow.addWidget(rb)
        mrow.addStretch()
        root.addLayout(mrow)
        self.mode_stack = QStackedWidget()
        root.addWidget(self.mode_stack)
        page0 = QWidget()
        lay = QVBoxLayout(page0)
        lay.setContentsMargins(0, 4, 0, 0)
        self.mode_stack.addWidget(page0)
        self.word_page = PdfToWordPanel(self)
        self.mode_stack.addWidget(self.word_page)
        self.mode_group.idClicked.connect(self.mode_stack.setCurrentIndex)
        lay.addWidget(QLabel("將 .docx / .jpg / .jpeg / .png 檔案拖曳到下方清單（或按「加入檔案」）："))
        self.list = DropListWidget(IMAGE_EXTS + WORD_EXTS)
        self.list.filesDropped.connect(self.add_files)
        lay.addWidget(self.list)

        row = QHBoxLayout()
        self.btn_add = QPushButton("加入檔案…")
        self.btn_remove = QPushButton("移除選取")
        self.btn_clear = QPushButton("清空")
        for b in (self.btn_add, self.btn_remove, self.btn_clear):
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)

        out_row = QHBoxLayout()
        out_row.addWidget(QLabel("輸出資料夾："))
        self.chk_same = QCheckBox("與原檔相同資料夾")
        self.chk_same.setChecked(True)
        self.out_dir = QLineEdit(os.path.expanduser("~"))
        self.out_dir.setEnabled(False)
        self.btn_browse = QPushButton("瀏覽…")
        self.btn_browse.setEnabled(False)
        out_row.addWidget(self.chk_same)
        out_row.addWidget(self.out_dir, 1)
        out_row.addWidget(self.btn_browse)
        lay.addLayout(out_row)

        self.chk_combine = QCheckBox("把所有圖片合併成一個 PDF（依清單順序）")
        lay.addWidget(self.chk_combine)
        if word_available():
            _engine = "將使用 Microsoft Word"
        elif find_soffice():
            _engine = "未偵測到 Word，將使用 LibreOffice"
        else:
            _engine = (f"未偵測到 Word 或 LibreOffice，請先安裝 "
                       f"<a href='{LIBREOFFICE_URL}'>免費的 LibreOffice</a>")
        _hint = QLabel(f"※ Word 轉檔：{_engine}。")
        _hint.setOpenExternalLinks(True)
        lay.addWidget(_hint)

        self.btn_go = QPushButton("開始轉換")
        lay.addWidget(self.btn_go, 0, Qt.AlignmentFlag.AlignRight)
        lay.addWidget(self.progress)

        # 事件綁定
        self.btn_add.clicked.connect(self.pick_files)
        self.btn_remove.clicked.connect(lambda: [self.list.takeItem(self.list.row(i))
                                                 for i in self.list.selectedItems()])
        self.btn_clear.clicked.connect(self.list.clear)
        self.chk_same.toggled.connect(lambda on: (self.out_dir.setEnabled(not on), self.btn_browse.setEnabled(not on)))
        self.btn_browse.clicked.connect(self.pick_dir)
        self.btn_go.clicked.connect(self.start)

    def add_files(self, paths):
        existing = set(self.files())
        for p in paths:
            if p not in existing:
                self.list.addItem(p)
        self.status(f"已加入 {len(paths)} 個檔案")

    def files(self):
        return [self.list.item(i).text() for i in range(self.list.count())]

    def pick_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "選擇檔案", "",
                                                "支援的檔案 (*.docx *.jpg *.jpeg *.png)")
        if paths:
            self.add_files(paths)

    def pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, "選擇輸出資料夾", self.out_dir.text())
        if d:
            self.out_dir.setText(d)

    def start(self):
        files = self.files()
        if not files:
            return self.error("請先加入要轉換的檔案")
        out_dir = None if self.chk_same.isChecked() else self.out_dir.text().strip()
        if out_dir is not None and not os.path.isdir(out_dir):
            return self.error("輸出資料夾不存在")
        self.run_worker(run_conversion, files, out_dir, self.chk_combine.isChecked(),
                        on_done=self.done, busy_widgets=[self.btn_go])

    def done(self, res):
        outputs, errors = res
        msg = f"成功產生 {len(outputs)} 個 PDF。"
        if errors:
            msg += "\n\n以下檔案失敗：\n" + "\n".join(errors)
        elif outputs:
            msg += "\n\n" + "\n".join(outputs)
        box = QMessageBox(QMessageBox.Icon.Warning if errors else QMessageBox.Icon.Information,
                          "轉換完成（部分失敗）" if errors else "轉換完成", msg, parent=self)
        btn_open = box.addButton("開啟資料夾", QMessageBox.ButtonRole.ActionRole) if outputs else None
        box.addButton("確定", QMessageBox.ButtonRole.AcceptRole)
        box.exec()
        if btn_open is not None and box.clickedButton() is btn_open:
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(outputs[0])))
        self.status(msg.splitlines()[0])


# =====================================================================
# 分頁 2：合併檔案
# =====================================================================

class MergeTab(BaseTab):
    def __init__(self, main):
        super().__init__(main)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("將多個 PDF 拖曳到清單中，用右側按鈕調整順序，再按「合併」："))
        tip = QLabel("💡 想一邊合併、一邊調整頁面順序、旋轉或刪頁？請改用左側「組織頁面」→「插入檔案」，一個畫面就能完成。")
        tip.setObjectName("PanelHint")
        tip.setWordWrap(True)
        lay.addWidget(tip)
        row = QHBoxLayout()
        self.list = DropListWidget(PDF_EXTS)
        self.list.filesDropped.connect(self.add_files)
        row.addWidget(self.list)
        col = QVBoxLayout()
        self.btn_add = QPushButton("加入 PDF…")
        self.btn_up = QPushButton("▲ 上移")
        self.btn_down = QPushButton("▼ 下移")
        self.btn_remove = QPushButton("移除")
        self.btn_clear = QPushButton("清空")
        for b in (self.btn_add, self.btn_up, self.btn_down, self.btn_remove, self.btn_clear):
            col.addWidget(b)
        col.addStretch()
        row.addLayout(col)
        lay.addLayout(row)
        self.btn_go = QPushButton("合併並另存…")
        lay.addWidget(self.btn_go, 0, Qt.AlignmentFlag.AlignRight)
        lay.addWidget(self.progress)

        self.btn_add.clicked.connect(self.pick_files)
        self.btn_up.clicked.connect(lambda: self.move(-1))
        self.btn_down.clicked.connect(lambda: self.move(1))
        self.btn_remove.clicked.connect(lambda: [self.list.takeItem(self.list.row(i))
                                                 for i in self.list.selectedItems()])
        self.btn_clear.clicked.connect(self.list.clear)
        self.btn_go.clicked.connect(self.start)

    def add_files(self, paths):
        for p in paths:
            self.list.addItem(p)
        self.status(f"已加入 {len(paths)} 個 PDF")

    def files(self):
        return [self.list.item(i).text() for i in range(self.list.count())]

    def pick_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "選擇 PDF", "", "PDF (*.pdf)")
        if paths:
            self.add_files(paths)

    def move(self, delta):
        """把目前選取的項目上移 / 下移一格"""
        row = self.list.currentRow()
        new = row + delta
        if row < 0 or not (0 <= new < self.list.count()):
            return
        item = self.list.takeItem(row)
        self.list.insertItem(new, item)
        self.list.setCurrentRow(new)

    def start(self):
        files = self.files()
        if len(files) < 2:
            return self.error("請至少加入 2 個 PDF")
        out, _ = QFileDialog.getSaveFileName(self, "儲存合併結果", "合併.pdf", "PDF (*.pdf)")
        if not out:
            return
        self.run_worker(merge_pdfs, files, out, busy_widgets=[self.btn_go],
                        on_done=lambda p: (QMessageBox.information(self, "完成", f"已輸出：\n{p}"),
                                           self.status("合併完成：" + p)))


# =====================================================================
# 分頁 3：組織頁面
# =====================================================================

class PageGrid(QListWidget):
    """組織頁面的縮圖網格：可拖曳排序，也可從檔案總管把 PDF / 圖片拖進來插入到放開的位置"""
    filesDropped = pyqtSignal(list, int)          # (檔案路徑, 插入位置)

    def _ext_files(self, e):
        return local_files_from_event(e, PDF_EXTS + IMAGE_EXTS)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            if self._ext_files(e):
                e.setDropAction(Qt.DropAction.CopyAction)
                e.accept()
            else:
                e.ignore()
        else:
            super().dragEnterEvent(e)

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            e.setDropAction(Qt.DropAction.CopyAction)
            e.accept()
        else:
            super().dragMoveEvent(e)

    def dropEvent(self, e):
        if e.mimeData().hasUrls():
            paths = self._ext_files(e)
            if paths:
                idx = self.indexAt(e.position().toPoint())
                row = idx.row() if idx.isValid() else self.count()
                if idx.isValid():                       # 放在縮圖右半邊 → 插在它後面
                    r = self.visualRect(idx)
                    if e.position().x() > r.center().x():
                        row += 1
                self.filesDropped.emit(paths, row)
            e.accept()
        else:
            super().dropEvent(e)


class OrganizeTab(FileDropMixin, BaseTab):
    """
    組織頁面（類似 Acrobat）：開啟一或多個 PDF / 圖片，在同一個畫面插入其他檔案、
    拖曳排序、旋轉、刪除、插入空白頁，最後匯出成一個新 PDF。
    """
    ROLE_INDEX = Qt.ItemDataRole.UserRole        # 原始頁碼（空白頁為 (寬, 高)）
    ROLE_ROT = Qt.ItemDataRole.UserRole + 1      # 額外旋轉角度
    ROLE_PIX = Qt.ItemDataRole.UserRole + 2      # 原始縮圖（旋轉時用）
    ROLE_SRC = Qt.ItemDataRole.UserRole + 3      # 來源檔路徑（空白頁為 None）
    drop_exts = PDF_EXTS + IMAGE_EXTS
    SRC_COLORS = ["#E8780F", "#1E7FD8", "#2E9E5B", "#7B4FD6", "#C0398E", "#0F9AA8", "#B7282E", "#6B6B6B"]

    def __init__(self, main):
        super().__init__(main)
        self.setAcceptDrops(True)
        self.src_path = None            # 第一個來源檔（相容舊介面）
        self.sources = []               # 來源檔路徑，依加入順序（決定 A、B、C 代號與顏色）
        lay = QVBoxLayout(self)
        top = QHBoxLayout()
        self.btn_open = QPushButton("開啟 PDF…")
        self.btn_insert = QPushButton("＋ 插入檔案…")
        self.btn_insert.setToolTip("把其他 PDF 或圖片插入到選取頁之後（沒選取時加在最後）")
        self.btn_blank = QPushButton("＋ 空白頁")
        self.btn_blank.setToolTip("在選取頁之後插入一張空白頁")
        self.lbl = QLabel("尚未載入（可直接把 PDF 或圖片拖進來，一次拖多個就是合併）")
        self.lbl.setWordWrap(True)
        top.addWidget(self.btn_open)
        top.addWidget(self.btn_insert)
        top.addWidget(self.btn_blank)
        top.addWidget(self.lbl, 1)
        lay.addLayout(top)

        # 縮圖網格：IconMode、由左到右自動換行，搭配 InternalMove 可用滑鼠拖曳排序
        self.grid = PageGrid()
        self.grid.setViewMode(QListView.ViewMode.IconMode)
        self.grid.setFlow(QListView.Flow.LeftToRight)
        self.grid.setWrapping(True)
        self.grid.setResizeMode(QListView.ResizeMode.Adjust)
        self.grid.setMovement(QListView.Movement.Snap)
        self.grid.setIconSize(QSize(140, 180))
        self.grid.setGridSize(QSize(165, 230))
        self.grid.setSpacing(6)
        self.grid.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.grid.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.grid.setDefaultDropAction(Qt.DropAction.MoveAction)
        # IconMode 拖放後項目位置可能錯亂，順序改變後重新排版網格並更新頁碼
        self.grid.model().rowsMoved.connect(lambda *a: self._after_change())
        self.grid.model().rowsInserted.connect(lambda *a: self.grid.doItemsLayout())
        self.grid.filesDropped.connect(self.insert_files)
        lay.addWidget(self.grid, 1)

        row = QHBoxLayout()
        self.btn_left = QPushButton("⟲ 左轉 90°")
        self.btn_right = QPushButton("⟳ 右轉 90°")
        self.btn_del = QPushButton("刪除選取頁")
        self.btn_fwd = QPushButton("◀ 往前移")
        self.btn_back = QPushButton("往後移 ▶")
        self.btn_export = QPushButton("匯出新 PDF…")
        self.btn_export.setProperty("variant", "primary")
        for b in (self.btn_left, self.btn_right, self.btn_del, self.btn_fwd, self.btn_back):
            row.addWidget(b)
        row.addStretch()
        row.addWidget(self.btn_export)
        lay.addLayout(row)
        lay.addWidget(QLabel("提示：Ctrl / Shift 可多選，直接拖曳縮圖調整順序，Delete 鍵刪除。"
                             "從檔案總管把 PDF / 圖片拖到兩頁之間即可插入；縮圖左上角的字母代表來源檔。"))
        lay.addWidget(self.progress)

        self.btn_open.clicked.connect(self.pick)
        self.btn_insert.clicked.connect(self.pick_insert)
        self.btn_blank.clicked.connect(self.insert_blank)
        self.btn_left.clicked.connect(lambda: self.rotate_selected(-90))
        self.btn_right.clicked.connect(lambda: self.rotate_selected(90))
        self.btn_del.clicked.connect(self.delete_selected)
        self.btn_fwd.clicked.connect(lambda: self.move_selected(-1))
        self.btn_back.clicked.connect(lambda: self.move_selected(1))
        self.btn_export.clicked.connect(self.export)
        QShortcut(QKeySequence(Qt.Key.Key_Delete), self.grid, activated=self.delete_selected)

    # ---------------- 開啟 / 插入 ----------------
    def on_files_dropped(self, paths):
        """拖到分頁空白處：沒有內容時開啟，已有內容時加在最後"""
        if self.grid.count() == 0:
            self.load(paths[0])
            if len(paths) > 1:
                self.insert_files(paths[1:], self.grid.count())
        else:
            self.insert_files(paths, self.grid.count())

    def pick(self):
        if self.grid.count() and QMessageBox.question(
                self, "開啟", "要清除目前的頁面，重新開啟檔案嗎？\n（若要加入其他檔案，請改用「插入檔案」）") \
                != QMessageBox.StandardButton.Yes:
            return
        paths, _ = QFileDialog.getOpenFileNames(self, "開啟 PDF（可多選，依選取順序合併）", "",
                                                "PDF 或圖片 (*.pdf *.jpg *.jpeg *.png)")
        if paths:
            self.grid.clear()
            self.sources = []
            self.on_files_dropped(paths)

    def pick_insert(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "插入檔案", os.path.dirname(self.src_path or ""),
                                                "PDF 或圖片 (*.pdf *.jpg *.jpeg *.png)")
        if paths:
            self.insert_files(paths, self._insert_row())

    def _insert_row(self):
        """插入位置：最後一個選取頁之後；沒有選取時加在最後"""
        rows = [self.grid.row(it) for it in self.grid.selectedItems()]
        return max(rows) + 1 if rows else self.grid.count()

    def load(self, path):
        """清空並開啟一個 PDF（或圖片）"""
        self.grid.clear()
        self.sources = []
        self.src_path = None
        self.insert_files([path], 0)

    def _make_item(self, src, idx, page):
        zoom = 160 / max(page.rect.width, page.rect.height)
        pix = fitz_pixmap_to_qpixmap(page.get_pixmap(matrix=fitz.Matrix(zoom, zoom)))
        item = QListWidgetItem()
        item.setData(self.ROLE_SRC, src)
        item.setData(self.ROLE_INDEX, idx)
        item.setData(self.ROLE_ROT, 0)
        item.setData(self.ROLE_PIX, pix)
        item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter)
        return item

    def insert_files(self, paths, row):
        """把 PDF / 圖片的所有頁面插入到指定位置"""
        added = 0
        for path in paths:
            try:
                doc = open_as_pdf(path)
            except Exception as e:
                self.error(f"無法開啟 {os.path.basename(path)}：{e}")
                continue
            if path not in self.sources:
                self.sources.append(path)
            if self.src_path is None:
                self.src_path = path
            for i, page in enumerate(doc):
                self.grid.insertItem(row, self._make_item(path, i, page))
                row += 1
                added += 1
                if i % 10 == 0:
                    QApplication.processEvents()   # 大檔載入時讓介面保持回應
            doc.close()
        self._after_change()
        if added:
            self.status(f"已加入 {added} 頁")

    def insert_blank(self):
        if not self.grid.count():
            return self.status("請先開啟 PDF")
        row = self._insert_row()
        ref = self.grid.item(max(0, row - 1))
        w, h = 595, 842                                  # 預設 A4，與前一頁同尺寸
        try:
            if ref.data(self.ROLE_SRC):
                with open_as_pdf(ref.data(self.ROLE_SRC)) as d:
                    r = d[ref.data(self.ROLE_INDEX)].rect
                    w, h = r.width, r.height
            else:
                w, h = ref.data(self.ROLE_INDEX)
        except Exception:
            pass
        pm = QPixmap(int(160 * w / max(w, h)), int(160 * h / max(w, h)))
        pm.fill(QColor("#FFFFFF"))
        item = QListWidgetItem()
        item.setData(self.ROLE_SRC, None)
        item.setData(self.ROLE_INDEX, (w, h))
        item.setData(self.ROLE_ROT, 0)
        item.setData(self.ROLE_PIX, pm)
        item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.grid.insertItem(row, item)
        self._after_change()
        self.grid.clearSelection()
        item.setSelected(True)
        self.status("已插入空白頁")

    # ---------------- 畫面 ----------------
    def _src_tag(self, src):
        if src is None:
            return "", "#9A9A9A"
        i = self.sources.index(src) if src in self.sources else 0
        return chr(ord("A") + i % 26), self.SRC_COLORS[i % len(self.SRC_COLORS)]

    def _render_icon(self, it):
        pix = it.data(self.ROLE_PIX).transformed(QTransform().rotate(it.data(self.ROLE_ROT)))
        multi = len(self.sources) > 1
        if True:                                         # 頁面外框；多個來源時左上角標示來源字母
            pm = QPixmap(pix)
            p = QPainter(pm)
            p.setRenderHint(QPainter.RenderHint.Antialiasing)
            if it.data(self.ROLE_SRC) is not None:
                p.setPen(QPen(QColor("#B8B8B8"), 1))
                p.setBrush(Qt.BrushStyle.NoBrush)
                p.drawRect(pm.rect().adjusted(0, 0, -1, -1))
            if it.data(self.ROLE_SRC) is None:
                p.setPen(QPen(QColor("#BBBBBB"), 1, Qt.PenStyle.DashLine))
                p.drawRect(pm.rect().adjusted(0, 0, -1, -1))
                p.setPen(QColor("#9A9A9A"))
                p.drawText(pm.rect(), Qt.AlignmentFlag.AlignCenter, "空白頁")
            elif multi:
                tag, color = self._src_tag(it.data(self.ROLE_SRC))
                p.setPen(Qt.PenStyle.NoPen)
                p.setBrush(QColor(color))
                p.drawRoundedRect(QRectF(3, 3, 22, 18), 4, 4)
                p.setPen(QColor("#FFFFFF"))
                f = QFont()
                f.setBold(True)
                f.setPointSize(9)
                p.setFont(f)
                p.drawText(QRectF(3, 3, 22, 18), Qt.AlignmentFlag.AlignCenter, tag)
            p.end()
            pix = pm
        it.setIcon(thumb_icon(pix))

    def _after_change(self):
        """更新每頁的新頁碼標籤、來源標示與上方說明"""
        multi = len(self.sources) > 1
        for i in range(self.grid.count()):
            it = self.grid.item(i)
            src = it.data(self.ROLE_SRC)
            if src is None:
                sub = "空白頁"
            elif multi:
                sub = f"{self._src_tag(src)[0]}・原第 {it.data(self.ROLE_INDEX) + 1} 頁"
            else:
                sub = f"原第 {it.data(self.ROLE_INDEX) + 1} 頁"
            it.setText(f"第 {i + 1} 頁\n{sub}")
            self._render_icon(it)
        self.grid.doItemsLayout()
        if not self.sources:
            self.lbl.setText("尚未載入（可直接把 PDF 或圖片拖進來，一次拖多個就是合併）")
        elif not multi:
            self.lbl.setText(f"{os.path.basename(self.sources[0])}（目前共 {self.grid.count()} 頁）")
        else:
            parts = []
            for s_ in self.sources:
                tag, color = self._src_tag(s_)
                parts.append(f"<span style='background:{color};color:#fff;padding:0 4px;border-radius:3px'>"
                             f"<b>{tag}</b></span> {os.path.basename(s_)}")
            self.lbl.setText("　".join(parts) + f"　｜ 目前共 {self.grid.count()} 頁")

    # ---------------- 編輯 ----------------
    def rotate_selected(self, deg):
        items = self.grid.selectedItems()
        if not items:
            return self.status("請先選取頁面")
        for it in items:
            it.setData(self.ROLE_ROT, (it.data(self.ROLE_ROT) + deg) % 360)
            self._render_icon(it)

    def delete_selected(self):
        items = self.grid.selectedItems()
        if not items:
            return self.status("請先選取頁面")
        for it in items:
            self.grid.takeItem(self.grid.row(it))
        self._after_change()
        self.status(f"已刪除 {len(items)} 頁（匯出後才會生效）")

    def move_selected(self, delta):
        row = self.grid.currentRow()
        new = row + delta
        if row < 0 or not (0 <= new < self.grid.count()):
            return
        it = self.grid.takeItem(row)
        self.grid.insertItem(new, it)
        self._after_change()
        self.grid.setCurrentRow(new)

    def page_specs(self):
        """依目前畫面順序取得 [(原始頁碼, 旋轉角度)]（單一來源用；保留相容）"""
        return [(self.grid.item(i).data(self.ROLE_INDEX), self.grid.item(i).data(self.ROLE_ROT))
                for i in range(self.grid.count())]

    def page_items(self):
        """依目前畫面順序取得 [(來源檔, 原始頁碼, 旋轉角度)]"""
        return [(it.data(self.ROLE_SRC), it.data(self.ROLE_INDEX), it.data(self.ROLE_ROT))
                for it in (self.grid.item(i) for i in range(self.grid.count()))]

    def export(self):
        if not self.grid.count():
            return self.error("請先開啟 PDF")
        base = os.path.splitext(self.src_path or "整理後")[0]
        out, _ = QFileDialog.getSaveFileName(self, "匯出 PDF", unique_path(base + "_整理後.pdf"), "PDF (*.pdf)")
        if not out:
            return
        if not out.lower().endswith(".pdf"):
            out += ".pdf"
        self.run_worker(export_organized, None, self.page_items(), out,
                        busy_widgets=[self.btn_export], on_done=self.export_done)

    def export_done(self, path):
        self.status("匯出完成：" + path)
        box = QMessageBox(QMessageBox.Icon.Information, "完成", f"已輸出：\n{path}", parent=self)
        btn_edit = box.addButton("在「編輯 PDF」開啟", QMessageBox.ButtonRole.ActionRole)
        btn_dir = box.addButton("開啟資料夾", QMessageBox.ButtonRole.ActionRole)
        box.addButton("確定", QMessageBox.ButtonRole.AcceptRole)
        box.exec()
        if box.clickedButton() is btn_edit:
            self.main.open_pdf(path)
        elif box.clickedButton() is btn_dir:
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(path)))


# =====================================================================
# 分頁 4：文字浮水印
# =====================================================================

class WatermarkTab(FileDropMixin, BaseTab):
    def __init__(self, main):
        super().__init__(main)
        self.setAcceptDrops(True)
        self.color = QColor(255, 0, 0)
        lay = QVBoxLayout(self)
        form = QFormLayout()
        frow = QHBoxLayout()
        self.src = QLineEdit()
        self.src.setPlaceholderText("選擇或拖曳 PDF 到此分頁")
        self.btn_src = QPushButton("瀏覽…")
        frow.addWidget(self.src)
        frow.addWidget(self.btn_src)
        form.addRow("來源 PDF：", frow)
        self.text = QLineEdit("機密文件 CONFIDENTIAL")
        form.addRow("浮水印文字：", self.text)
        self.opacity = QSlider(Qt.Orientation.Horizontal)
        self.opacity.setRange(5, 100)
        self.opacity.setValue(30)
        self.lbl_op = QLabel("30%")
        orow = QHBoxLayout()
        orow.addWidget(self.opacity)
        orow.addWidget(self.lbl_op)
        form.addRow("不透明度：", orow)
        self.size = QSpinBox()
        self.size.setRange(8, 300)
        self.size.setValue(60)
        form.addRow("字級 (pt)：", self.size)
        self.angle = QSpinBox()
        self.angle.setRange(-180, 180)
        self.angle.setValue(45)
        form.addRow("角度 (°)：", self.angle)
        self.btn_color = QPushButton()
        self.update_color_btn()
        form.addRow("顏色：", self.btn_color)
        fnt = find_cjk_font()
        form.addRow("中文字型：", QLabel(fnt or "內建 china-t（未找到系統中文字型）"))
        lay.addLayout(form)
        lay.addStretch()
        self.btn_go = QPushButton("加上浮水印並另存…")
        lay.addWidget(self.btn_go, 0, Qt.AlignmentFlag.AlignRight)
        lay.addWidget(self.progress)

        self.btn_src.clicked.connect(self.pick)
        self.opacity.valueChanged.connect(lambda v: self.lbl_op.setText(f"{v}%"))
        self.btn_color.clicked.connect(self.pick_color)
        self.btn_go.clicked.connect(self.start)

    def update_color_btn(self):
        self.btn_color.setText(self.color.name())
        self.btn_color.setStyleSheet(f"background:{self.color.name()};color:white;")

    def pick_color(self):
        c = QColorDialog.getColor(self.color, self, "選擇顏色")
        if c.isValid():
            self.color = c
            self.update_color_btn()

    def on_files_dropped(self, paths):
        self.src.setText(paths[0])

    def pick(self):
        p, _ = QFileDialog.getOpenFileName(self, "選擇 PDF", "", "PDF (*.pdf)")
        if p:
            self.src.setText(p)

    def start(self):
        src = self.src.text().strip()
        if not os.path.isfile(src):
            return self.error("請選擇有效的 PDF 檔")
        if not self.text.text().strip():
            return self.error("請輸入浮水印文字")
        default = os.path.splitext(src)[0] + "_浮水印.pdf"
        out, _ = QFileDialog.getSaveFileName(self, "另存新檔", default, "PDF (*.pdf)")
        if not out:
            return
        self.run_worker(add_text_watermark, src, out, self.text.text(),
                        opacity=self.opacity.value() / 100, fontsize=self.size.value(),
                        angle=self.angle.value(), color=color_to_rgb01(self.color),
                        busy_widgets=[self.btn_go],
                        on_done=lambda p: (QMessageBox.information(self, "完成", f"已輸出：\n{p}"),
                                           self.status("浮水印完成：" + p)))


# =====================================================================
# 外觀：應用程式名稱、配色、自繪圖示、QSS 樣式表
# （版面參考一般 PDF 閱讀器的配置；所有圖示皆由 QPainter 自行繪製，未使用任何商標素材）
# =====================================================================

APP_NAME = "長庚大學PDF簡易編輯工具"
APP_TITLE = f"{APP_NAME} {APP_VERSION}"   # 視窗標題
# ---- 主題色（橘色系）集中管理 ----
THEME = {
    "accent": "#EE8200",        # 主色：標題列底線、選取框、勾選圖示、進度條
    "accent_btn": "#D96A00",    # 主要按鈕 / 版本標籤底色（較深，白字對比較好）
    "accent_dark": "#B35500",   # 按鈕滑過 / 按下、選取文字
    "tint": "#FFF1E0",          # 選取 / 啟用項目的淡橘底
    "tint_border": "#F6C58A",   # 淡橘邊框
    "disabled": "#F2C594",      # 停用的主要按鈕
    "annot": "#E8780F",         # 註解預設顏色
}
ACCENT = THEME["accent"]
ACCENT_BTN = THEME["accent_btn"]
ACCENT_DARK = THEME["accent_dark"]
TINT = THEME["tint"]
TINT_BORDER = THEME["tint_border"]
HEADER_BG = "#2B2B2B"     # 頂部標題列
INK = "#4B4B4B"           # 圖示預設顏色


def logo_pixmap(height=None):
    """從內嵌 base64 取得長庚大學校徽 QPixmap"""
    pm = QPixmap()
    pm.loadFromData(base64.b64decode("".join(APP_LOGO_B64)), "PNG")
    if height:
        pm = pm.scaledToHeight(height, Qt.TransformationMode.SmoothTransformation)
    return pm


def app_window_icon():
    """視窗 / 工作列圖示：校徽是橫式，置中放進正方形避免被壓扁"""
    pm = logo_pixmap()
    side = max(pm.width(), pm.height())
    sq = QPixmap(side, side)
    sq.fill(Qt.GlobalColor.transparent)
    p = QPainter(sq)
    p.drawPixmap((side - pm.width()) // 2, (side - pm.height()) // 2, pm)
    p.end()
    return QIcon(sq)


def _legacy_icon(kind, color=INK, size=24):
    """用 QPainter 畫出簡單的線條圖示（自製，不使用外部圖檔）"""
    pm = QPixmap(size * 2, size * 2)          # 2 倍解析度，高 DPI 螢幕較清晰
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.scale(2 * size / 24, 2 * size / 24)     # 以 24×24 座標系繪製
    pen = QPen(QColor(color), 1.8, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap,
               Qt.PenJoinStyle.RoundJoin)
    p.setPen(pen)
    p.setBrush(Qt.BrushStyle.NoBrush)
    c = QColor(color)

    def page(x=5, y=3, w=14, h=18):
        path = QPainterPath()
        path.moveTo(x, y); path.lineTo(x + w - 4, y); path.lineTo(x + w, y + 4)
        path.lineTo(x + w, y + h); path.lineTo(x, y + h); path.closeSubpath()
        p.drawPath(path)

    if kind == "open":
        path = QPainterPath()
        path.moveTo(3, 6); path.lineTo(9, 6); path.lineTo(11, 8); path.lineTo(21, 8)
        path.lineTo(21, 19); path.lineTo(3, 19); path.closeSubpath()
        p.drawPath(path); p.drawLine(3, 11, 21, 11)
    elif kind == "save":
        p.drawRoundedRect(QRectF(4, 4, 16, 16), 2, 2)
        p.drawRect(QRectF(8, 4, 8, 5)); p.drawRect(QRectF(7, 13, 10, 7))
    elif kind == "saveas":
        p.drawRoundedRect(QRectF(3, 5, 14, 14), 2, 2)
        p.drawRect(QRectF(6, 5, 7, 4))
        p.drawLine(15, 16, 21, 10); p.drawLine(19, 8, 22, 11)
    elif kind in ("undo", "redo"):
        if kind == "redo":
            p.translate(24, 0); p.scale(-1, 1)
        path = QPainterPath()
        path.moveTo(8, 9); path.lineTo(15, 9)
        path.cubicTo(22, 9, 22, 19, 15, 19); path.lineTo(9, 19)
        p.drawPath(path); p.drawLine(8, 9, 12, 5); p.drawLine(8, 9, 12, 13)
    elif kind in ("zoomin", "zoomout"):
        p.drawEllipse(QRectF(4, 4, 12, 12)); p.drawLine(14, 14, 20, 20)
        p.drawLine(7, 10, 13, 10)
        if kind == "zoomin":
            p.drawLine(10, 7, 10, 13)
    elif kind == "fit":
        page(6, 4, 12, 16)
        p.drawLine(2, 12, 5, 12); p.drawLine(19, 12, 22, 12)
        p.drawLine(2, 12, 4, 10); p.drawLine(2, 12, 4, 14)
        p.drawLine(22, 12, 20, 10); p.drawLine(22, 12, 20, 14)
    elif kind in ("prev", "next"):
        if kind == "prev":
            p.drawLine(15, 6, 9, 12); p.drawLine(9, 12, 15, 18)
        else:
            p.drawLine(9, 6, 15, 12); p.drawLine(15, 12, 9, 18)
    elif kind == "thumbs":
        p.drawRect(QRectF(3, 4, 18, 16)); p.drawLine(9, 4, 9, 20)
        for y in (7, 11, 15):
            p.drawLine(5, y, 7, y)
    elif kind == "edit":
        page(4, 3, 13, 18)
        p.setPen(QPen(QColor(ACCENT) if color == INK else c, 2))
        p.drawLine(11, 17, 20, 8); p.drawLine(18, 6, 22, 10)
    elif kind == "merge":
        page(3, 3, 9, 12); page(12, 9, 9, 12)
        p.drawLine(10, 18, 7, 18); p.drawLine(7, 18, 7, 15)
    elif kind == "organize":
        for x, y in ((4, 4), (13, 4), (4, 13), (13, 13)):
            p.drawRoundedRect(QRectF(x, y, 7, 7), 1, 1)
    elif kind == "convert":
        page(3, 3, 10, 13)
        p.drawLine(13, 18, 21, 18); p.drawLine(18, 15, 21, 18); p.drawLine(18, 21, 21, 18)
        p.setFont(QFont("Arial", 5, QFont.Weight.Bold)); p.drawText(QRectF(3, 6, 10, 9), Qt.AlignmentFlag.AlignCenter, "PDF")
    elif kind == "watermark":
        page(4, 3, 16, 18)
        p.setPen(QPen(c, 1.5, Qt.PenStyle.DashLine)); p.drawLine(7, 17, 17, 8)
    elif kind == "select":
        p.setFont(QFont("Times New Roman", 15)); p.drawText(QRectF(2, 1, 16, 20), Qt.AlignmentFlag.AlignCenter, "T")
        p.drawLine(18, 6, 18, 20)
    elif kind == "text":
        p.drawRect(QRectF(3, 4, 18, 16))
        p.setFont(QFont("Arial", 10, QFont.Weight.Bold)); p.drawText(QRectF(3, 4, 18, 16), Qt.AlignmentFlag.AlignCenter, "T+")
    elif kind == "image":
        p.drawRoundedRect(QRectF(3, 5, 18, 14), 2, 2); p.drawEllipse(QRectF(6, 8, 3, 3))
        p.drawPolyline(QPolygonF([QPointF(4, 18), QPointF(10, 12), QPointF(14, 16), QPointF(16, 14), QPointF(20, 18)]))
    elif kind == "highlight":
        p.fillRect(QRectF(3, 15, 18, 5), QColor(255, 210, 0, 200))
        p.drawLine(8, 14, 16, 4); p.drawLine(12, 15, 19, 6)
    elif kind == "rect":
        p.drawRect(QRectF(4, 6, 16, 12))
    elif kind == "line":
        p.drawLine(4, 20, 20, 4)
    elif kind == "pen":
        path = QPainterPath()
        path.moveTo(3, 17); path.cubicTo(7, 7, 11, 21, 14, 12); path.cubicTo(16, 6, 19, 9, 21, 7)
        p.drawPath(path)
    elif kind == "note":
        path = QPainterPath()
        path.moveTo(4, 4); path.lineTo(20, 4); path.lineTo(20, 15); path.lineTo(15, 20)
        path.lineTo(4, 20); path.closeSubpath()
        p.drawPath(path); p.drawLine(8, 9, 16, 9); p.drawLine(8, 13, 13, 13)
    elif kind == "whiteout":
        p.drawRect(QRectF(3, 6, 18, 12))
        p.fillRect(QRectF(6, 9, 12, 6), c)
    elif kind == "ocr":
        for (x1, y1, x2, y2) in ((3, 8, 3, 3), (3, 3, 8, 3), (16, 3, 21, 3), (21, 3, 21, 8),
                                 (3, 16, 3, 21), (3, 21, 8, 21), (21, 16, 21, 21), (16, 21, 21, 21)):
            p.drawLine(x1, y1, x2, y2)
        p.setFont(QFont("Arial", 8, QFont.Weight.Bold)); p.drawText(QRectF(3, 3, 18, 18), Qt.AlignmentFlag.AlignCenter, "文")
    elif kind == "delannot":
        p.drawLine(6, 6, 18, 18); p.drawLine(18, 6, 6, 18)
    p.end()
    return QIcon(pm)


# Lucide 圖示（ISC License, https://lucide.dev ，v1.59.0）以 SVG 內嵌，exe 不需外部檔案
LUCIDE_SVGS = {
    'rotate-ccw': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" /> <path d="M3 3v5h5" /> </svg>',
    'rotate-cw': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8" /> <path d="M21 3v5h-5" /> </svg>',
    'square-check': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <rect width="18" height="18" x="3" y="3" rx="2" /> <path d="m9 12 2 2 4-4" /> </svg>',
    'shield-check': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z" /> <path d="m9 12 2 2 4-4" /> </svg>',
    'signature': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="m21 17-2.156-1.868A.5.5 0 0 0 18 15.5v.5a1 1 0 0 1-1 1h-2a1 1 0 0 1-1-1c0-2.545-3.991-3.97-8.5-4a1 1 0 0 0 0 5c4.153 0 4.745-11.295 5.708-13.5a2.5 2.5 0 1 1 3.31 3.284" /> <path d="M3 21h18" /> </svg>',
    'link': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" /> <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" /> </svg>',
    'printer': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" /> <path d="M6 9V3a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v6" /> <rect x="6" y="14" width="12" height="8" rx="1" /> </svg>',
    'chevron-left': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="m15 18-6-6 6-6" /> </svg>',
    'chevron-right': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="m9 18 6-6-6-6" /> </svg>',
    'eraser': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M21 21H8a2 2 0 0 1-1.42-.587l-3.994-3.999a2 2 0 0 1 0-2.828l10-10a2 2 0 0 1 2.829 0l5.999 6a2 2 0 0 1 0 2.828L12.834 21" /> <path d="m5.082 11.09 8.828 8.828" /> </svg>',
    'file-output': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M4.226 20.925A2 2 0 0 0 6 22h12a2 2 0 0 0 2-2V8a2.4 2.4 0 0 0-.706-1.706l-3.588-3.588A2.4 2.4 0 0 0 14 2H6a2 2 0 0 0-2 2v3.127" /> <path d="M14 2v5a1 1 0 0 0 1 1h5" /> <path d="m5 11-3 3" /> <path d="m5 17-3-3h10" /> </svg>',
    'file-pen-line': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M14.364 13.634a2 2 0 0 0-.506.854l-.837 2.87a.5.5 0 0 0 .62.62l2.87-.837a2 2 0 0 0 .854-.506l4.013-4.009a1 1 0 0 0-3.004-3.004z" /> <path d="M14.487 7.858A1 1 0 0 1 14 7V2" /> <path d="M20 19.645V20a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706l2.516 2.516" /> <path d="M8 18h1" /> </svg>',
    'files': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M15 2h-4a2 2 0 0 0-2 2v11a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V8" /> <path d="M16.706 2.706A2.4 2.4 0 0 0 15 2v5a1 1 0 0 0 1 1h5a2.4 2.4 0 0 0-.706-1.706z" /> <path d="M5 7a2 2 0 0 0-2 2v11a2 2 0 0 0 2 2h8a2 2 0 0 0 1.732-1" /> </svg>',
    'folder-open': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="m6 14 1.5-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.54 6a2 2 0 0 1-1.95 1.5H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 1.67.9H18a2 2 0 0 1 2 2v2" /> </svg>',
    'highlighter': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="m9 11-6 6v3h9l3-3" /> <path d="m22 12-4.6 4.6a2 2 0 0 1-2.8 0l-5.2-5.2a2 2 0 0 1 0-2.8L14 4" /> </svg>',
    'image-plus': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M16 5h6" /> <path d="M19 2v6" /> <path d="M21 11.5V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h7.5" /> <path d="m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21" /> <circle cx="9" cy="9" r="2" /> </svg>',
    'info': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <circle cx="12" cy="12" r="10" /> <path d="M12 16v-4" /> <path d="M12 8h.01" /> </svg>',
    'layout-grid': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <rect width="7" height="7" x="3" y="3" rx="1" /> <rect width="7" height="7" x="14" y="3" rx="1" /> <rect width="7" height="7" x="14" y="14" rx="1" /> <rect width="7" height="7" x="3" y="14" rx="1" /> </svg>',
    'maximize': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M8 3H5a2 2 0 0 0-2 2v3" /> <path d="M21 8V5a2 2 0 0 0-2-2h-3" /> <path d="M3 16v3a2 2 0 0 0 2 2h3" /> <path d="M16 21h3a2 2 0 0 0 2-2v-3" /> </svg>',
    'minus': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M5 12h14" /> </svg>',
    'move-horizontal': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="m18 8 4 4-4 4" /> <path d="M2 12h20" /> <path d="m6 8-4 4 4 4" /> </svg>',
    'panel-left': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <rect width="18" height="18" x="3" y="3" rx="2" /> <path d="M9 3v18" /> </svg>',
    'pencil': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M21.174 6.812a1 1 0 0 0-3.986-3.987L3.842 16.174a2 2 0 0 0-.5.83l-1.321 4.352a.5.5 0 0 0 .623.622l4.353-1.32a2 2 0 0 0 .83-.497z" /> <path d="m15 5 4 4" /> </svg>',
    'redo-2': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="m15 14 5-5-5-5" /> <path d="M20 9H9.5A5.5 5.5 0 0 0 4 14.5A5.5 5.5 0 0 0 9.5 20H13" /> </svg>',
    'save-all': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M10 2v3a1 1 0 0 0 1 1h5" /> <path d="M18 18v-6a1 1 0 0 0-1-1h-6a1 1 0 0 0-1 1v6" /> <path d="M18 22H4a2 2 0 0 1-2-2V6" /> <path d="M8 18a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9.172a2 2 0 0 1 1.414.586l2.828 2.828A2 2 0 0 1 22 6.828V16a2 2 0 0 1-2.01 2z" /> </svg>',
    'save': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M15.2 3a2 2 0 0 1 1.4.6l3.8 3.8a2 2 0 0 1 .6 1.4V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z" /> <path d="M17 21v-7a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v7" /> <path d="M7 3v4a1 1 0 0 0 1 1h7" /> </svg>',
    'scan-text': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M3 7V5a2 2 0 0 1 2-2h2" /> <path d="M17 3h2a2 2 0 0 1 2 2v2" /> <path d="M21 17v2a2 2 0 0 1-2 2h-2" /> <path d="M7 21H5a2 2 0 0 1-2-2v-2" /> <path d="M7 8h8" /> <path d="M7 12h10" /> <path d="M7 16h6" /> </svg>',
    'square': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <rect width="18" height="18" x="3" y="3" rx="2" /> </svg>',
    'stamp': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M14 13V8.5C14 7 15 7 15 5a3 3 0 0 0-6 0c0 2 1 2 1 3.5V13" /> <path d="M20 15.5a2.5 2.5 0 0 0-2.5-2.5h-11A2.5 2.5 0 0 0 4 15.5V17a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1z" /> <path d="M5 22h14" /> </svg>',
    'sticky-note': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M21 9a2.4 2.4 0 0 0-.706-1.706l-3.588-3.588A2.4 2.4 0 0 0 15 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2z" /> <path d="M15 3v5a1 1 0 0 0 1 1h5" /> </svg>',
    'text-cursor': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M17 22h-1a4 4 0 0 1-4-4V6a4 4 0 0 1 4-4h1" /> <path d="M7 22h1a4 4 0 0 0 4-4" /> <path d="M7 2h1a4 4 0 0 1 4 4" /> </svg>',
    'trash-2': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M10 11v6" /> <path d="M14 11v6" /> <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6" /> <path d="M3 6h18" /> <path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" /> </svg>',
    'type': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M12 4v16" /> <path d="M4 7V5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v2" /> <path d="M9 20h6" /> </svg>',
    'undo-2': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <path d="M9 14 4 9l5-5" /> <path d="M4 9h10.5a5.5 5.5 0 0 1 5.5 5.5a5.5 5.5 0 0 1-5.5 5.5H11" /> </svg>',
    'zoom-in': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <circle cx="11" cy="11" r="8" /> <line x1="21" x2="16.65" y1="21" y2="16.65" /> <line x1="11" x2="11" y1="8" y2="14" /> <line x1="8" x2="14" y1="11" y2="11" /> </svg>',
    'zoom-out': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" > <circle cx="11" cy="11" r="8" /> <line x1="21" x2="16.65" y1="21" y2="16.65" /> <line x1="8" x2="14" y1="11" y2="11" /> </svg>',
}

# 功能代號 → Lucide 圖示名稱
ICON_NAMES = {
    "open": "folder-open", "save": "save", "saveas": "save-all", "print": "printer", "dsign": "shield-check", "undo": "undo-2", "redo": "redo-2",
    "zoomin": "zoom-in", "zoomout": "zoom-out", "fit": "move-horizontal", "fitpage": "maximize",
    "rotl": "rotate-ccw", "rotr": "rotate-cw",
    "prev": "chevron-left", "next": "chevron-right", "thumbs": "panel-left",
    "edit": "file-pen-line", "merge": "files", "organize": "layout-grid", "convert": "file-output",
    "watermark": "stamp", "select": "text-cursor", "text": "type", "image": "image-plus",
    "highlight": "highlighter", "rect": "square", "line": "minus", "pen": "pencil",
    "note": "sticky-note", "link": "link", "sign": "signature", "check": "square-check", "whiteout": "eraser", "delannot": "trash-2", "ocr": "scan-text", "about": "info",
}


# 各工具專屬顏色（彩色線條 + 淡色填滿；主工具列維持中性深灰）
TOOL_COLORS = {
    "edit": "#D93A49", "merge": "#7B4FD6", "organize": "#2E9E5B", "convert": "#1E7FD8", "watermark": "#C0398E",
    "select": "#D93A49", "text": "#1E7FD8", "image": "#2E9E5B", "highlight": "#D9A400", "rect": "#7B4FD6",
    "line": "#0F9AA8", "pen": "#C0398E", "note": "#E0A100", "link": "#0B6BCB", "sign": "#B7282E", "check": "#2E9E5B", "whiteout": "#E0603A", "delannot": "#C62828",
    "ocr": "#7B4FD6",
}


def svg_pixmap(name, color, size, stroke=1.75, fill=None):
    """把 Lucide SVG 以指定顏色 / 線寬繪製成 2 倍解析度 QPixmap；fill 給定時加上淡色填滿"""
    svg = LUCIDE_SVGS[name].replace("currentColor", color).replace('stroke-width="2"', f'stroke-width="{stroke}"')
    if fill:
        svg = svg.replace('fill="none"', f'fill="{fill}" fill-opacity="0.16"', 1)
    pm = QPixmap(size * 2, size * 2)
    pm.fill(Qt.GlobalColor.transparent)
    pm.setDevicePixelRatio(2)
    r = QSvgRenderer(QByteArray(svg.encode("utf-8")))
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    r.render(p, QRectF(0, 0, size, size))
    p.end()
    return pm


def make_icon(kind, color=INK, size=24, active=ACCENT, disabled="#B8B8B8", hover=None, tool=False):
    """
    一致風格的圖示（Lucide 線條圖示）：
      一般 = color；滑過 = hover（預設同色）；勾選/啟用中 = 紅色主色；停用 = 淺灰
    找不到對應 SVG 時退回舊的 QPainter 自繪圖示。
    """
    name = ICON_NAMES.get(kind)
    if not name or name not in LUCIDE_SVGS:
        return _legacy_icon(kind, color, size)
    if tool and kind in TOOL_COLORS:     # 彩色工具圖示：啟用中也保留自己的顏色（底色由樣式表標示）
        c = TOOL_COLORS[kind]
        icon = QIcon()
        pm = svg_pixmap(name, c, size, fill=c)
        for m in (QIcon.Mode.Normal, QIcon.Mode.Active, QIcon.Mode.Selected):
            for st in (QIcon.State.Off, QIcon.State.On):
                icon.addPixmap(pm, m, st)
        icon.addPixmap(svg_pixmap(name, disabled, size), QIcon.Mode.Disabled, QIcon.State.Off)
        return icon
    icon = QIcon()
    icon.addPixmap(svg_pixmap(name, color, size), QIcon.Mode.Normal, QIcon.State.Off)
    icon.addPixmap(svg_pixmap(name, hover or color, size), QIcon.Mode.Active, QIcon.State.Off)
    icon.addPixmap(svg_pixmap(name, active, size), QIcon.Mode.Normal, QIcon.State.On)
    icon.addPixmap(svg_pixmap(name, active, size), QIcon.Mode.Active, QIcon.State.On)
    icon.addPixmap(svg_pixmap(name, active, size), QIcon.Mode.Selected, QIcon.State.On)
    icon.addPixmap(svg_pixmap(name, disabled, size), QIcon.Mode.Disabled, QIcon.State.Off)
    return icon



def _write_arrow(name, path_d):
    """把 Bootstrap 風格的箭頭存成 SVG 供樣式表使用（數值框 / 下拉選單）"""
    import tempfile
    d = os.path.join(tempfile.gettempdir(), "cgu_pdftool_ui")
    os.makedirs(d, exist_ok=True)
    f = os.path.join(d, name)
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16">'
           f'<path fill="none" stroke="#495057" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" d="{path_d}"/></svg>')
    try:
        if not os.path.exists(f) or open(f, encoding="utf-8").read() != svg:
            with open(f, "w", encoding="utf-8") as fh:
                fh.write(svg)
    except OSError:
        pass
    return f.replace("\\", "/")


_ARROW_UP = _write_arrow("up.svg", "M3 10.5 8 5.5l5 5")
_ARROW_DOWN = _write_arrow("down.svg", "m3 5.5 5 5 5-5")


def bs(btn, variant):
    """套用 Bootstrap 按鈕樣式：primary / outline-primary / success / danger / outline-danger / outline-secondary"""
    btn.setProperty("variant", variant)
    btn.setProperty("primary", False)
    st = btn.style()
    st.unpolish(btn); st.polish(btn)
    return btn

STYLE = f"""
* {{ font-family: "Microsoft JhengHei UI", "Microsoft JhengHei", "Noto Sans CJK TC", "PingFang TC", sans-serif; font-size: 13px; color: #2C2C2C; }}
QMainWindow, #Body {{ background: #F0F0F0; }}

/* ---- 頂部標題列 ---- */
#Header {{ background: #1F1F1F; border-bottom: 2px solid {ACCENT}; }}
#Header QLabel {{ color: #B3B3B3; }}
#Header QLabel#AppName {{ font-size: 14px; font-weight: 600; color: #FFFFFF; }}
#Header QLabel#DocName {{ color: #B3B3B3; font-size: 12px; }}
#Header QLabel#VerBadge {{ background: {ACCENT}; color: #1F1F1F; font-size: 10px; font-weight: bold; border-radius: 7px; padding: 0px 6px; }}
#MiniBar {{ background: #FFFFFF; border: 1px solid #D0D0D0; border-radius: 6px; }}
#MiniBar QToolButton {{ border: 1px solid transparent; border-radius: 4px; padding: 2px 6px; }}
#MiniBar QToolButton:hover {{ background: {TINT}; border-color: {TINT_BORDER}; }}
#MiniBar QToolButton#StyleToggle {{ min-width: 18px; color: #6c757d; border: 1px solid #6c757d; border-radius: 6px; padding: 1px 6px; font-size: 13px; }}
#MiniBar QToolButton#StyleToggle:hover {{ background: #e9ecef; border-color: #6c757d; }}
#MiniBar QToolButton#StyleToggle:checked {{ background: #6c757d; color: #FFFFFF; border-color: #6c757d; }}
#Header QLabel#Copyright {{ color: #A8A8A8; font-size: 11px; }}
#AboutBtn {{ color: #CCCCCC; background: transparent; border: 1px solid #3A3A3A; border-radius: 11px; padding: 1px 10px; font-size: 12px; }}
#AboutBtn:hover {{ background: #2D2D2D; color: #FFFFFF; border-color: #4A4A4A; }}

/* ---- 主工具列 ---- */
#MainToolbar {{ background: #FFFFFF; border: none; border-bottom: 1px solid #D5D5D5; padding: 3px 8px; spacing: 2px; }}
QToolBar QToolButton {{ border: 1px solid transparent; border-radius: 4px; padding: 4px; }}
QToolBar QToolButton:hover {{ background: #EDEDED; border-color: #DDDDDD; }}
QToolBar QToolButton:pressed {{ background: #E0E0E0; }}
QToolBar QToolButton:checked {{ background: {TINT}; border-color: {TINT_BORDER}; }}
QToolBar::separator {{ background: #DADADA; width: 1px; margin: 5px 6px; }}
#PageBox {{ min-width: 40px; max-height: 24px; padding: 1px 4px; }}

/* ---- 左側工具軌 ---- */
#Rail {{ background: #FFFFFF; border-right: 1px solid #D5D5D5; }}
#Rail QToolButton {{ border: none; border-left: 3px solid transparent; padding: 8px 2px 6px 2px;
                     color: #505050; font-size: 12px; min-width: 78px; }}
#Rail QToolButton:hover {{ background: #F3F3F3; }}
#Rail QToolButton:checked {{ background: {TINT}; border-left: 3px solid {ACCENT}; color: {ACCENT_DARK}; font-weight: bold; }}

/* ---- 工具頁面 ---- */
#ToolPage {{ background: #F0F0F0; }}
#ToolTitle {{ font-size: 20px; font-weight: bold; }}
#ToolDesc {{ color: #6E6E6E; }}
#Card {{ background: #FFFFFF; border: 1px solid #DCDCDC; border-radius: 8px; }}

/* ---- 編輯檢視：縮圖面板、畫布、屬性面板 ---- */
#ThumbPanel, #PropPanel {{ background: #FAFAFA; }}
#ThumbPanel {{ border-right: 1px solid #D5D5D5; }}
#PropPanel {{ border-left: 1px solid #D5D5D5; }}
#PanelTitle {{ font-weight: bold; font-size: 14px; padding: 2px 0; }}
#PanelHint {{ color: #6E6E6E; background: #FFFFFF; border: 1px solid #E2E2E2; border-radius: 6px; padding: 8px; }}
#CanvasBg, #CanvasScroll {{ background: #DCDCDC; border: none; }}
#Canvas {{ background: #FFFFFF; color: #8A8A8A; font-size: 15px; }}
#EditToolbar {{ background: #FFFFFF; border: none; border-bottom: 1px solid #D5D5D5; padding: 2px 8px; }}
#EditToolbar QToolButton {{ padding: 3px 2px; font-size: 12px; min-width: 46px; }}
#ThumbList {{ background: #FAFAFA; border: none; }}
#ThumbList::item {{ border: 2px solid transparent; border-radius: 4px; padding: 4px; color: #555; }}
#ThumbList::item:selected {{ border-color: {ACCENT}; background: {TINT}; color: {ACCENT_DARK}; }}
#ThumbList::item:hover {{ background: #EFEFEF; }}

/* ---- 一般控制項 ---- */
/* ---- 按鈕 / 表單：Bootstrap 5 配色與形狀（MIT） ---- */
QPushButton {{ background: #FFFFFF; border: 1px solid #6c757d; color: #6c757d; border-radius: 6px; padding: 6px 12px; font-size: 14px; min-height: 18px; }}
QPushButton:hover {{ background: #6c757d; color: #FFFFFF; }}
QPushButton:pressed {{ background: #565e64; border-color: #51585e; color: #FFFFFF; }}
QPushButton:focus {{ border: 1px solid #86b7fe; outline: none; }}
QPushButton:disabled {{ color: rgba(108,117,125,0.65); border-color: rgba(108,117,125,0.65); background: transparent; }}
QPushButton[variant="primary"], QPushButton[primary="true"] {{ background: #0d6efd; border-color: #0d6efd; color: #FFFFFF; }}
QPushButton[variant="primary"]:hover, QPushButton[primary="true"]:hover {{ background: #0b5ed7; border-color: #0a58ca; }}
QPushButton[variant="primary"]:pressed, QPushButton[primary="true"]:pressed {{ background: #0a58ca; border-color: #0a53be; }}
QPushButton[variant="primary"]:disabled, QPushButton[primary="true"]:disabled {{ background: rgba(13,110,253,0.65); border-color: rgba(13,110,253,0.0); color: #FFFFFF; }}
QPushButton[variant="outline-primary"] {{ background: #FFFFFF; border-color: #0d6efd; color: #0d6efd; }}
QPushButton[variant="outline-primary"]:hover {{ background: #0d6efd; color: #FFFFFF; }}
QPushButton[variant="outline-primary"]:pressed {{ background: #0a58ca; color: #FFFFFF; }}
QPushButton[variant="success"] {{ background: #198754; border-color: #198754; color: #FFFFFF; }}
QPushButton[variant="success"]:hover {{ background: #157347; border-color: #146c43; }}
QPushButton[variant="success"]:pressed {{ background: #146c43; border-color: #13653f; }}
QPushButton[variant="success"]:disabled {{ background: rgba(25,135,84,0.65); color: #FFFFFF; border-color: transparent; }}
QPushButton[variant="danger"] {{ background: #dc3545; border-color: #dc3545; color: #FFFFFF; }}
QPushButton[variant="danger"]:hover {{ background: #bb2d3b; border-color: #b02a37; }}
QPushButton[variant="danger"]:pressed {{ background: #b02a37; }}
QPushButton[variant="outline-danger"] {{ background: #FFFFFF; border-color: #dc3545; color: #dc3545; }}
QPushButton[variant="outline-danger"]:hover {{ background: #dc3545; color: #FFFFFF; }}
QPushButton[variant="outline-danger"]:pressed {{ background: #b02a37; color: #FFFFFF; }}
QPushButton[variant="danger"]:disabled, QPushButton[variant="outline-danger"]:disabled {{ color: rgba(220,53,69,0.65); border-color: rgba(220,53,69,0.65); background: transparent; }}
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox, QPlainTextEdit {{ background: #FFFFFF; border: 1px solid #dee2e6; border-radius: 6px; padding: 4px 8px; color: #212529; selection-background-color: #0d6efd; }}
QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus, QPlainTextEdit:focus {{ border: 1px solid #86b7fe; background: #FFFFFF; }}
QLineEdit:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled, QComboBox:disabled {{ background: #e9ecef; color: #6c757d; }}
QSpinBox, QDoubleSpinBox {{ padding-right: 22px; }}
QSpinBox::up-button, QDoubleSpinBox::up-button {{ subcontrol-origin: border; subcontrol-position: top right; width: 20px; border: none; border-left: 1px solid #dee2e6; border-top-right-radius: 6px; background: #f8f9fa; }}
QSpinBox::down-button, QDoubleSpinBox::down-button {{ subcontrol-origin: border; subcontrol-position: bottom right; width: 20px; border: none; border-left: 1px solid #dee2e6; border-top: 1px solid #dee2e6; border-bottom-right-radius: 6px; background: #f8f9fa; }}
QSpinBox::up-button:hover, QSpinBox::down-button:hover, QDoubleSpinBox::up-button:hover, QDoubleSpinBox::down-button:hover {{ background: #e9ecef; }}
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {{ image: url({_ARROW_UP}); width: 10px; height: 10px; }}
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {{ image: url({_ARROW_DOWN}); width: 10px; height: 10px; }}
QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox::down-arrow {{ image: url({_ARROW_DOWN}); width: 10px; height: 10px; }}
QListWidget {{ background: #FFFFFF; border: 1px solid #D5D5D5; border-radius: 6px; padding: 4px; }}
QListWidget::item {{ padding: 5px; border-radius: 4px; }}
QListWidget::item:selected {{ background: {TINT}; color: {ACCENT_DARK}; }}
QListWidget::item:hover {{ background: #F4F4F4; }}
QProgressBar {{ background: #e9ecef; border: none; border-radius: 6px; max-height: 14px; min-height: 14px; text-align: center; font-size: 10px; color: #212529; }}
QProgressBar::chunk {{ background: #0d6efd; border-radius: 6px; }}
QSlider::groove:horizontal {{ height: 4px; background: #D0D0D0; border-radius: 2px; }}
QSlider::sub-page:horizontal {{ background: {ACCENT}; border-radius: 2px; }}
QSlider::handle:horizontal {{ background: #FFFFFF; border: 2px solid {ACCENT}; width: 12px; margin: -6px 0; border-radius: 8px; }}
QCheckBox::indicator:checked {{ background: {ACCENT}; border: 1px solid {ACCENT}; border-radius: 3px; }}
QStatusBar {{ background: #FFFFFF; border-top: 1px solid #D5D5D5; color: #555; }}
QStatusBar QLabel {{ color: #555; padding: 0 8px; }}
QScrollBar:vertical {{ background: transparent; width: 10px; }}
QScrollBar::handle:vertical {{ background: #B5B5B5; border-radius: 4px; min-height: 30px; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; }}
QScrollBar::handle:horizontal {{ background: #B5B5B5; border-radius: 4px; min-width: 30px; }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QToolTip {{ background: #333; color: #FFF; border: none; padding: 4px 8px; }}
"""


# =====================================================================
# 工具 5：編輯 PDF
# =====================================================================

class PageCanvas(QLabel):
    """
    顯示 PDF 頁面的畫布，處理滑鼠事件：
      * clicked(QPointF)                 ：單擊（選文字、加文字、便利貼…）
      * dragged(QPointF, QPointF)        ：拖曳出矩形 / 線條
      * strokeFinished(list[QPointF])    ：手繪筆畫
    座標皆為「畫面像素」，由 EditTab 轉為 PDF 座標。
    """
    clicked = pyqtSignal(QPointF)
    contextRequested = pyqtSignal(QPointF, QPoint)   # 右鍵（畫面座標, 螢幕座標）
    dragged = pyqtSignal(QPointF, QPointF)
    strokeFinished = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setObjectName("Canvas")
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.mode = "select"
        self.start = None
        self.current = None
        self.points = []

    def mousePressEvent(self, e):
        if e.button() != Qt.MouseButton.LeftButton or self.pixmap() is None or self.pixmap().isNull():
            return
        self.start = self.current = e.position()
        self.points = [e.position()]

    def mouseMoveEvent(self, e):
        if self.start is None:
            return
        self.current = e.position()
        if self.mode == "pen":
            self.points.append(e.position())
        self.update()

    def contextMenuEvent(self, e):
        self.contextRequested.emit(QPointF(e.pos()), e.globalPos())

    def mouseReleaseEvent(self, e):
        if self.start is None:
            return
        end = e.position()
        start, self.start = self.start, None
        moved = (end - start).manhattanLength() > 4
        if self.mode == "pen" and len(self.points) > 1:
            self.strokeFinished.emit(list(self.points))
        elif self.mode in ("highlight", "rect", "line", "whiteout", "link", "sign", "check") and moved:
            self.dragged.emit(start, end)
        else:
            self.clicked.emit(end)
        self.points = []
        self.update()

    def paintEvent(self, e):
        super().paintEvent(e)
        if self.start is None:
            return
        # 拖曳時即時畫出預覽框 / 線 / 筆跡
        p = QPainter(self)
        p.setPen(QPen(QColor(ACCENT), 2, Qt.PenStyle.DashLine))
        if self.mode == "pen":
            for a, b in zip(self.points, self.points[1:]):
                p.drawLine(a, b)
        elif self.mode == "line":
            p.drawLine(self.start, self.current)
        elif self.mode in ("highlight", "rect", "whiteout", "link", "sign", "check"):
            p.drawRect(QRect(self.start.toPoint(), self.current.toPoint()).normalized())
        p.end()


def centered_rect(parent_rect, size):
    """計算讓對話框置中於主視窗的位置（純函式，方便測試）"""
    x = parent_rect.x() + (parent_rect.width() - size.width()) // 2
    y = parent_rect.y() + (parent_rect.height() - size.height()) // 2
    return QRect(x, y, size.width(), size.height())


def settings():
    """使用者偏好設定（Windows 存在登錄檔 HKCU\\Software\\CGU\\PDFTool）"""
    return QSettings("CGU", "PDFTool")


def render_font_preview(text, fontfile=None, fontbuffer=None, size=28, color=(0, 0, 0)):
    """用指定字型把文字畫成 QPixmap（預覽用，與實際寫入 PDF 的效果一致）"""
    text = text or " "
    try:
        f = fitz.Font(fontfile=fontfile) if fontfile else fitz.Font(fontbuffer=fontbuffer) if fontbuffer else fitz.Font("helv")
        w = max(40, f.text_length(text, fontsize=size) + 16)
        d = fitz.open()
        p = d.new_page(width=min(w, 1400), height=size * 1.7)
        if fontfile:
            p.insert_text((8, size * 1.2), text, fontsize=size, fontfile=fontfile, fontname="P", color=color)
        elif fontbuffer:
            p.insert_font(fontname="P", fontbuffer=fontbuffer)
            p.insert_text((8, size * 1.2), text, fontsize=size, fontname="P", color=color)
        else:
            p.insert_text((8, size * 1.2), text, fontsize=size, fontname="helv", color=color)
        return fitz_pixmap_to_qpixmap(p.get_pixmap(alpha=False))
    except Exception:
        return QPixmap()


class EditSpanDialog(QDialog):
    """
    修改既有文字的對話框（強制回應、置中於主視窗）：
      * 預先填入點選的文字；顯示偵測到的字型、大小、顏色
      * 字型下拉選單（系統字型 + 「沿用原字型」+ 自動），即時預覽
      * 可改字級；按 Enter = 確定；可勾選「文字過長時自動縮小」
    """
    AUTO = "auto"
    ORIGINAL = "original"

    def __init__(self, page, span, parent=None, preferred_font=None):
        super().__init__(parent)
        self.page, self.span = page, span
        self.setWindowTitle("修改文字")
        self.setModal(True)
        self.setMinimumWidth(520)
        st = span_style(page, span)
        self.st = st
        self.ocr = st["ocr"]
        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        title = QLabel("修改文字")
        title.setObjectName("PanelTitle")
        lay.addWidget(title)
        form = QFormLayout()
        r, g, b = (int(c * 255) for c in st["color"])
        col = QColor(r, g, b)
        form.addRow("原字型：", QLabel(st["font"] or "（未知）"))
        crow = QHBoxLayout()
        chip = QLabel()
        chip.setFixedSize(16, 16)
        chip.setStyleSheet(f"background:{col.name()}; border:1px solid #888; border-radius:3px;")
        crow.addWidget(chip)
        crow.addWidget(QLabel(col.name().upper()))
        crow.addStretch()
        form.addRow("顏色：", crow)

        # ---- 字型下拉選單 ----
        self.cmb_font = QComboBox()
        self.cmb_font.addItem("自動（依原字型對應）", self.AUTO)
        if not self.ocr and choose_font(page, span, span["text"])["kind"] == "embedded":
            self.cmb_font.addItem(f"沿用原字型（{strip_subset(span.get('font', ''))}）", self.ORIGINAL)
        for label, path in list_ui_fonts():
            self.cmb_font.addItem(label, path)
        if preferred_font:
            i = self.cmb_font.findData(preferred_font)
            if i >= 0:
                self.cmb_font.setCurrentIndex(i)
        form.addRow("字型：", self.cmb_font)

        self.spin_size = QDoubleSpinBox()
        self.spin_size.setRange(4, 200)
        self.spin_size.setDecimals(1)
        self.spin_size.setSingleStep(0.5)
        self.spin_size.setSuffix(" pt")
        base = default_edit_size(page, span, resolve_font(page, span, preferred_font, span["text"], st), st)
        self.spin_size.setValue(round(base, 1))
        self.orig_size = round(base, 1)
        form.addRow("字級：", self.spin_size)
        self.lbl_font = QLabel()
        self.lbl_font.setWordWrap(True)
        self.lbl_font.setObjectName("PanelHint")
        form.addRow("將使用：", self.lbl_font)
        lay.addLayout(form)

        self.edit = QLineEdit(span["text"])
        self.edit.selectAll()
        self.edit.returnPressed.connect(self.accept)   # Enter 直接套用
        lay.addWidget(self.edit)
        # ---- 即時預覽 ----
        self.preview = QLabel()
        self.preview.setMinimumHeight(60)
        self.preview.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.preview.setStyleSheet("background:#FFFFFF; border:1px dashed #C8C8C8; border-radius:4px; padding:4px;")
        lay.addWidget(QLabel("預覽："))
        lay.addWidget(self.preview)
        srow = QHBoxLayout()
        ob, oi = (False, False) if self.ocr else span_style_flags(span)
        self.orig_style = (ob, oi, False)
        self.chk_bold = QCheckBox("粗體")
        self.chk_bold.setChecked(ob)
        self.chk_italic = QCheckBox("斜體")
        self.chk_italic.setChecked(oi)
        self.chk_under = QCheckBox("底線")
        for c in (self.chk_bold, self.chk_italic, self.chk_under):
            srow.addWidget(c)
        srow.addStretch()
        lay.addLayout(srow)
        self.chk_shrink = QCheckBox("新文字較寬時自動縮小字級以符合原範圍")
        self.chk_shrink.setChecked(False)      # 預設關閉：保持原字級，文字可往右延伸
        lay.addWidget(self.chk_shrink)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("確定")
        bb.button(QDialogButtonBox.StandardButton.Ok).setProperty("primary", True)
        bb.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.edit.textChanged.connect(self.refresh)
        self.cmb_font.currentIndexChanged.connect(self.refresh)
        self.spin_size.valueChanged.connect(self.refresh)
        self.refresh()

    def resolved_font(self):
        """依下拉選單決定實際字型 → (字型資訊 dict, 傳給 replace_span_text 的 fontfile 參數)"""
        t = self.text() or " "
        data = self.cmb_font.currentData()
        if data not in (self.AUTO, self.ORIGINAL, None):
            return font_info_from_path(data, self.cmb_font.currentText()), data
        if data == self.ORIGINAL or not self.ocr:
            return choose_font(self.page, self.span, t), None
        if OCR_DEFAULT_FONT and os.path.exists(OCR_DEFAULT_FONT):
            return font_info_from_path(OCR_DEFAULT_FONT, "OCR 預設字型"), None
        fake = {"font": "PMingLiU" if has_cjk(t) else "Times New Roman", "flags": 4}
        return choose_font(self.page, fake, t), None

    def refresh(self):
        t = self.text() or " "
        fi, _ = self.resolved_font()
        label = fi["label"]
        if not font_covers(fi["font"], t):
            label += "\n⚠ 此字型缺少部分字元，可能顯示為空白"
        if self.ocr:
            label = "掃描文字：以背景色覆蓋原字後重新寫入\n" + label
        self.lbl_font.setText(label)
        ins = fi["insert"]
        pv_size = max(14, min(40, self.spin_size.value() * 2))
        pm = render_font_preview(t, ins.get("fontfile"), ins.get("fontbuffer") if "fontfile" not in ins and "fontbuffer" in ins else None,
                                 size=pv_size, color=self.st["color"]) if ("fontfile" in ins or "fontbuffer" in ins) else \
            render_font_preview(t, size=pv_size, color=self.st["color"])
        if not pm.isNull() and pm.width() > 470:
            pm = pm.scaledToWidth(470, Qt.TransformationMode.SmoothTransformation)
        self.preview.setPixmap(pm)

    def center_on_parent(self):
        """置中於主視窗（沒有主視窗時置中於螢幕）"""
        self.adjustSize()
        par = self.parentWidget()
        ref = par.window().frameGeometry() if par else self.screen().availableGeometry()
        self.setGeometry(centered_rect(ref, self.sizeHint().expandedTo(self.minimumSize())))

    def showEvent(self, e):
        self.center_on_parent()
        super().showEvent(e)

    def text(self):
        return self.edit.text()

    def shrink(self):
        return self.chk_shrink.isChecked()

    def style(self):
        return (self.chk_bold.isChecked(), self.chk_italic.isChecked(), self.chk_under.isChecked())

    def style_changed(self):
        return self.style() != self.orig_style

    def font_choice(self):
        """回傳 replace_span_text 的 fontfile 參數：字型檔路徑 / 'original' / None(自動)"""
        data = self.cmb_font.currentData()
        return None if data == self.AUTO else data

    def size(self):
        """使用者有改字級才回傳，否則 None（沿用原大小）"""
        v = round(self.spin_size.value(), 1)
        return None if abs(v - self.orig_size) < 0.05 else v


class InlineLineEdit(QLineEdit):
    """畫布上的就地編輯框：Enter 確定、Esc 取消、Tab 確定並編輯下一行"""
    commit = pyqtSignal()
    cancel = pyqtSignal()
    nextLine = pyqtSignal()
    styleKey = pyqtSignal(str)       # Ctrl+B / Ctrl+I / Ctrl+U

    def keyPressEvent(self, e):
        k = e.key()
        if e.modifiers() & Qt.KeyboardModifier.ControlModifier and k in (Qt.Key.Key_B, Qt.Key.Key_I, Qt.Key.Key_U):
            self.styleKey.emit({Qt.Key.Key_B: "B", Qt.Key.Key_I: "I", Qt.Key.Key_U: "U"}[k])
            return
        if k in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.commit.emit()
        elif k == Qt.Key.Key_Escape:
            self.cancel.emit()
        elif k == Qt.Key.Key_Tab:
            self.nextLine.emit()
        else:
            super().keyPressEvent(e)

    def focusNextPrevChild(self, nxt):
        return False      # 讓 Tab 交給 keyPressEvent 處理


_QFONT_FAMILY = {}


def qt_family_for(fi):
    """把替換字型載入 Qt（QFontDatabase），回傳字型家族名稱，讓編輯框顯示相同字型"""
    ins = fi["insert"]
    key = ins.get("fontfile") or fi["label"]
    if key in _QFONT_FAMILY:
        return _QFONT_FAMILY[key]
    fam = None
    try:
        if "fontfile" in ins:
            fid = QFontDatabase.addApplicationFont(ins["fontfile"])
        elif "fontbuffer" in ins:
            fid = QFontDatabase.addApplicationFontFromData(QByteArray(ins["fontbuffer"]))
        else:
            fid = -1
        if fid >= 0:
            fams = QFontDatabase.applicationFontFamilies(fid)
            fam = fams[0] if fams else None
    except Exception:
        fam = None
    if not fam:   # 內建 Base-14 字型對應
        code = ins.get("fontname", "")
        fam = "Times New Roman" if code.startswith("ti") else "Courier New" if code.startswith("co") else "Arial"
    _QFONT_FAMILY[key] = fam
    return fam


class InlineEditor(QObject):
    """
    類似 Acrobat 的就地編輯：在頁面文字上方疊一個編輯框（同字型 / 字級 / 顏色，白底蓋住原字），
    上方附迷你工具列（字型、字級、顏色、確定、取消、進階）。
    """
    finished = pyqtSignal(bool)      # True = 已套用（或無變更），False = 取消

    def __init__(self, tab, span):
        super().__init__(tab)
        self.tab, self.span = tab, span
        self.page = tab.doc[tab.page_no]
        self.st = span_style(self.page, span)
        self.ocr = self.st["ocr"]
        self.color = QColor(*(int(c * 255) for c in self.st["color"]))
        self.color_changed = False
        canvas = tab.canvas
        pref = (tab.doc_font_choice or settings().value("last_font", None)) if self.ocr else None
        # ---- 編輯框 ----
        self.edit = InlineLineEdit(canvas)
        self.edit.setObjectName("InlineEdit")
        self.edit.setText(span["text"])
        self.edit.setFrame(False)
        self.edit.setTextMargins(0, 0, 0, 0)
        self.edit.commit.connect(self.apply)
        self.edit.cancel.connect(self.discard)
        self.edit.nextLine.connect(self.apply_and_next)
        self.edit.textChanged.connect(self.relayout)
        # ---- 迷你工具列 ----
        self.bar = QFrame(canvas)
        self.bar.setObjectName("MiniBar")
        bl = QHBoxLayout(self.bar)
        bl.setContentsMargins(6, 4, 6, 4)
        bl.setSpacing(4)
        self.cmb = QComboBox()
        self.cmb.addItem("自動", None)
        if not self.ocr and choose_font(self.page, span, span["text"])["kind"] == "embedded":
            self.cmb.addItem("沿用原字型", "original")
        for label, path in list_ui_fonts():
            self.cmb.addItem(label, path)
        if pref:
            i = self.cmb.findData(pref)
            if i >= 0:
                self.cmb.setCurrentIndex(i)
        self.cmb.setMinimumWidth(120)
        self.spin = QDoubleSpinBox()
        self.spin.setRange(4, 200)
        self.spin.setDecimals(1)
        self.spin.setSingleStep(0.5)
        self.spin.setSuffix(" pt")
        self.btn_color = QToolButton()
        self.btn_color.setToolTip("文字顏色")
        self.btn_ok = QToolButton()
        self.btn_ok.setText("✓")
        self.btn_ok.setToolTip("套用（Enter）")
        self.btn_cancel = QToolButton()
        self.btn_cancel.setText("✕")
        self.btn_cancel.setToolTip("取消（Esc）")
        self.btn_adv = QToolButton()
        self.btn_adv.setText("進階…")
        self.btn_adv.setToolTip("開啟進階編輯對話框")
        # 粗體 / 斜體 / 底線（原生文字預先帶入原樣式）
        ob, oi = (False, False) if self.ocr else span_style_flags(span)
        self.base_style = (ob, oi, False)
        self.btn_b, self.btn_i, self.btn_u = (self._style_btn(t, tip, on) for t, tip, on in (
            ("B", "粗體 (Ctrl+B)", ob), ("I", "斜體 (Ctrl+I)", oi), ("U", "底線 (Ctrl+U)", False)))
        f_ = self.btn_b.font(); f_.setBold(True); self.btn_b.setFont(f_)
        f_ = self.btn_i.font(); f_.setItalic(True); self.btn_i.setFont(f_)
        f_ = self.btn_u.font(); f_.setUnderline(True); self.btn_u.setFont(f_)
        for w in (self.cmb, self.spin, self.btn_color, self.btn_b, self.btn_i, self.btn_u,
                  self.btn_ok, self.btn_cancel, self.btn_adv):
            bl.addWidget(w)
        self.edit.styleKey.connect(lambda k: {"B": self.btn_b, "I": self.btn_i, "U": self.btn_u}[k].toggle())
        self.fi = resolve_font(self.page, span, self.cmb.currentData(), span["text"], self.st)
        self.base_size = default_edit_size(self.page, span, self.fi, self.st)
        self.spin.setValue(round(self.base_size, 1))
        self.base_choice = self.cmb.currentData()
        self.update_color_btn()
        self.cmb.currentIndexChanged.connect(self.on_font)
        self.spin.valueChanged.connect(self.relayout)
        self.btn_color.clicked.connect(self.pick_color)
        self.btn_ok.clicked.connect(self.apply)
        self.btn_cancel.clicked.connect(self.discard)
        self.btn_adv.clicked.connect(self.advanced)
        self.done = False
        self.relayout()
        self.edit.show()
        self.bar.show()
        self.edit.setFocus()
        self.edit.selectAll()

    def _style_btn(self, text, tip, on):
        b = QToolButton()
        b.setObjectName("StyleToggle")
        b.setText(text)
        b.setToolTip(tip)
        b.setCheckable(True)
        b.setChecked(on)
        b.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        b.toggled.connect(lambda _=False: (self.relayout(), self.edit.setFocus()))
        return b

    def style(self):
        return (self.btn_b.isChecked(), self.btn_i.isChecked(), self.btn_u.isChecked())

    # ---- 外觀 / 位置 ----
    def update_color_btn(self):
        self.btn_color.setStyleSheet(f"QToolButton{{background:{self.color.name()};border:1px solid #888;"
                                     f"border-radius:3px;min-width:18px;max-width:18px;min-height:18px;}}")

    def on_font(self):
        self.fi = resolve_font(self.page, self.span, self.cmb.currentData(), self.edit.text(), self.st)
        if self.ocr and not self.size_changed():   # OCR：換字型時重新校正字級
            self.base_size = default_edit_size(self.page, self.span, self.fi, self.st)
            self.spin.blockSignals(True)
            self.spin.setValue(round(self.base_size, 1))
            self.spin.blockSignals(False)
        self.relayout()

    def size_changed(self):
        return abs(self.spin.value() - round(self.base_size, 1)) >= 0.05

    def relayout(self):
        """依目前縮放倍率把編輯框放到原文字上，字型 / 字級 / 顏色與替換結果一致"""
        z = self.tab.zoom
        size = self.spin.value()
        f = QFont(qt_family_for(self.fi))
        f.setPixelSize(max(4, int(round(size * z))))
        b_, i_, u_ = self.style() if hasattr(self, "btn_u") else (False, False, False)
        f.setBold(b_)
        f.setItalic(i_)
        f.setUnderline(u_)
        self.edit.setFont(f)
        bg = self.st["bg"] or (1, 1, 1)
        bgc = QColor(*(int(c * 255) for c in bg))
        self.edit.setStyleSheet(f"QLineEdit#InlineEdit{{background:{bgc.name()};color:{self.color.name()};"
                                f"border:1.5px solid {ACCENT};border-radius:2px;padding:0 1px;"
                                f"font-family:'{f.family()}';font-size:{f.pixelSize()}px;"
                                f"font-weight:{'bold' if b_ else 'normal'};font-style:{'italic' if i_ else 'normal'};"
                                f"text-decoration:{'underline' if u_ else 'none'};"
                                f"selection-background-color:{TINT_BORDER};selection-color:#000;}}")
        bb = fitz.Rect(self.span["bbox"])
        tw = self.fi["font"].text_length(self.edit.text() or " ", fontsize=size)
        x = bb.x0 * z - 3
        h = max(bb.height, size * 1.25) * z + 4
        y = (bb.y0 + bb.y1) / 2 * z - h / 2
        w = max(bb.width, tw) * z + 12
        self.edit.setGeometry(int(x), int(y), int(w), int(h))
        self.bar.adjustSize()
        by = y - self.bar.height() - 6
        if by < 0:
            by = y + h + 6
        bx = min(max(0, x), max(0, self.tab.canvas.width() - self.bar.width()))
        self.bar.move(int(bx), int(by))
        self.bar.raise_()

    def pick_color(self):
        c = QColorDialog.getColor(self.color, self.tab, "文字顏色")
        if c.isValid():
            self.color = c
            self.color_changed = True
            self.update_color_btn()
            self.relayout()
        self.edit.setFocus()

    # ---- 結束 ----
    def close_widgets(self):
        self.done = True
        self.edit.hide()
        self.bar.hide()
        self.edit.deleteLater()
        self.bar.deleteLater()

    def changed(self):
        return (self.edit.text() != self.span["text"] or self.cmb.currentData() != self.base_choice
                or self.size_changed() or self.color_changed or self.style() != self.base_style)

    def apply(self):
        """套用修改（可復原）；沒有變更就直接關閉"""
        if self.done:
            return
        text, choice = self.edit.text(), self.cmb.currentData()
        size = self.spin.value() if (self.size_changed() or not self.ocr) else None
        color = color_to_rgb01(self.color) if self.color_changed else None
        changed = self.changed()
        b_, i_, u_ = self.style()
        self.close_widgets()
        if changed:
            self.tab.apply_span_edit(self.span, text, choice, size, color, self.ocr, bold=b_, italic=i_, underline=u_)
        self.finished.emit(True)

    def apply_and_next(self):
        span = self.span
        self.apply()
        self.tab.edit_next_after(span)

    def discard(self):
        if self.done:
            return
        self.close_widgets()
        self.finished.emit(False)
        self.tab.status("已取消修改")

    def advanced(self):
        span = self.span
        self.discard()
        self.tab.advanced_edit(span)


class EditTab(FileDropMixin, BaseTab):
    """
    編輯檢視：左側頁面縮圖（可收合）＋ 中央灰色畫布（頁面帶陰影）＋ 右側屬性面板。
    檔案 / 復原 / 縮放 / 換頁等按鈕放在主視窗上方工具列（由 MainWindow 呼叫本類別的方法）。
    """
    TOOLS = [
        ("select", "修改文字", "點選頁面上的文字片段即可修改原有文字。"),
        ("text", "新增文字", "在頁面上點一下要放文字的位置，輸入內容即可。可在下方設定字級與顏色。"),
        ("image", "插入圖片", "點一下要放圖片的位置，再選擇 PNG / JPG 圖檔。"),
        ("sign", "簽名印章", "在下方選一個簽名或印章，點一下頁面放上去，再拖曳移動、拖四角縮放，"
                          "按 Enter 或點旁邊確定。已蓋好的章點一下可重新調整或刪除。"),
        ("check", "打勾", "點一下框框就打勾，再點一次取消；拖曳框選可一次勾選範圍內所有框框。"
                       "沒有偵測到框框（例如掃描檔）時，會在點的位置打勾。"),
        ("highlight", "螢光標記", "拖曳框選要標記的文字。可調整顏色與不透明度。"),
        ("rect", "方框", "拖曳畫出方框。可調整顏色、線寬與不透明度。"),
        ("line", "直線", "拖曳畫出直線。可調整顏色、線寬與不透明度。"),
        ("pen", "手繪", "按住滑鼠左鍵自由繪製。"),
        ("note", "便利貼", "點一下新增註解（便利貼）。"),
        ("link", "超連結", "點一下文字（整行）或拖曳框選範圍，設定連到網址或本檔第幾頁。"
                         "藍色虛線框是已有的連結，點一下可修改或刪除。"),
        ("whiteout", "塗銷", "拖曳框選要永久清除的內容（文字與圖片），以白色填滿。"),
        ("delannot", "刪除註解", "點選要刪除的註解、標記或方框。"),
    ]
    MAX_UNDO = 30
    statusChanged = pyqtSignal()     # 頁碼 / 縮放 / 檔名改變時通知主視窗更新

    def __init__(self, main):
        super().__init__(main)
        self.setAcceptDrops(True)
        self.doc = None
        self.path = None
        self.signed_count = 0
        self.page_no = 0
        self.zoom = 1.25
        self.undo_stack = []
        self.redo_stack = []
        self.dirty = False
        self.doc_font_choice = None      # 本文件 OCR 修改時使用者選過的字型
        self.inline = None               # 目前的就地編輯器
        self.sign_overlay = None         # 目前的可調整簽名框
        self.color = QColor(THEME["annot"])
        self.fontsize = 14

        # ---- 給主工具列使用的頁碼 / 縮放元件（由 MainWindow 放進工具列）----
        self.page_spin = QSpinBox()
        self.page_spin.setObjectName("PageBox")
        self.page_spin.setMinimum(1)
        self.page_spin.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.page_spin.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_spin.valueChanged.connect(lambda v: self.goto(v - 1))
        self.lbl_pages = QLabel(" / 0 ")
        self.lbl_zoom = QLabel("125%")
        self.lbl_zoom.setMinimumWidth(48)
        self.lbl_zoom.setAlignment(Qt.AlignmentFlag.AlignCenter)

        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---- 左：頁面縮圖面板（可收合）----
        self.thumb_panel = QFrame()
        self.thumb_panel.setObjectName("ThumbPanel")
        self.thumb_panel.setFixedWidth(150)
        tl = QVBoxLayout(self.thumb_panel)
        tl.setContentsMargins(8, 10, 4, 6)
        t = QLabel("頁面縮圖")
        t.setObjectName("PanelTitle")
        tl.addWidget(t)
        self.thumbs = QListWidget()
        self.thumbs.setObjectName("ThumbList")
        self.thumbs.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.thumbs.customContextMenuRequested.connect(self.thumb_menu)
        self.thumbs.setViewMode(QListView.ViewMode.IconMode)
        self.thumbs.setFlow(QListView.Flow.TopToBottom)
        self.thumbs.setWrapping(False)
        self.thumbs.setMovement(QListView.Movement.Static)
        self.thumbs.setIconSize(QSize(110, 140))
        self.thumbs.setSpacing(4)
        self.thumbs.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.thumbs.currentRowChanged.connect(lambda r: r >= 0 and r != self.page_no and self.goto(r))
        tl.addWidget(self.thumbs, 1)
        root.addWidget(self.thumb_panel)

        # ---- 中：編輯工具列 + 灰色畫布 ----
        center = QVBoxLayout()
        center.setSpacing(0)
        bar = QToolBar()
        bar.setObjectName("EditToolbar")
        bar.setIconSize(QSize(20, 20))
        bar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        group = QActionGroup(self)
        group.setExclusive(True)
        self.tool_actions = {}
        for key, name, tip in self.TOOLS:
            act = QAction(make_icon(key, tool=True), name, self, checkable=True)
            act.setToolTip(tip)
            act.triggered.connect(lambda _=False, k=key: self.set_tool(k))
            group.addAction(act)
            bar.addAction(act)
            self.tool_actions[key] = act
        self.tool_actions["select"].setChecked(True)
        center.addWidget(bar)

        self.canvas = PageCanvas()
        self.canvas.setText("請按上方「開啟」，或直接把 PDF 拖曳到這裡")
        self.canvas.setFixedSize(595, 760)
        self.canvas.clicked.connect(self.on_click)
        self.canvas.dragged.connect(self.on_drag)
        self.canvas.strokeFinished.connect(self.on_stroke)
        self.canvas.contextRequested.connect(self.on_context)
        shadow = QGraphicsDropShadowEffect(self)       # 頁面陰影
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 3)
        shadow.setColor(QColor(0, 0, 0, 90))
        self.canvas.setGraphicsEffect(shadow)
        bg = QWidget()
        bg.setObjectName("CanvasBg")
        bl = QVBoxLayout(bg)
        bl.setContentsMargins(40, 30, 40, 30)
        bl.addWidget(self.canvas, 0, Qt.AlignmentFlag.AlignCenter)
        self.scroll = QScrollArea()
        self.scroll.setObjectName("CanvasScroll")
        self.scroll.setWidget(bg)
        self.scroll.setWidgetResizable(True)
        center.addWidget(self.scroll, 1)
        root.addLayout(center, 1)

        # ---- 右：屬性 / 情境面板 ----
        prop = QFrame()
        prop.setObjectName("PropPanel")
        prop.setFixedWidth(232)
        pl = QVBoxLayout(prop)
        pl.setContentsMargins(14, 12, 14, 12)
        pl.setSpacing(10)
        self.lbl_tool = QLabel(self.TOOLS[0][1])
        self.lbl_tool.setObjectName("PanelTitle")
        pl.addWidget(self.lbl_tool)
        self.hint = QLabel(self.TOOLS[0][2])
        self.hint.setObjectName("PanelHint")
        self.hint.setWordWrap(True)
        pl.addWidget(self.hint)
        self.generic_box = QWidget()
        form = QFormLayout(self.generic_box)
        form.setContentsMargins(0, 0, 0, 0)
        form.setVerticalSpacing(10)
        self.btn_color = QPushButton()
        self.btn_color.clicked.connect(self.pick_color)
        form.addRow("顏色", self.btn_color)
        self.spin_font = QSpinBox()
        self.spin_font.setRange(6, 96)
        self.spin_font.setValue(self.fontsize)
        self.spin_font.setSuffix(" pt")
        form.addRow("字級", self.spin_font)
        self.spin_width = QDoubleSpinBox()
        self.spin_width.setRange(0.5, 12)
        self.spin_width.setSingleStep(0.5)
        self.spin_width.setValue(2)
        self.spin_width.setSuffix(" pt")
        form.addRow("線寬", self.spin_width)
        orow = QHBoxLayout()
        self.opacity = QSlider(Qt.Orientation.Horizontal)
        self.opacity.setRange(10, 100)
        self.opacity.setValue(100)
        self.lbl_op = QLabel("100%")
        self.lbl_op.setMinimumWidth(36)
        self.opacity.valueChanged.connect(lambda v: self.lbl_op.setText(f"{v}%"))
        orow.addWidget(self.opacity)
        orow.addWidget(self.lbl_op)
        form.addRow("不透明度", orow)
        self.spin_imgw = QSpinBox()
        self.spin_imgw.setRange(20, 1000)
        self.spin_imgw.setValue(200)
        self.spin_imgw.setSuffix(" pt")
        form.addRow("圖片寬度", self.spin_imgw)
        pl.addWidget(self.generic_box)

        # ---- 簽名 / 印章庫（選「簽名印章」工具時顯示）----
        self.sign_box = QWidget()
        sl = QVBoxLayout(self.sign_box)
        sl.setContentsMargins(0, 0, 0, 0)
        sl.setSpacing(6)
        self.sign_list = QListWidget()
        self.sign_list.setViewMode(QListView.ViewMode.IconMode)
        self.sign_list.setIconSize(QSize(88, 56))
        self.sign_list.setGridSize(QSize(98, 66))
        self.sign_list.setMovement(QListView.Movement.Static)
        self.sign_list.setResizeMode(QListView.ResizeMode.Adjust)
        self.sign_list.setMinimumHeight(220)
        self.sign_list.currentItemChanged.connect(self.on_sign_selected)
        sl.addWidget(self.sign_list)
        srow = QHBoxLayout()
        self.btn_sign_import = QPushButton("匯入圖檔")
        self.btn_sign_draw = QPushButton("手寫")
        self.btn_sign_del = QPushButton("刪除")
        self.btn_sign_del.setProperty("variant", "outline-danger")
        for b_ in (self.btn_sign_import, self.btn_sign_draw, self.btn_sign_del):
            b_.setStyleSheet("padding:4px 6px;")
            srow.addWidget(b_)
        sl.addLayout(srow)
        self.chk_sign_bg = QCheckBox("匯入時自動去除白底")
        self.chk_sign_bg.setChecked(True)
        sl.addWidget(self.chk_sign_bg)
        wrow = QFormLayout()
        self.spin_signw = QSpinBox()
        self.spin_signw.setRange(20, 600)
        self.spin_signw.setValue(100)
        self.spin_signw.setSuffix(" pt")
        self.spin_signw.setToolTip("點一下頁面時的蓋章寬度（72 pt ≈ 2.5 公分）")
        self.spin_signw.valueChanged.connect(self.on_signw_changed)
        wrow.addRow("蓋章寬度", self.spin_signw)
        sl.addLayout(wrow)
        self.btn_sign_import.clicked.connect(self.sign_import)
        self.btn_sign_draw.clicked.connect(self.sign_draw)
        self.btn_sign_del.clicked.connect(self.sign_delete)
        self.sign_box.setVisible(False)
        pl.addWidget(self.sign_box)
        self.refresh_signatures()

        # ---- 打勾工具選項 ----
        self.check_box = QWidget()
        cf = QFormLayout(self.check_box)
        cf.setContentsMargins(0, 0, 0, 0)
        cf.setVerticalSpacing(10)
        self.cmb_check_sym = QComboBox()
        for label, key in (("✓ 打勾", "check"), ("✗ 打叉", "cross"), ("● 實心點", "dot")):
            self.cmb_check_sym.addItem(label, key)
        self.cmb_check_color = QComboBox()
        for label, rgb in (("黑色", (0, 0, 0)), ("藍色", (0.11, 0.25, 0.63)), ("紅色", (0.8, 0.1, 0.1))):
            self.cmb_check_color.addItem(label, rgb)
        self.spin_check_size = QSpinBox()
        self.spin_check_size.setRange(6, 40)
        self.spin_check_size.setValue(12)
        self.spin_check_size.setSuffix(" pt")
        self.spin_check_size.setToolTip("沒有偵測到框框時（例如掃描檔）打勾的大小")
        cf.addRow("符號", self.cmb_check_sym)
        cf.addRow("顏色", self.cmb_check_color)
        cf.addRow("大小", self.spin_check_size)
        self.check_box.setVisible(False)
        pl.addWidget(self.check_box)

        # ---- OCR：掃描檔辨識文字 ----
        self.ocr_box = QWidget()
        ocr_outer = pl
        pl = QVBoxLayout(self.ocr_box)
        pl.setContentsMargins(0, 0, 0, 0)
        pl.setSpacing(10)
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#DDDDDD;")
        pl.addWidget(sep)
        t2 = QLabel("辨識文字 (OCR)")
        t2.setObjectName("PanelTitle")
        pl.addWidget(t2)
        h2 = QLabel("讓掃描 / 拍照的 PDF 變成可選取、可搜尋、可修改的文字（繁中＋英文，離線辨識）。")
        h2.setWordWrap(True)
        h2.setObjectName("ToolDesc")
        pl.addWidget(h2)
        rrow = QHBoxLayout()
        self.rb_ocr_cur = QRadioButton("目前頁面")
        self.rb_ocr_all = QRadioButton("全部頁面")
        self.rb_ocr_cur.setChecked(True)
        rrow.addWidget(self.rb_ocr_cur)
        rrow.addWidget(self.rb_ocr_all)
        pl.addLayout(rrow)
        frow = QFormLayout()
        self.cmb_ocr_font = QComboBox()
        self.cmb_ocr_font.addItem("自動", "")
        for label, path in list_ui_fonts():
            self.cmb_ocr_font.addItem(label, path)
        saved = settings().value("ocr_font", None)
        target = saved if saved is not None else (default_ocr_font_path() or "")
        i = self.cmb_ocr_font.findData(target)
        self.cmb_ocr_font.setCurrentIndex(i if i >= 0 else 0)
        self.cmb_ocr_font.currentIndexChanged.connect(self.on_ocr_font_changed)
        self.cmb_ocr_font.setToolTip("修改掃描（OCR）文字時預設使用的字型")
        frow.addRow("替換字型", self.cmb_ocr_font)
        pl.addLayout(frow)
        self.on_ocr_font_changed(save=False)
        self.btn_ocr = QPushButton("開始辨識")
        self.btn_ocr.setProperty("variant", "primary")
        self.btn_ocr.clicked.connect(lambda: self.run_ocr())
        pl.addWidget(self.btn_ocr, 0, Qt.AlignmentFlag.AlignRight)
        pl.addWidget(self.progress)
        pl = ocr_outer
        pl.addWidget(self.ocr_box)
        pl.addStretch()
        self.lbl_docinfo = QLabel("尚未開啟文件")
        self.lbl_docinfo.setObjectName("PanelHint")
        self.lbl_docinfo.setWordWrap(True)
        pl.addWidget(self.lbl_docinfo)
        root.addWidget(prop)
        self.update_color_btn()

    # ---------------- 檔案操作 ----------------
    def on_files_dropped(self, paths):
        if self.confirm_discard():
            self.load(paths[0])

    def pick(self):
        if not self.confirm_discard():
            return
        p, _ = QFileDialog.getOpenFileName(self, "開啟 PDF", "", "PDF (*.pdf)")
        if p:
            self.load(p)

    def confirm_discard(self):
        if self.dirty:
            r = QMessageBox.question(self, "尚未儲存", "目前的修改尚未儲存，確定要放棄嗎？")
            return r == QMessageBox.StandardButton.Yes
        return True

    def load(self, path):
        """開啟 PDF。讀入記憶體後開啟，避免存回同一檔案時檔案被鎖住"""
        if self.inline and not self.inline.done:
            self.inline.discard()
        self.inline = None
        try:
            with open(path, "rb") as f:
                data = f.read()
            doc = fitz.open("pdf", data)
            if doc.needs_pass:
                pw, ok = QInputDialog.getText(self, "需要密碼", "請輸入 PDF 密碼：",
                                              QLineEdit.EchoMode.Password)
                if not ok or not doc.authenticate(pw):
                    raise RuntimeError("密碼錯誤")
        except Exception as e:
            return self.error(f"無法開啟 PDF：{e}")
        if self.doc:
            self.doc.close()
        self.doc, self.path = doc, path
        self.signed_count = pdf_signature_count(doc)
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.dirty = False
        self.doc_font_choice = None
        self.page_no = 0
        self.after_doc_change()
        self.status(f"已開啟 {path}")

    def after_doc_change(self):
        """文件整份換掉（開啟 / 復原 / 重做）後：更新頁數、縮圖與畫面"""
        n = len(self.doc)
        self.page_no = min(self.page_no, n - 1)
        self.page_spin.blockSignals(True)
        self.page_spin.setMaximum(n)
        self.page_spin.setValue(self.page_no + 1)
        self.page_spin.blockSignals(False)
        self.lbl_pages.setText(f" / {n} ")
        self.refresh_thumbs()
        self.render()

    def write_to(self, path):
        """存檔：先輸出成位元組再寫入，可安全覆蓋原檔"""
        if getattr(self, "signed_count", 0):
            if not self.dirty:
                # 已簽章且沒修改：重新輸出會破壞簽章，改用原檔位元組直接複製
                if os.path.abspath(path) != os.path.abspath(self.path or ""):
                    import shutil
                    shutil.copyfile(self.path, path)
                    self.path = path
                    self.statusChanged.emit()
                self.status(f"已儲存（保留數位簽章）：{path}")
                return True
            r = QMessageBox.warning(self, "文件含數位簽章",
                                    f"此文件有 {self.signed_count} 個數位簽章。儲存修改後，原有簽章會顯示「無效」。\n\n仍要儲存嗎？",
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                    QMessageBox.StandardButton.No)
            if r != QMessageBox.StandardButton.Yes:
                self.status("已取消儲存")
                return False
        data = self.doc.tobytes(garbage=3, deflate=True)
        with open(path, "wb") as f:
            f.write(data)
        self.path = path
        self.dirty = False
        self.statusChanged.emit()
        self.status(f"已儲存：{path}")

    def save(self):
        if not self.doc:
            return self.status("請先開啟 PDF")
        try:
            self.write_to(self.path)
        except Exception as e:
            self.error(f"儲存失敗（檔案可能被其他程式開啟）：{e}")

    def digital_sign(self):
        """以 .pfx 憑證數位簽署目前文件，另存新檔"""
        if not self.doc:
            return self.status("請先開啟 PDF")
        self.commit_inline()
        if self.dirty or not self.path:
            r = QMessageBox.question(self, "數位簽章", "簽署前需要先儲存目前的修改。要現在儲存嗎？")
            if r != QMessageBox.StandardButton.Yes:
                return
            if self.path:
                self.save()
            else:
                self.save_as()
            if self.dirty or not self.path:
                return
        base, _ = os.path.splitext(self.path)
        dlg = DigitalSignDialog(self, unique_path(base + "_已簽章.pdf"))
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        out = dlg.ed_out.text().strip()
        if not out.lower().endswith(".pdf"):
            out += ".pdf"
        try:
            with open(self.path, "rb") as f:       # 用磁碟上的原檔簽署：不重新輸出，保留既有簽章
                data = f.read()
            n = pdf_signature_count(self.doc)
            signed = sign_pdf_bytes(data, dlg.signer, dlg.cmb_reason.currentText(), dlg.ed_loc.text(),
                                    field_name=f"Signature{n + 1}")
            with open(out, "wb") as f:
                f.write(signed)
        except Exception as e:
            return self.error(f"數位簽章失敗：{e}")
        info = dlg.cert_info
        box = QMessageBox(QMessageBox.Icon.Information, "數位簽章完成",
                          f"已簽署並另存為：\n{out}\n\n簽署人：{info['name']}\n發行者：{info['issuer']}\n\n"
                          "※ 請勿再修改簽署後的檔案，否則簽章會失效。", parent=self)
        btn_open = box.addButton("用預設程式開啟", QMessageBox.ButtonRole.ActionRole)
        btn_dir = box.addButton("開啟資料夾", QMessageBox.ButtonRole.ActionRole)
        box.addButton("確定", QMessageBox.ButtonRole.AcceptRole)
        box.exec()
        if box.clickedButton() is btn_open:
            QDesktopServices.openUrl(QUrl.fromLocalFile(out))
        elif box.clickedButton() is btn_dir:
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(out)))
        self.status(f"已數位簽署：{out}")

    def print_pdf(self):
        """列印目前文件（含尚未存檔的修改與註解）。每頁依印表機解析度點陣化後縮放置中列印。"""
        if not self.doc:
            return self.status("請先開啟 PDF")
        n = len(self.doc)
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setDocName(os.path.basename(self.path) if self.path else "文件.pdf")
        printer.setFromTo(1, n)
        first = self.doc[0].rect
        printer.setPageOrientation(QPageLayout.Orientation.Landscape if first.width > first.height
                                   else QPageLayout.Orientation.Portrait)
        dlg = QPrintDialog(printer, self)
        dlg.setWindowTitle("列印")
        dlg.setMinMax(1, n)
        dlg.setOption(QAbstractPrintDialog.PrintDialogOption.PrintPageRange, True)
        dlg.setOption(QAbstractPrintDialog.PrintDialogOption.PrintCurrentPage, True)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        rng = printer.printRange()
        if rng == QPrinter.PrintRange.PageRange and printer.fromPage() > 0:
            pages = list(range(printer.fromPage() - 1, min(printer.toPage(), n)))
        elif rng == QPrinter.PrintRange.CurrentPage:
            pages = [self.page_no]
        else:
            pages = list(range(n))
        if not pages:
            return
        dpi = max(150, min(printer.resolution(), 300))      # 150~300 dpi：清晰度與速度兼顧
        prog = QProgressDialog("正在送出列印…", "取消", 0, len(pages), self)
        prog.setWindowTitle("列印")
        prog.setWindowModality(Qt.WindowModality.WindowModal)
        prog.setMinimumDuration(300)
        painter = QPainter()
        sent = 0
        try:
            if not painter.begin(printer):
                return self.error("無法連線到印表機，請確認印表機已開啟。")
            for i, pno in enumerate(pages):
                if prog.wasCanceled():
                    break
                prog.setValue(i)
                QApplication.processEvents()
                if i > 0:
                    printer.newPage()
                pix = self.doc[pno].get_pixmap(dpi=dpi, annots=True, alpha=False)
                img = QImage(pix.samples, pix.width, pix.height, pix.stride,
                             QImage.Format.Format_RGB888).copy()
                area = painter.viewport()
                # 自動旋轉：頁面與紙張方向不同時轉 90°，避免橫式頁被縮得很小
                if (img.width() > img.height()) != (area.width() > area.height()):
                    img = img.transformed(QTransform().rotate(90))
                tgt = img.size().scaled(area.size(), Qt.AspectRatioMode.KeepAspectRatio)
                x = area.x() + (area.width() - tgt.width()) // 2
                y = area.y() + (area.height() - tgt.height()) // 2
                painter.drawImage(QRect(x, y, tgt.width(), tgt.height()), img)
                sent += 1
            prog.setValue(len(pages))
        except Exception as e:
            self.error(f"列印失敗：{e}")
        finally:
            if painter.isActive():
                painter.end()
            prog.close()
        if sent:
            self.status(f"已送出列印（{sent} 頁）")

    def save_as(self):
        if not self.doc:
            return self.status("請先開啟 PDF")
        p, _ = QFileDialog.getSaveFileName(self, "另存新檔", self.path or "", "PDF (*.pdf)")
        if p:
            try:
                self.write_to(p)
            except Exception as e:
                self.error(f"儲存失敗：{e}")

    # ---------------- 復原 / 重做（快照堆疊） ----------------
    def snapshot(self):
        """每次修改前把整份文件存成位元組放入堆疊，復原時重新載入"""
        self.undo_stack.append(self.doc.tobytes())
        if len(self.undo_stack) > self.MAX_UNDO:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def _swap(self, src, dst, msg):
        if self.inline and not self.inline.done:
            self.inline.discard()
        if not self.doc or not src:
            return self.status("沒有可" + msg + "的動作")
        dst.append(self.doc.tobytes())
        data = src.pop()
        self.doc.close()
        self.doc = fitz.open("pdf", data)
        self.dirty = True
        self.after_doc_change()
        self.status("已" + msg)

    def undo(self):
        self._swap(self.undo_stack, self.redo_stack, "復原")

    def redo(self):
        self._swap(self.redo_stack, self.undo_stack, "重做")

    def modify(self, func, *args):
        """包裝所有編輯動作：快照 → 執行 → 重繪；失敗時自動還原"""
        if not self.doc:
            return False
        self.snapshot()
        try:
            func(self.doc[self.page_no], *args)
        except Exception as e:
            data = self.undo_stack.pop()
            self.doc.close()
            self.doc = fitz.open("pdf", data)
            self.render()
            self.error(f"編輯失敗：{e}")
            return False
        self.dirty = True
        self.render()
        self.update_thumb(self.page_no)
        return True

    # ---------------- 就地編輯 / 進階編輯 ----------------
    def start_inline(self, span):
        """在畫布上直接編輯點選的文字（Acrobat 式就地編輯）"""
        self.commit_inline()
        self.inline = InlineEditor(self, span)
        self.status("就地編輯：Enter 套用、Esc 取消、Tab 下一行；右鍵 →「進階編輯…」")
        return self.inline

    def commit_inline(self):
        if self.inline and not self.inline.done:
            self.inline.apply()
        self.inline = None
        if getattr(self, "sign_overlay", None):     # 換頁 / 縮放 / 換工具 / 存檔前先確定簽名位置
            self.sign_overlay.commit()

    def apply_span_edit(self, span, text, choice, size, color, is_ocr, bold=None, italic=None, underline=False):
        """套用一次文字修改（就地編輯與進階對話框共用）"""
        if choice and choice != "original":
            settings().setValue("last_font", choice)       # 記住上次選擇
            if is_ocr:
                self.doc_font_choice = choice               # 本文件後續 OCR 修改沿用
        info = {}
        self.modify(lambda pg, sp, t: info.update(
            replace_span_text(pg, sp, t, fontfile=choice, size=size, color=color,
                              bold=bold, italic=italic, underline=underline) or {}), span, text)
        if info:
            self.status(f"已修改文字（{info.get('label', '').splitlines()[0]}，{info.get('size', 0):.1f} pt）")

    def edit_next_after(self, span):
        """Tab：編輯下一行文字（依閱讀順序）"""
        page = self.doc[self.page_no]
        spans = [s for b in page.get_text("dict")["blocks"] if b.get("type") == 0
                 for l in b["lines"] for s in l["spans"] if s["text"].strip()]
        spans.sort(key=lambda s: (round(s["bbox"][1] / 4), s["bbox"][0]))
        y0, x0 = span["bbox"][1], span["bbox"][0]
        for s2 in spans:
            if (round(s2["bbox"][1] / 4), s2["bbox"][0]) > (round(y0 / 4), x0 + 0.5):
                return self.start_inline(dict(s2))

    def advanced_edit(self, span):
        """原本的置中對話框（右鍵 →「進階編輯…」）"""
        page = self.doc[self.page_no]
        is_ocr = is_invisible_span(page, span)
        pref = (self.doc_font_choice or settings().value("last_font", None)) if is_ocr else None
        dlg = EditSpanDialog(page, span, self.main, preferred_font=pref)
        if dlg.exec() and (dlg.text() != span["text"] or dlg.font_choice() or dlg.size() or dlg.style_changed()):
            b_, i_, u_ = dlg.style()
            info = {}
            choice = dlg.font_choice()
            if choice and choice != dlg.ORIGINAL:
                settings().setValue("last_font", choice)
                if is_ocr:
                    self.doc_font_choice = choice
            self.modify(lambda pg, sp, t: info.update(
                replace_span_text(pg, sp, t, fontfile=choice, shrink=dlg.shrink(), size=dlg.size(),
                                  bold=b_, italic=i_, underline=u_) or {}),
                span, dlg.text())

    def on_context(self, qp, gpos):
        """右鍵選單：在文字上提供「進階編輯…」"""
        if not self.doc:
            return
        self.commit_inline()
        page = self.doc[self.page_no]
        span = find_span_at(page, self.to_pdf(qp))
        menu = QMenu(self)
        act_edit = menu.addAction(make_icon("select"), "就地編輯")
        act_adv = menu.addAction(make_icon("edit"), "進階編輯…")
        if not span:
            act_edit.setEnabled(False)
            act_adv.setEnabled(False)
        chosen = menu.exec(gpos)
        if chosen == act_edit:
            self.start_inline(span)
        elif chosen == act_adv:
            self.advanced_edit(span)

    # ---------------- OCR ----------------
    def on_ocr_font_changed(self, *_, save=True):
        """屬性面板的 OCR 替換字型（以 QSettings 保存）"""
        global OCR_DEFAULT_FONT
        path = self.cmb_ocr_font.currentData() or None
        OCR_DEFAULT_FONT = path
        if save:
            settings().setValue("ocr_font", path or "")

    def run_ocr(self, all_pages=None):
        """在背景執行緒辨識目前頁或全部頁面，完成後加入隱形文字層（可復原）"""
        if not self.doc:
            return self.status("請先開啟 PDF")
        if all_pages is None:
            all_pages = self.rb_ocr_all.isChecked()
        pages = list(range(len(self.doc))) if all_pages else [self.page_no]
        if all(self.doc[i].get_text("text").strip() for i in pages):
            return self.show_message("文字辨識", "此頁已有文字，可直接編輯。" if len(pages) == 1
                                     else "這些頁面都已有文字，可直接編輯。")
        self.status("正在辨識文字（第一次需要載入模型，請稍候）…", 0)
        self.run_worker(ocr_document_bytes, self.doc.tobytes(), pages,
                        on_done=self.ocr_done, busy_widgets=[self.btn_ocr])

    def ocr_done(self, res):
        data, n, done_pages, skipped = res
        if done_pages:
            self.undo_stack.append(self.doc.tobytes())
            self.redo_stack.clear()
            self.doc.close()
            self.doc = fitz.open("pdf", data)
            self.dirty = True
            self.after_doc_change()
        self.progress.setValue(100)
        # 組合完成訊息
        if done_pages and n:
            msg = f"文字辨識完成！共辨識 {len(done_pages)} 頁、{n} 行文字。\n現在可以使用『修改文字』工具點選文字進行編輯。"
            icon = QMessageBox.Icon.Information
        elif done_pages:
            msg = f"已辨識 {len(done_pages)} 頁，但沒有找到可辨識的文字。\n請確認頁面是否為清楚的掃描影像。"
            icon = QMessageBox.Icon.Warning
        else:
            msg = "此頁已有文字，可直接編輯。" if len(skipped) == 1 else f"這 {len(skipped)} 頁都已有文字，可直接編輯。"
            icon = QMessageBox.Icon.Information
        if done_pages and skipped:
            msg += f"\n（另有 {len(skipped)} 頁原本已有文字，已略過）"
        self.status(msg.splitlines()[0])
        self.show_message("文字辨識", msg, icon)
        self.set_tool("select")          # 按下確定後自動切換到「修改文字」

    def show_message(self, title, text, icon=QMessageBox.Icon.Information):
        """置中於主視窗的強制回應訊息框"""
        box = QMessageBox(icon, title, text, QMessageBox.StandardButton.Ok, self.main)
        box.button(QMessageBox.StandardButton.Ok).setText("確定")
        box.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.last_message = text
        box.exec()

    # ---------------- 檢視 ----------------
    def thumb_pixmap(self, i):
        page = self.doc[i]
        z = 140 / max(page.rect.width, page.rect.height)
        return fitz_pixmap_to_qpixmap(page.get_pixmap(matrix=fitz.Matrix(z, z), annots=True))

    def refresh_thumbs(self):
        self.thumbs.blockSignals(True)
        self.thumbs.clear()
        for i in range(len(self.doc)):
            it = QListWidgetItem(thumb_icon(self.thumb_pixmap(i)), str(i + 1))
            it.setTextAlignment(Qt.AlignmentFlag.AlignHCenter)
            self.thumbs.addItem(it)
            if i % 10 == 0:
                QApplication.processEvents()
        self.thumbs.setCurrentRow(self.page_no)
        self.thumbs.blockSignals(False)

    def rotate_page(self, delta, pno=None):
        """旋轉單一頁面（預設為目前頁面），可復原"""
        if not self.doc:
            return self.status("請先開啟 PDF")
        if pno is not None and pno != self.page_no:
            self.goto(pno)
        self.commit_inline()
        if self.modify(lambda pg: pg.set_rotation((pg.rotation + delta) % 360)):
            self.status(f"第 {self.page_no + 1} 頁已{'向右' if delta > 0 else '向左'}旋轉 90°")

    def delete_page(self, pno):
        """刪除單一頁面，可復原"""
        if not self.doc:
            return
        if len(self.doc) <= 1:
            return self.status("文件只剩一頁，無法刪除")
        if QMessageBox.question(self, "刪除頁面", f"確定要刪除第 {pno + 1} 頁嗎？（可用 Ctrl+Z 復原）") \
                != QMessageBox.StandardButton.Yes:
            return
        self.commit_inline()
        self.snapshot()
        self.doc.delete_page(pno)
        self.dirty = True
        self.page_no = min(pno, len(self.doc) - 1)
        self.after_doc_change()
        self.status(f"已刪除第 {pno + 1} 頁")

    def thumb_menu(self, pos):
        it = self.thumbs.itemAt(pos)
        if not it or not self.doc:
            return
        pno = self.thumbs.row(it)
        menu = QMenu(self)
        a_l = menu.addAction(make_icon("rotl"), "向左旋轉 90°")
        a_r = menu.addAction(make_icon("rotr"), "向右旋轉 90°")
        menu.addSeparator()
        a_d = menu.addAction(make_icon("delannot"), "刪除此頁")
        act = menu.exec(self.thumbs.viewport().mapToGlobal(pos))
        if act is a_l:
            self.rotate_page(-90, pno)
        elif act is a_r:
            self.rotate_page(90, pno)
        elif act is a_d:
            self.delete_page(pno)

    def update_thumb(self, i):
        it = self.thumbs.item(i)
        if it:
            it.setIcon(thumb_icon(self.thumb_pixmap(i)))

    def toggle_thumbs(self, show=None):
        show = not self.thumb_panel.isVisible() if show is None else show
        self.thumb_panel.setVisible(show)

    def render(self):
        if not self.doc:
            return
        page = self.doc[self.page_no]
        pix = fitz_pixmap_to_qpixmap(page.get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom), annots=True))
        if self.canvas.mode == "link":
            self.draw_link_boxes(page, pix)
        self.canvas.setPixmap(pix)
        self.canvas.setFixedSize(pix.size())
        self.lbl_zoom.setText(f"{int(self.zoom * 100)}%")
        self.thumbs.blockSignals(True)
        self.thumbs.setCurrentRow(self.page_no)
        self.thumbs.blockSignals(False)
        name = os.path.basename(self.path or "")
        self.lbl_docinfo.setText(f"檔案：{name}\n頁數：{len(self.doc)}\n頁面尺寸：{page.rect.width:.0f} × {page.rect.height:.0f} pt"
                                 + ("\n（尚未儲存）" if self.dirty else ""))
        self.statusChanged.emit()

    def draw_link_boxes(self, page, pix):
        """超連結工具：在畫面上用藍色虛線框標出已有的連結（只顯示，不寫入 PDF）"""
        links = read_links(page)
        if not links:
            return
        m = page.rotation_matrix * fitz.Matrix(self.zoom, self.zoom)
        p = QPainter(pix)
        p.setPen(QPen(QColor("#0B6BCB"), 1.5, Qt.PenStyle.DashLine))
        p.setBrush(QColor(11, 107, 203, 28))
        for ln in links:
            r = ln["rect"] * m
            p.drawRect(QRectF(r.x0, r.y0, r.width, r.height))
        p.end()

    def goto(self, n):
        self.commit_inline()               # 換頁前先套用就地編輯
        if not self.doc or not (0 <= n < len(self.doc)):
            return
        self.page_no = n
        self.page_spin.blockSignals(True)
        self.page_spin.setValue(n + 1)
        self.page_spin.blockSignals(False)
        self.render()

    def set_zoom(self, z):
        self.commit_inline()               # 縮放前先套用就地編輯
        self.zoom = max(0.25, min(z, 6.0))
        self.render()

    def fit_width(self):
        if self.doc:
            w = self.doc[self.page_no].rect.width
            self.set_zoom((self.scroll.viewport().width() - 90) / w)

    def fit_page(self):
        if self.doc:
            r = self.doc[self.page_no].rect
            vp = self.scroll.viewport()
            self.set_zoom(min((vp.width() - 90) / r.width, (vp.height() - 70) / r.height))

    def to_pdf(self, qp):
        """畫面像素座標 → PDF 頁面座標（考慮縮放與頁面旋轉）"""
        page = self.doc[self.page_no]
        return fitz.Point(qp.x() / self.zoom, qp.y() / self.zoom) * page.derotation_matrix

    # ---------------- 工具 ----------------
    def set_tool(self, key):
        self.commit_inline()
        self.canvas.mode = key
        info = {k: (n, t) for k, n, t in self.TOOLS}[key]
        self.lbl_tool.setText(info[0])
        self.hint.setText(info[1])
        if not self.tool_actions[key].isChecked():
            self.tool_actions[key].setChecked(True)
        self.generic_box.setVisible(key not in ("sign", "check"))
        self.check_box.setVisible(key == "check")
        self.sign_box.setVisible(key == "sign")
        self.ocr_box.setVisible(key != "sign")
        self.render()                      # 切換到 / 離開超連結工具時更新連結外框

    def update_color_btn(self):
        self.btn_color.setText(self.color.name().upper())
        fg = "#000" if self.color.lightness() > 150 else "#FFF"
        self.btn_color.setStyleSheet(f"QPushButton{{background:{self.color.name()};color:{fg};border-color:#999;}}")

    def pick_color(self):
        c = QColorDialog.getColor(self.color, self, "選擇顏色")
        if c.isValid():
            self.color = c
            self.update_color_btn()

    def style_annot(self, an, fill=False):
        """依右側屬性面板設定註解的顏色、線寬、不透明度"""
        an.set_colors(stroke=color_to_rgb01(self.color))
        an.set_border(width=self.spin_width.value())
        an.set_opacity(self.opacity.value() / 100)
        an.update()

    def on_click(self, qp):
        if not self.doc:
            return
        pt = self.to_pdf(qp)
        page = self.doc[self.page_no]
        mode = self.canvas.mode
        rgb = color_to_rgb01(self.color)
        if mode == "select":
            if self.inline and not self.inline.done:      # 點其他地方 = 套用目前的就地編輯
                self.commit_inline()
                page = self.doc[self.page_no]
            span = find_span_at(page, pt)
            if not span:
                return self.status("此處沒有可編輯的文字")
            self.start_inline(span)
        elif mode == "text":
            text, ok = QInputDialog.getMultiLineText(self, "新增文字", "請輸入文字：")
            if ok and text:
                self.modify(add_text_box, pt, text, self.spin_font.value(), rgb)
        elif mode == "image":
            p, _ = QFileDialog.getOpenFileName(self, "選擇圖片", "", "圖片 (*.png *.jpg *.jpeg)")
            if p:
                self.modify(insert_image_at, pt, p, self.spin_imgw.value())
        elif mode == "note":
            text, ok = QInputDialog.getMultiLineText(self, "便利貼", "註解內容：")
            if ok and text:
                def add_note(pg, point, t):
                    a = pg.add_text_annot(point, t, icon="Note")
                    a.set_info(title=os.environ.get("USERNAME", "使用者"))
                    a.update()
                self.modify(add_note, pt, text)
        elif mode == "check":
            sym = self.cmb_check_sym.currentData()
            rgb = tuple(self.cmb_check_color.currentData())
            box = box_at(page, pt)
            if box is None:                            # 沒有框框（掃描檔等）：點哪裡就勾哪裡
                existing = [a for a in (page.annots() or []) if a.info.get("title") == CHECK_TAG
                            and a.rect.contains(pt)]
                if existing:
                    return self.modify(lambda pg: [pg.delete_annot(a) for a in
                                                   [x for x in (pg.annots() or []) if x.info.get("title") == CHECK_TAG
                                                    and x.rect.contains(pt)]])
                s_ = self.spin_check_size.value()
                vp = pt * page.rotation_matrix
                box = (fitz.Rect(vp.x - s_ / 2, vp.y - s_ / 2, vp.x + s_ / 2, vp.y + s_ / 2)
                       * page.derotation_matrix).normalize()
            res = {}
            self.modify(lambda pg, b: res.update(on=toggle_check(pg, b, sym, rgb)), box)
            self.status("已打勾" if res.get("on") else "已取消勾選")
        elif mode == "sign":
            if self.sign_overlay:                      # 正在調整時點旁邊 = 確定
                return self.sign_overlay.commit()
            hit = signature_at(page, pt)
            if hit:
                return self.edit_placed_signature(*hit)
            path = self.current_signature()
            if not path:
                return self.status("請先在右側選擇或新增簽名 / 印章")
            with Image.open(path) as im:
                iw, ih = im.size
            w_ = float(self.spin_signw.value()) * self.zoom
            h_ = w_ * ih / max(iw, 1)
            self.start_sign_overlay(path, QRectF(qp.x() - w_ / 2, qp.y() - h_ / 2, w_, h_))
        elif mode == "link":
            ln = link_at(page, pt)
            if ln:
                return self.edit_link(ln)
            hit = find_line_at(page, pt)
            if not hit:
                return self.status("此處沒有文字：請拖曳框選要加連結的範圍")
            self.new_link(hit[0], hit[1])
        elif mode == "delannot":
            if not any(a.rect.contains(pt) for a in (page.annots() or [])):
                return self.status("此處沒有註解")
            self.modify(delete_annot_at, pt)

    def on_drag(self, a, b):
        if not self.doc:
            return
        p1, p2 = self.to_pdf(a), self.to_pdf(b)
        rect = fitz.Rect(p1, p2).normalize()
        mode = self.canvas.mode
        if mode == "highlight":
            def add_hl(pg, r):
                an = highlight_area(pg, r, color_to_rgb01(self.color))
                an.set_opacity(self.opacity.value() / 100)
                an.update()
            self.modify(add_hl, rect)
        elif mode == "rect":
            self.modify(lambda pg, r: self.style_annot(pg.add_rect_annot(r)), rect)
        elif mode == "line":
            self.modify(lambda pg, x, y: self.style_annot(pg.add_line_annot(x, y)), p1, p2)
        elif mode == "check":
            page = self.doc[self.page_no]
            if not find_boxes(page, rect):
                return self.status("框選範圍內沒有偵測到框框")
            sym = self.cmb_check_sym.currentData()
            rgb = tuple(self.cmb_check_color.currentData())
            res = {}
            self.modify(lambda pg, r: res.update(n=check_all_in(pg, r, sym, rgb)), rect)
            self.status(f"已勾選 {res.get('n', 0)} 個框框")
        elif mode == "sign":
            if self.sign_overlay:
                return self.sign_overlay.commit()
            path = self.current_signature()
            if not path:
                return self.status("請先在右側選擇或新增簽名 / 印章")
            with Image.open(path) as im:
                aspect = im.width / max(im.height, 1)
            box = QRectF(a, b).normalized()
            w_ = min(box.width(), box.height() * aspect)       # 放進框內、保持比例
            h_ = w_ / aspect
            self.start_sign_overlay(path, QRectF(box.center().x() - w_ / 2, box.center().y() - h_ / 2, w_, h_))
        elif mode == "link":
            page = self.doc[self.page_no]
            text = page.get_textbox(rect).strip().replace("\n", " ")
            self.new_link(rect, text or None)
        elif mode == "whiteout":
            r = QMessageBox.question(self, "塗銷", "將永久刪除框選範圍內的文字與圖片，確定嗎？")
            if r == QMessageBox.StandardButton.Yes:
                self.modify(whiteout_area, rect)

    # ---------------- 簽名 / 印章：放置與調整 ----------------
    def _screen_to_visual(self, r):
        z = self.zoom
        return fitz.Rect(r.x() / z, r.y() / z, r.right() / z, r.bottom() / z)

    def start_sign_overlay(self, path, screen_rect):
        """新蓋一個章：先浮在頁面上讓使用者調整，確定後才寫入"""
        with open(path, "rb") as f:
            png = f.read()
        pm = QPixmap()
        pm.loadFromData(png)
        self.sign_overlay = SignOverlay(
            self.canvas, pm, screen_rect,
            on_commit=lambda r: self._finish_overlay(lambda: self.modify(place_signature, self._screen_to_visual(r), png),
                                                     "已蓋上：點一下章可再調整"),
            on_cancel=lambda: self._finish_overlay(None, "已取消"))
        self.status("拖曳移動、拖四角縮放；Enter 確定、Esc 取消（方向鍵可微調，Shift 加快）")

    def edit_placed_signature(self, xref, bbox):
        """重新調整已蓋的章：先從頁面移除（畫面上改用可調整框顯示），確定後於新位置重蓋"""
        page = self.doc[self.page_no]
        try:
            png = signature_png(self.doc, xref)
        except Exception as e:
            return self.error(f"無法讀取這個章：{e}")
        vis = (bbox * page.rotation_matrix).normalize()
        z = self.zoom
        screen = QRectF(vis.x0 * z, vis.y0 * z, vis.width * z, vis.height * z)
        self.snapshot()                                   # 移除 + 重蓋算同一個復原步驟
        remove_signature(page, xref)
        self.render()
        pm = QPixmap()
        pm.loadFromData(png)

        def commit(r):
            place_signature(self.doc[self.page_no], self._screen_to_visual(r), png)
            self._after_direct_edit("已更新位置")

        def cancel():
            data = self.undo_stack.pop()
            self.doc.close()
            self.doc = fitz.open("pdf", data)
            self.render()
            self.status("已取消調整")

        def delete():
            self._after_direct_edit("已刪除這個章")

        self.sign_overlay = SignOverlay(self.canvas, pm, screen,
                                        on_commit=lambda r: self._finish_overlay(lambda: commit(r)),
                                        on_cancel=lambda: self._finish_overlay(cancel),
                                        on_delete=lambda: self._finish_overlay(delete))
        self.status("拖曳移動、拖四角縮放；Enter 確定、Esc 取消、Delete 刪除")

    def _after_direct_edit(self, msg):
        self.dirty = True
        self.render()
        self.update_thumb(self.page_no)
        self.status(msg)

    def _finish_overlay(self, action, msg=None):
        self.sign_overlay = None
        if action:
            action()
        if msg:
            self.status(msg)

    def commit_sign_overlay(self):
        if self.sign_overlay:
            self.sign_overlay.commit()

    # ---------------- 簽名 / 印章庫 ----------------
    def refresh_signatures(self, select=None):
        self.sign_list.blockSignals(True)
        self.sign_list.clear()
        for path in list_signatures():
            pm = QPixmap(path)
            if pm.isNull():
                continue
            canvas = QPixmap(88, 56)
            canvas.fill(QColor("#FFFFFF"))
            p = QPainter(canvas)
            sc = pm.scaled(84, 52, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            p.drawPixmap((88 - sc.width()) // 2, (56 - sc.height()) // 2, sc)
            p.end()
            it = QListWidgetItem(QIcon(canvas), "")
            it.setData(Qt.ItemDataRole.UserRole, path)
            it.setToolTip(os.path.basename(path))
            self.sign_list.addItem(it)
        self.sign_list.blockSignals(False)
        if self.sign_list.count():
            row = self.sign_list.count() - 1
            if select:
                for i in range(self.sign_list.count()):
                    if self.sign_list.item(i).data(Qt.ItemDataRole.UserRole) == select:
                        row = i
            self.sign_list.setCurrentRow(row)
        self.btn_sign_del.setEnabled(self.sign_list.count() > 0)

    def current_signature(self):
        it = self.sign_list.currentItem()
        return it.data(Qt.ItemDataRole.UserRole) if it else None

    def on_sign_selected(self, cur, _prev=None):
        if not cur:
            return
        path = cur.data(Qt.ItemDataRole.UserRole)
        w_ = settings().value("sign_w/" + os.path.basename(path), None)
        self.spin_signw.blockSignals(True)
        self.spin_signw.setValue(int(w_) if w_ else default_sign_width(path))
        self.spin_signw.blockSignals(False)

    def on_signw_changed(self, v):
        path = self.current_signature()
        if path:
            settings().setValue("sign_w/" + os.path.basename(path), int(v))

    def sign_import(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "匯入簽名 / 印章圖檔", "",
                                                "圖片 (*.png *.jpg *.jpeg *.bmp)")
        last = None
        for p_ in paths:
            try:
                last = import_signature_image(p_, self.chk_sign_bg.isChecked())
            except Exception as e:
                self.error(f"無法匯入 {os.path.basename(p_)}：{e}")
        if last:
            self.refresh_signatures(select=last)
            self.status("已加入簽名 / 印章：在頁面上點一下即可蓋上")

    def sign_draw(self):
        dlg = SignatureDrawDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.refresh_signatures(select=dlg.save())
            self.status("已儲存手寫簽名：在頁面上點一下即可蓋上")

    def sign_delete(self):
        path = self.current_signature()
        if not path:
            return
        if QMessageBox.question(self, "刪除簽名 / 印章", "要從簽名庫刪除這個簽名 / 印章嗎？\n"
                                "（已經蓋在文件上的不受影響）") != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
        except OSError as e:
            return self.error(f"刪除失敗：{e}")
        settings().remove("sign_w/" + os.path.basename(path))
        self.refresh_signatures()

    def new_link(self, rect, text=None):
        dlg = LinkDialog(self, len(self.doc), self.page_no, text=text)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            target = dlg.target()
            if self.modify(add_link, rect, target, dlg.chk_bg.isChecked()):
                self.status(self.link_desc(target, "已新增連結"))

    def edit_link(self, ln):
        page = self.doc[self.page_no]
        text = page.get_textbox(ln["rect"]).strip().replace("\n", " ")
        dlg = LinkDialog(self, len(self.doc), self.page_no, text=text or None, link=ln,
                         bg=has_link_bg(page, ln["rect"]))
        r = dlg.exec()
        if r == LinkDialog.DELETE:
            if self.modify(remove_link, ln):
                self.status("已刪除連結")
        elif r == QDialog.DialogCode.Accepted:
            target = dlg.target()
            if self.modify(replace_link, ln, ln["rect"], target, dlg.chk_bg.isChecked()):
                self.status(self.link_desc(target, "已更新連結"))

    @staticmethod
    def link_desc(target, prefix):
        kind, val = target
        return f"{prefix}：{'第 ' + str(val + 1) + ' 頁' if kind == 'page' else val}"

    def on_stroke(self, points):
        if not self.doc:
            return
        pts = [tuple(self.to_pdf(p)) for p in points]
        self.modify(lambda pg, stroke: self.style_annot(pg.add_ink_annot([stroke])), pts)


# =====================================================================
# 主視窗：頂部標題列 + 主工具列 + 左側工具軌 + 內容區（QStackedWidget）+ 狀態列
# =====================================================================

STARTUP_NOTICE = "本軟體由長庚大學資訊中心 RC 開發"


class StartupNotice(QDialog):
    """每次開啟程式時顯示的開發單位提示"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(APP_NAME)
        self.setModal(True)
        self.setMinimumWidth(380)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 16)
        lay.setSpacing(12)
        top = QHBoxLayout()
        logo = QLabel()
        logo.setPixmap(logo_pixmap(56))
        top.addWidget(logo)
        info = QLabel(f"<div style='font-size:17px;font-weight:bold'>{APP_NAME}</div>"
                      f"<div style='color:{ACCENT_BTN};font-weight:bold'>版本 {APP_VERSION}</div>")
        top.addWidget(info, 1)
        lay.addLayout(top)
        msg = QLabel(STARTUP_NOTICE)
        msg.setStyleSheet("font-size:15px; padding:6px 0;")
        msg.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(msg)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("確定")
        bb.button(QDialogButtonBox.StandardButton.Ok).setProperty("primary", True)
        bb.accepted.connect(self.accept)
        lay.addWidget(bb, 0, Qt.AlignmentFlag.AlignCenter)


class AboutDialog(QDialog):
    """「關於」對話框：名稱、版本、作者、著作權與開放原始碼授權聲明"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("關於 " + APP_NAME)
        self.setModal(True)
        self.setMinimumWidth(500)
        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        top = QHBoxLayout()
        logo = QLabel()
        logo.setPixmap(logo_pixmap(64))
        top.addWidget(logo)
        info = QLabel(f"<div style='font-size:18px;font-weight:bold'>{APP_NAME}</div>"
                      f"<div style='color:{ACCENT_BTN};font-weight:bold'>版本 {APP_VERSION}</div>"
                      f"<div>Created by {APP_AUTHOR}</div>"
                      f"<div>以 {APP_LICENSE} 授權開源發布，可自由使用、修改與散布</div>"
                      + (f"<div>原始碼：<a href='{APP_SOURCE_URL}'>{APP_SOURCE_URL}</a></div>" if APP_SOURCE_URL else "")
                      + f"<div style='color:#777'>Copyright © {APP_YEAR} {APP_AUTHOR}</div>")
        info.setOpenExternalLinks(True)
        top.addWidget(info, 1)
        lay.addLayout(top)
        lic = QLabel(
            "<b>開放原始碼元件與授權</b><br>"
            "本程式使用下列開放原始碼軟體，各元件依其授權條款散布：<ul style='margin-left:-20px'>"
            "<li>PyQt6 / Qt 6 — GPL v3（Riverbank Computing / The Qt Company）</li>"
            "<li>PyMuPDF / MuPDF — AGPL v3（Artifex Software）</li>"
            "<li>RapidOCR、ONNX Runtime — Apache 2.0；PaddleOCR 模型 — Apache 2.0</li>"
            "<li>OpenCV — Apache 2.0；NumPy — BSD</li>"
            "<li>pypdf — BSD 3-Clause</li>"
            "<li>img2pdf — LGPL v3</li>"
            "<li>pdf2docx — MIT License</li><li>python-docx — MIT License</li>"
            "<li>Pillow — MIT-CMU (HPND)</li>"
            "<li>docx2pdf — MIT；pywin32 — PSF</li>"
            "<li>pyHanko — MIT（數位簽章）</li>"
            "<li>Lucide 圖示 — ISC License（© Lucide Icons and Contributors）</li></ul>"
            "<span style='color:#777'>本程式以 AGPL-3.0 授權發布，散布時應一併提供原始碼。"
            "長庚大學校徽著作權屬長庚大學所有。</span>")
        lic.setWordWrap(True)
        lic.setObjectName("PanelHint")
        lay.addWidget(lic)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("確定")
        bb.button(QDialogButtonBox.StandardButton.Ok).setProperty("primary", True)
        bb.accepted.connect(self.accept)
        lay.addWidget(bb)

    def showEvent(self, e):
        self.adjustSize()
        par = self.parentWidget()
        if par:
            self.setGeometry(centered_rect(par.window().frameGeometry(), self.sizeHint().expandedTo(self.minimumSize())))
        super().showEvent(e)


def tool_page(title, desc, widget):
    """把一般工具（合併 / 組織 / 轉換 / 浮水印）包成：標題 + 說明 + 白色卡片"""
    page = QWidget()
    page.setObjectName("ToolPage")
    lay = QVBoxLayout(page)
    lay.setContentsMargins(28, 20, 28, 20)
    lay.setSpacing(6)
    t = QLabel(title)
    t.setObjectName("ToolTitle")
    d = QLabel(desc)
    d.setObjectName("ToolDesc")
    lay.addWidget(t)
    lay.addWidget(d)
    card = QFrame()
    card.setObjectName("Card")
    cl = QVBoxLayout(card)
    cl.setContentsMargins(18, 16, 18, 16)
    cl.addWidget(widget)
    lay.addSpacing(8)
    lay.addWidget(card, 1)
    return page


class MainWindow(QMainWindow):
    # 左側工具軌：(圖示, 名稱, 頁面標題說明)
    RAIL = [
        ("edit", "編輯 PDF", "直接在頁面上修改文字、加入圖片、標記與註解"),
        ("merge", "合併檔案", "將多個 PDF 依序合併成一個檔案"),
        ("organize", "組織頁面", "合併多個檔案、插入、刪除、旋轉、重新排列頁面，一次匯出新的 PDF"),
        ("convert", "格式轉換", "Word（.docx）/ 圖片轉 PDF，或將 PDF 轉為可編輯的 Word"),
        ("watermark", "文字浮水印", "在所有頁面中央壓印半透明文字"),
    ]

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(app_window_icon())
        self.resize(1280, 820)
        self.setStatusBar(QStatusBar())

        self.convert_tab = ConvertTab(self)
        self.merge_tab = MergeTab(self)
        self.organize_tab = OrganizeTab(self)
        self.watermark_tab = WatermarkTab(self)
        self.edit_tab = EditTab(self)
        # 主要動作按鈕使用紅色強調樣式
        c, m, o, w = self.convert_tab, self.merge_tab, self.organize_tab, self.watermark_tab
        for b, v in ((c.btn_go, "primary"), (c.btn_add, "outline-primary"), (c.btn_remove, "outline-danger"),
                     (c.btn_clear, "outline-danger"), (c.btn_browse, "outline-secondary"),
                     (m.btn_go, "success"), (m.btn_add, "outline-primary"), (m.btn_up, "outline-secondary"),
                     (m.btn_down, "outline-secondary"), (m.btn_remove, "outline-danger"), (m.btn_clear, "outline-danger"),
                     (o.btn_open, "primary"), (o.btn_export, "success"), (o.btn_del, "danger"),
                     (o.btn_left, "outline-secondary"), (o.btn_right, "outline-secondary"),
                     (o.btn_fwd, "outline-secondary"), (o.btn_back, "outline-secondary"),
                     (w.btn_go, "success"), (w.btn_src, "outline-secondary"), (self.edit_tab.btn_ocr, "primary")):
            bs(b, v)

        root = QWidget()
        root.setObjectName("Body")
        v = QVBoxLayout(root)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        v.addWidget(self.build_header())
        v.addWidget(self.build_toolbar())

        body = QHBoxLayout()
        body.setSpacing(0)
        body.addWidget(self.build_rail())
        self.stack = QStackedWidget()
        self.stack.addWidget(self.edit_tab)
        pages = [self.merge_tab, self.organize_tab, self.convert_tab, self.watermark_tab]
        for (icon, name, desc), w in zip(self.RAIL[1:], pages):
            self.stack.addWidget(tool_page(name, desc, w))
        body.addWidget(self.stack, 1)
        v.addLayout(body, 1)
        self.setCentralWidget(root)

        self.lbl_state = QLabel()
        self.statusBar().addPermanentWidget(self.lbl_state)
        self.edit_tab.statusChanged.connect(self.update_state)
        self.statusBar().showMessage("就緒")
        self.set_view(0)
        self.update_state()

    # ---- 頂部標題列：程式圖示 + 應用程式名稱 + 目前檔名 ----
    def build_header(self):
        """精簡標題列（約 34px）：程式圖示 + 程式名稱 + 版本 | 檔名 · 著作權 · 關於"""
        h = QFrame()
        h.setObjectName("Header")
        h.setFixedHeight(34)
        lay = QHBoxLayout(h)
        lay.setContentsMargins(10, 2, 8, 2)
        lay.setSpacing(8)
        logo = QLabel()
        logo.setPixmap(logo_pixmap(22))
        lay.addWidget(logo)
        name = QLabel(APP_NAME)
        name.setObjectName("AppName")
        name.setStyleSheet("color: #FFFFFF; font-size: 14px; font-weight: 600;")   # 深色標題列白字
        lay.addWidget(name)
        badge = QLabel(APP_VERSION)          # 版本標籤
        badge.setObjectName("VerBadge")
        lay.addWidget(badge, 0, Qt.AlignmentFlag.AlignVCenter)
        lay.addStretch()
        self.lbl_doc = QLabel("")
        self.lbl_doc.setObjectName("DocName")
        lay.addWidget(self.lbl_doc)
        lay.addSpacing(12)
        copy = QLabel(f"Created by {APP_AUTHOR} © {APP_YEAR}")   # 著作權標示
        copy.setObjectName("Copyright")
        copy.setStyleSheet("color: #A8A8A8; font-size: 11px;")
        lay.addWidget(copy)
        about = QToolButton()
        about.setObjectName("AboutBtn")
        about.setText("關於")
        about.setIcon(make_icon("about", color="#CCCCCC", size=15, active="#FFFFFF", hover="#FFFFFF"))
        about.setIconSize(QSize(15, 15))
        about.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        about.setToolTip("關於本程式、版本與授權資訊")
        about.setCursor(Qt.CursorShape.PointingHandCursor)
        about.clicked.connect(self.show_about)
        self.btn_about = about
        lay.addWidget(about)
        return h

    def show_about(self):
        AboutDialog(self).exec()

    # ---- 主工具列：開啟、儲存、復原/重做、換頁、縮放 ----
    def build_toolbar(self):
        e = self.edit_tab
        tb = QToolBar()
        tb.setObjectName("MainToolbar")
        tb.setIconSize(QSize(20, 20))
        tb.setMovable(False)

        def act(icon, tip, fn, shortcut=None, text=None):
            a = QAction(make_icon(icon, tool=(icon == "ocr")), text or tip, self)
            a.setToolTip(tip + (f"（{QKeySequence(shortcut).toString()}）" if shortcut else ""))
            if shortcut:
                a.setShortcut(shortcut)
            a.triggered.connect(lambda: (self.set_view(0), fn()))
            tb.addAction(a)
            return a

        self.act_thumbs = QAction(make_icon("thumbs"), "頁面縮圖", self, checkable=True, checked=True)
        self.act_thumbs.setToolTip("顯示 / 隱藏頁面縮圖")
        self.act_thumbs.toggled.connect(lambda on: (self.set_view(0), e.toggle_thumbs(on)))
        tb.addAction(self.act_thumbs)
        tb.addSeparator()
        self.act_open = act("open", "開啟", e.pick, QKeySequence.StandardKey.Open)
        self.act_save = act("save", "儲存", e.save, QKeySequence.StandardKey.Save)
        self.act_saveas = act("saveas", "另存新檔", e.save_as, QKeySequence.StandardKey.SaveAs)
        self.act_print = act("print", "列印", e.print_pdf, QKeySequence.StandardKey.Print)
        self.act_dsign = act("dsign", "數位簽章（.pfx 憑證）", e.digital_sign, text="數位簽章")
        tb.widgetForAction(self.act_dsign).setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        tb.addSeparator()
        self.act_undo = act("undo", "復原", e.undo, QKeySequence.StandardKey.Undo)
        self.act_redo = act("redo", "重做", e.redo, QKeySequence.StandardKey.Redo)
        tb.addSeparator()
        self.act_ocr = act("ocr", "辨識文字 (OCR)", lambda: e.run_ocr(), text="辨識文字 (OCR)")
        tb.widgetForAction(self.act_ocr).setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        tb.addSeparator()
        act("prev", "上一頁", lambda: e.goto(e.page_no - 1), QKeySequence(Qt.Key.Key_PageUp))
        tb.addWidget(e.page_spin)
        tb.addWidget(e.lbl_pages)
        act("next", "下一頁", lambda: e.goto(e.page_no + 1), QKeySequence(Qt.Key.Key_PageDown))
        tb.addSeparator()
        act("zoomout", "縮小", lambda: e.set_zoom(e.zoom / 1.2), QKeySequence.StandardKey.ZoomOut)
        tb.addWidget(e.lbl_zoom)
        act("zoomin", "放大", lambda: e.set_zoom(e.zoom * 1.2), QKeySequence.StandardKey.ZoomIn)
        act("fit", "符合寬度", e.fit_width)
        act("fitpage", "整頁顯示", e.fit_page)
        tb.addSeparator()
        act("rotl", "目前頁面向左旋轉 90°", lambda: e.rotate_page(-90),
            QKeySequence(Qt.Modifier.CTRL | Qt.Modifier.SHIFT | Qt.Key.Key_Minus))
        act("rotr", "目前頁面向右旋轉 90°", lambda: e.rotate_page(90),
            QKeySequence(Qt.Modifier.CTRL | Qt.Modifier.SHIFT | Qt.Key.Key_Plus))
        return tb

    # ---- 左側工具軌 ----
    def build_rail(self):
        rail = QFrame()
        rail.setObjectName("Rail")
        lay = QVBoxLayout(rail)
        lay.setContentsMargins(0, 8, 0, 8)
        lay.setSpacing(2)
        self.rail_group = QButtonGroup(self)
        self.rail_group.setExclusive(True)
        self.rail_buttons = []
        for i, (icon, name, _) in enumerate(self.RAIL):
            b = QToolButton()
            b.setIcon(make_icon(icon, size=24, tool=True))
            b.setIconSize(QSize(24, 24))
            b.setText(name)
            b.setToolTip(self.RAIL[i][2])
            b.setCheckable(True)
            b.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.clicked.connect(lambda _=False, k=i: self.set_view(k))
            self.rail_group.addButton(b, i)
            lay.addWidget(b)
            self.rail_buttons.append(b)
        lay.addStretch()
        return rail

    def open_pdf(self, path):
        """開啟指定的 PDF 到編輯頁（命令列 / 檔案關聯）"""
        self.set_view(0)
        self.edit_tab.load(path)

    def set_view(self, i):
        """切換左側工具（0=編輯 1=合併 2=組織 3=轉換 4=浮水印）"""
        self.stack.setCurrentIndex(i)
        self.rail_buttons[i].setChecked(True)
        self.update_state()

    def update_state(self):
        e = self.edit_tab
        if e.doc:
            name = os.path.basename(e.path or "") + (" *" if e.dirty else "")
            self.lbl_doc.setText(name)
            self.setWindowTitle(f"{name} - {APP_TITLE}")
            self.lbl_state.setText(f"第 {e.page_no + 1} / {len(e.doc)} 頁　·　{int(e.zoom * 100)}%")
        else:
            self.lbl_doc.setText("")
            self.setWindowTitle(APP_TITLE)
            self.lbl_state.setText(self.RAIL[self.stack.currentIndex()][1])

    def closeEvent(self, ev):
        # 關閉前確認編輯檢視是否有未儲存的修改
        if self.edit_tab.confirm_discard():
            ev.accept()
        else:
            ev.ignore()


def startup_pdf(argv):
    """從命令列參數找出要開啟的 PDF（檔案總管「開啟檔案」會傳入 "%1"）；略過 --選項及其參數"""
    skip = False
    for a in argv[1:]:
        if skip:
            skip = False
            continue
        if a.startswith("--"):
            skip = a in ("--selftest-ocr", "--check-open")
            continue
        if a.lower().endswith(".pdf") and os.path.isfile(a):
            return os.path.abspath(a)
    return None


def main():
    # 命令列自我測試（不開視窗）：CGU_PDF_Tool.exe --selftest-ocr 結果.txt
    if "--selftest-ocr" in sys.argv:
        i = sys.argv.index("--selftest-ocr")
        out = sys.argv[i + 1] if len(sys.argv) > i + 1 else None
        sys.exit(0 if ocr_selftest(out) else 1)
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setApplicationName(APP_NAME)
    app.setStyleSheet(STYLE)
    app.setWindowIcon(app_window_icon())
    w = MainWindow()
    # 驗證用：CGU_PDF_Tool.exe --check-open 結果.txt 檔案.pdf → 不顯示視窗，開檔後寫出結果並結束
    if "--check-open" in sys.argv:
        i = sys.argv.index("--check-open")
        out = sys.argv[i + 1] if len(sys.argv) > i + 1 else "check_open.txt"
        pdf = startup_pdf(sys.argv)
        if pdf:
            w.open_pdf(pdf)
        ok_ = bool(pdf and w.edit_tab.doc is not None)
        with open(out, "w", encoding="utf-8") as f:
            f.write(f"{'OPEN PASS' if ok_ else 'OPEN FAIL'} {pdf} pages={len(w.edit_tab.doc) if ok_ else 0}\n")
        sys.exit(0 if ok_ else 1)
    w.show()
    StartupNotice(w).exec()           # 每次開啟顯示開發單位提示
    pdf = startup_pdf(sys.argv)
    if pdf:
        w.open_pdf(pdf)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
