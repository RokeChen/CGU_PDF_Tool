"""PC 實測：Microsoft Word COM 路徑（文字 PDF + 掃描影像 PDF）→ 以 python-docx 檢查內容"""
import os, sys, tempfile
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pdf_tool as T
import docx
fitz = T.fitz
d = tempfile.mkdtemp(prefix="cgu_w_")
tp = os.path.join(d, "text_pdf.pdf")
doc = fitz.open(); pg = doc.new_page(); pg.insert_text((72, 100), "Hello Word COM 12345", fontsize=16)
ff = T.ocr_rebuild_font(); print("font:", ff)
pg.insert_font(fontname="C", fontfile=ff); pg.insert_text((72, 140), "長庚大學轉換測試", fontsize=16, fontname="C"); doc.save(tp)
# 內建 Droid 字型的 PDF：Word 會漏字 → auto 模式應自動改用 pdf2docx
dp = os.path.join(d, "droid_pdf.pdf")
doc = fitz.open(); pg = doc.new_page(); pg.insert_font(fontname="C", fontbuffer=fitz.Font("cjk").buffer); pg.insert_text((72, 100), "長庚大學轉換測試", fontsize=16, fontname="C"); doc.save(dp)
sp = T.make_scanned_test_pdf(os.path.join(d, "scan_pdf.pdf"))
print("word_available:", T.word_available())
res = T.run_pdf_to_word([tp, sp], None, True, "word")
def text(p):
    return T.docx_text(p)
ok = True
for o, eng, ocr in res["outputs"]:
    t = text(o)
    print(os.path.basename(o), eng, "OCR" if ocr else "", os.path.getsize(o), repr(t[:80]))
need = {"text_pdf.docx": ["Hello Word COM", "長庚大學"], "scan_pdf.docx": ["Hello", "長庚"]}
got = {os.path.basename(o): text(o) for o, *_ in res["outputs"]}
for k, words in need.items():
    good = k in got and all(w in got[k] for w in words)
    print(k, "PASS" if good else "FAIL"); ok &= good
ra = T.run_pdf_to_word([dp], None, True, "auto")
o = ra["outputs"][0]; good = "長庚大學轉換測試" in text(o[0]); print("droid auto:", o[1], "PASS" if good else "FAIL"); ok &= good
out = open(os.path.join(tempfile.gettempdir(), "cgu_wordtest.txt"), "w", encoding="utf-8")
for o_, eng, ocr in res["outputs"] + ra["outputs"]:
    out.write(f"{os.path.basename(o_)} | {eng} | OCR={ocr} | {text(o_)}\n")
out.write("WORDTEST " + ("PASS" if ok and not res["errors"] else "FAIL") + "\n"); out.close()
print("errors:", res["errors"])
print("WORDTEST", "PASS" if ok and not res["errors"] else "FAIL")
