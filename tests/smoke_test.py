import os, sys, tempfile
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pdf_tool as T
from pdf_tool import fitz
from pypdf import PdfReader
from PIL import Image
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QPointF, QEventLoop, QRect

D = tempfile.mkdtemp()
def mk(name, n, label):
    d = fitz.open()
    for i in range(n):
        p = d.new_page(); p.insert_text((72, 100), f"{label} page {i+1}", fontsize=20)
    path = os.path.join(D, name); d.save(path); return path
def text(path, i): return PdfReader(path).pages[i].extract_text()
ok = lambda c, m: (print("PASS" if c else "FAIL", m), c or sys.exit(1))

# 1 convert
Image.new("RGBA", (200, 100), (255, 0, 0, 128)).save(f"{D}/a.png")
Image.new("RGB", (100, 200), "blue").save(f"{D}/b.jpg")
open(f"{D}/c.docx", "wb").write(b"PK fake")
outs, errs = T.run_conversion([f"{D}/a.png", f"{D}/b.jpg", f"{D}/c.docx"], D, False)
ok(len(outs) == 2 and len(errs) == 1 and "Word" in errs[0], f"convert separate; docx error graceful: {errs}")
outs, errs = T.run_conversion([f"{D}/a.png", f"{D}/b.jpg"], D, True)
ok(len(PdfReader(outs[0]).pages) == 2, "convert combined images -> 2 pages (RGBA handled)")

# 2 merge
a, b = mk("A.pdf", 2, "AAA"), mk("B.pdf", 3, "BBB")
m = T.merge_pdfs([b, a], f"{D}/m.pdf")
ok(len(PdfReader(m).pages) == 5 and "BBB page 1" in text(m, 0) and "AAA page 2" in text(m, 4), "merge order/page count")

# 3 organize
src = mk("O.pdf", 4, "PG")
o = T.export_organized(src, [(3, 90), (0, 0), (2, 270)], f"{D}/o.pdf")
r = PdfReader(o)
ok(len(r.pages) == 3 and "PG page 4" in text(o, 0) and r.pages[0].rotation == 90 and r.pages[2].rotation == 270, "organize order/delete/rotation")

# 4 watermark
w = T.add_text_watermark(a, f"{D}/w.pdf", "機密 CONFIDENTIAL", opacity=0.3, angle=30)
ok(all("CONFIDENTIAL" in text(w, i) and "機密" in text(w, i) for i in range(2)), "watermark text (incl. Chinese) on all pages")

# 5 edit core
doc = fitz.open(mk("E.pdf", 1, "Hello"))
pg = doc[0]
span = T.find_span_at(pg, fitz.Point(80, 95))
ok(span and span["text"] == "Hello page 1", "find span")
T.replace_span_text(pg, span, "Goodbye 世界")
T.add_text_box(pg, fitz.Point(72, 200), "新增中文文字", 14, (0, 0, 1))
T.highlight_area(pg, fitz.Rect(60, 70, 300, 110))
T.whiteout_area(pg, fitz.Rect(0, 700, 600, 800))
T.insert_image_at(pg, fitz.Point(300, 300), f"{D}/b.jpg", 100)
pg.add_text_annot((400, 400), "note")
ok(T.delete_annot_at(pg, fitz.Point(405, 405)), "delete annotation")
doc.save(f"{D}/e.pdf"); t = text(f"{D}/e.pdf", 0)
ok("Goodbye" in t and "世界" in t and "Hello" not in t and "新增中文文字" in t, f"edited text replaced: {t!r}")

# GUI: build window, drive edit tab with undo, worker thread, screenshots
app = QApplication(sys.argv)
app.setStyle("Fusion"); app.setStyleSheet(T.STYLE)
win = T.MainWindow(); win.resize(1280, 820); win.show()
ok(win.windowTitle().startswith(T.APP_NAME) and not T.logo_pixmap().isNull(), "title + embedded CGU logo")
et = win.edit_tab; et.load(mk("G.pdf", 3, "GUI"))
et.set_tool("rect"); et.on_drag(QPointF(50, 50), QPointF(200, 150))
et.set_tool("pen"); et.on_stroke([QPointF(100, 300), QPointF(150, 320), QPointF(200, 300)])
ok(len(list(et.doc[0].annots())) == 2 and len(et.undo_stack) == 2, "GUI edit rect+pen annots")
et.undo(); ok(len(list(et.doc[0].annots())) == 1, "undo")
et.redo(); ok(len(list(et.doc[0].annots())) == 2, "redo")
et.undo()
ok(et.thumbs.count() == 3, "thumbnail panel")
et.opacity.setValue(50); et.set_tool("rect"); et.on_drag(QPointF(300, 300), QPointF(400, 400))
ok(abs(list(et.doc[0].annots())[-1].opacity - 0.5) < 0.01, "opacity property applied"); et.undo()
win.act_thumbs.setChecked(False); ok(not et.thumb_panel.isVisible(), "collapse thumbs"); win.act_thumbs.setChecked(True)
et.goto(1); et.set_zoom(1.5); et.write_to(f"{D}/g.pdf"); ok(len(PdfReader(f"{D}/g.pdf").pages) == 3, "save")
wk = T.Worker(T.merge_pdfs, [a, b], f"{D}/wm.pdf"); res = []; prog = []
wk.result.connect(res.append); wk.progress.connect(prog.append)
loop = QEventLoop(); wk.finished.connect(loop.quit); wk.start(); loop.exec()
ok(res and prog and prog[-1] == 100, "QThread worker signals")
win.organize_tab.load(src); g = win.organize_tab.grid
g.setCurrentRow(0); win.organize_tab.rotate_selected(90); win.organize_tab.move_selected(1)
ok(win.organize_tab.page_specs()[:2] == [(1, 0), (0, 90)], "organize UI rotate+move")
win.convert_tab.add_files([f"{D}/a.png", f"{D}/b.jpg"]); win.merge_tab.add_files([a, b]); win.watermark_tab.src.setText(a)
et.goto(0)
os.makedirs("screens", exist_ok=True)
et.set_tool("highlight"); et.fit_width()
for i, n in enumerate(["1_edit", "2_merge", "3_organize", "4_convert", "5_watermark"]):
    win.set_view(i); app.processEvents(); win.grab().save(f"screens/v2_{n}.png")
win.set_view(0); app.processEvents(); win.grab(QRect(0, 0, 1280, 100)).save("screens/v2_header.png")

# ---------- v3：字型比對 / OCR / 置中對話框 ----------
from PyQt6.QtCore import QSize
cjk = fitz.Font("cjk")
fd = fitz.open(); fp = fd.new_page()
fp.insert_font(fontname="EMB", fontbuffer=cjk.buffer)
fp.insert_text((72, 100), "Hello 長庚 page", fontsize=18, fontname="EMB", color=(0, 0, 1))
fsp = T.find_span_at(fp, fitz.Point(80, 95))
fi = T.choose_font(fp, fsp, "Hello 大學 world")
ok(fi["kind"] == "embedded", f"embedded font reused: {fi['label']}")
info = T.replace_span_text(fp, fsp, "Hello 大學 world", shrink=False)
sp2 = T.find_span_at(fp, fitz.Point(80, 95))
ok(info and info.get("kind") == "embedded" and "大學" in sp2["text"] and abs(sp2["size"] - 18) < 0.6
   and sp2["color"] == 0x0000ff and abs(sp2["origin"][1] - 100) < 0.6, f"replace keeps size/color/baseline: {sp2['text']!r} {sp2['size']:.1f}")
# 子集字型缺字 → 改用系統 / 內建字型
fd.subset_fonts()
fd2 = fitz.open("pdf", fd.tobytes()); fp2 = fd2[0]
fsp2 = T.find_span_at(fp2, fitz.Point(80, 95))
fi2 = T.choose_font(fp2, fsp2, "測試缺字龘")
ok(fi2["kind"] in ("system", "builtin") and T.font_covers(fi2["font"], "測試缺字龘"), f"subset missing glyphs -> fallback: {fi2['label']}")
# 名稱對應（Times → serif）
tdoc = fitz.open(); tp = tdoc.new_page(); tp.insert_text((72, 100), "Times text", fontname="tiro", fontsize=12)
tfi = T.choose_font(tp, T.find_span_at(tp, fitz.Point(80, 96)), "New words")
ok("serif" in tfi["label"].lower() or "times" in tfi["label"].lower() or "tiro" in tfi["label"], f"Times maps to serif: {tfi['label']}")
# shrink-to-fit
sdoc = fitz.open(); spg = sdoc.new_page(); spg.insert_text((72, 100), "short", fontsize=20, fontname="helv")
ssp = T.find_span_at(spg, fitz.Point(80, 95)); w0 = fitz.Rect(ssp["bbox"]).width
si = T.replace_span_text(spg, ssp, "a much much longer replacement", shrink=True)
ok(si["size"] < 20, f"shrink-to-fit size {si['size']:.1f}")
# OCR：產生只有影像的 PDF
scan = T.make_scanned_test_pdf(f"{D}/scan.pdf")
ok(not fitz.open(scan)[0].get_text().strip(), "scan PDF has no text before OCR")
data, n = T.ocr_document_bytes(open(scan, "rb").read(), [0])
od = fitz.open("pdf", data); ot = od[0].get_text()
ok(n >= 3 and "Hello" in ot and "長庚大學" in ot and "12345" in ot, f"OCR text layer: {ot!r}")
ok(len(od[0].search_for("Hello")) == 1, "OCR text searchable")
osp = T.find_span_at(od[0], od[0].search_for("Hello")[0].tl + (3, 3))
ok(osp and T.is_invisible_span(od[0], osp), "OCR span is invisible layer")
T.replace_span_text(od[0], osp, "Hi CGU 67890")
ot2 = od[0].get_text()
ok("Hi CGU 67890" in ot2 and "Hello" not in ot2, "edit OCR text (cover + rewrite)")
# GUI：OCR 透過背景執行緒、對話框置中
from PyQt6.QtCore import QRect
r = T.centered_rect(QRect(100, 50, 1200, 800), QSize(400, 200))
ok(r == QRect(500, 350, 400, 200), "centered_rect math")
et.load(scan); et.rb_ocr_cur.setChecked(True); et.run_ocr()
loop = QEventLoop(); et.worker.finished.connect(loop.quit); loop.exec(); app.processEvents()
ok("Hello" in et.doc[0].get_text() and len(et.undo_stack) == 1, "GUI OCR via QThread (+undo)")
win.set_view(0); app.processEvents()
win.grab().save("screens/v3_ocr.png")
hs = et.doc[0].search_for("Hello")[0]
dlg = T.EditSpanDialog(et.doc[0], T.find_span_at(et.doc[0], hs.tl + (3, 3)), win)
dlg.show(); app.processEvents()
c1, c2 = dlg.frameGeometry().center(), win.frameGeometry().center()
ok(abs(c1.x() - c2.x()) <= 30 and abs(c1.y() - c2.y()) <= 30 and dlg.isModal(), f"dialog centered on main window {c1} vs {c2}")
ok("Hello" in dlg.edit.text() and dlg.lbl_font.text(), "dialog prefilled + font label")
dlg.grab().save("screens/v3_edit_dialog.png")
# 把對話框畫進主視窗截圖
from PyQt6.QtGui import QPainter
full = win.grab(); pnt = QPainter(full); pnt.drawPixmap(dlg.geometry().topLeft() - win.geometry().topLeft(), dlg.grab()); pnt.end()
full.save("screens/v3_edit_dialog_in_window.png")
dlg.reject()
# 原生文字對話框
et.load(f"{D}/g.pdf"); app.processEvents()
d2 = T.EditSpanDialog(et.doc[0], T.find_span_at(et.doc[0], fitz.Point(80, 95)), win); d2.show(); app.processEvents()
d2.grab().save("screens/v3_edit_dialog_native.png"); d2.reject()
print("ALL PASSED")
