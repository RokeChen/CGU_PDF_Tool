"""以使用者提供的掃描 PDF 做實測：OCR → 修改一個辨識出的詞 → 輸出前後對照圖"""
import os, sys
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pdf_tool as T
from pdf_tool import fitz
src = os.path.join(os.path.dirname(__file__), "user_sample.pdf")
doc = fitz.open(src)
print("pages", len(doc), "native text chars:", len(doc[0].get_text().strip()))
doc[0].get_pixmap(dpi=110).save("screens/v3_user_before.png")
data, n = T.ocr_document_bytes(doc.tobytes(), list(range(len(doc))))
d2 = fitz.open("pdf", data)
pg = d2[0]
print("OCR lines:", n)
target = sys.argv[1] if len(sys.argv) > 1 else "保養日"
hits = pg.search_for(target)
print("search", target, "->", len(hits))
sp = T.find_span_at(pg, hits[0].tl + (2, 2))
print("span:", sp["text"], "ocr:", T.is_invisible_span(pg, sp))
new = sys.argv[2] if len(sys.argv) > 2 else "保養日期"
info = T.replace_span_text(pg, sp, sp["text"].replace(target, new))
print("font used:", info["label"].replace("\n", " "), f"{info['size']:.1f}pt")
pg.get_pixmap(dpi=110).save("screens/v3_user_after.png")
# 局部放大對照
clip = fitz.Rect(hits[0]) + (-60, -25, 120, 25)
fitz.open(src)[0].get_pixmap(dpi=220, clip=clip).save("screens/v3_user_before_zoom.png")
pg.get_pixmap(dpi=220, clip=clip).save("screens/v3_user_after_zoom.png")
d2.save("/tmp/user_sample_ocr_edited.pdf")
print("search after edit:", len(fitz.open("/tmp/user_sample_ocr_edited.pdf")[0].search_for(new)))
