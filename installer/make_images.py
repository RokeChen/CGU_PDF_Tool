"""產生 Inno Setup 精靈圖片（左側大圖 / 右上小圖，含多種 DPI 尺寸）"""
import os
from PIL import Image, ImageDraw, ImageFont
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(HERE, "img")
os.makedirs(OUT, exist_ok=True)
LOGO = Image.open(os.path.join(ROOT, "assets", "app_logo.png")).convert("RGBA")
# 字型：Linux 用 Noto Sans CJK（TC），Windows 用微軟正黑體
if os.path.exists("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"):
    FB = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
    FR = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
    TC = 3  # TC face index in the ttc（JP=0,KR=1,SC=2,TC=3）
else:
    FB = r"C:\Windows\Fonts\msjhbd.ttc"
    FR = r"C:\Windows\Fonts\msjh.ttc"
    TC = 0
VERSION = "V2.5"
DARK, DARK2, ORANGE = (31, 31, 31), (14, 14, 14), (238, 130, 0)

def font(path, size):
    return ImageFont.truetype(path, size, index=TC)

def logo_white(h):
    lg = LOGO.resize((round(LOGO.width * h / LOGO.height), h), Image.LANCZOS)
    return lg

def gradient(w, h):
    im = Image.new("RGB", (w, h), DARK)
    px = im.load()
    for y in range(h):
        t = y / (h - 1)
        for x in range(w):
            u = min(1.0, (x / w) * 0.3 + t * 0.7)
            px[x, y] = tuple(round(DARK[i] * (1 - u) + DARK2[i] * u) for i in range(3))
    # 底部橘色光暈
    glow = Image.new("RGBA", (w, h), (0, 0, 0, 0)); g = ImageDraw.Draw(glow)
    for i in range(40):
        a = int(70 * (1 - i / 40) ** 2)
        r = int(w * (0.9 + i * 0.03))
        g.ellipse((w * 0.5 - r, h * 1.02 - r * 0.55, w * 0.5 + r, h * 1.02 + r * 0.55), outline=None, fill=(ORANGE[0], ORANGE[1], ORANGE[2], max(0, a // 6)))
    im = Image.alpha_composite(im.convert("RGBA"), glow).convert("RGB")
    return im

def big(w, h):
    s = w / 164
    im = gradient(w, h).convert("RGBA"); d = ImageDraw.Draw(im)
    d.rectangle((w - max(2, round(3 * s)), 0, w, h), fill=ORANGE)          # 右側橘線
    lg = logo_white(round(46 * s)); im.alpha_composite(lg, ((w - lg.width) // 2, round(54 * s)))
    y = round(118 * s)
    for line, sz in (("長庚大學", 17), ("PDF簡易編輯工具", 15)):
        f = font(FB, round(sz * s)); tw = d.textlength(line, font=f)
        d.text(((w - tw) / 2, y), line, font=f, fill=(255, 255, 255)); y += round((sz + 9) * s)
    # 版本徽章
    f = font(FB, round(10 * s)); t = VERSION; tw = d.textlength(t, font=f)
    bw, bh = tw + 16 * s, 17 * s; bx, by = (w - bw) / 2, y + 6 * s
    d.rounded_rectangle((bx, by, bx + bw, by + bh), radius=bh / 2, fill=ORANGE)
    d.text((bx + 8 * s, by + 1.2 * s), t, font=f, fill=DARK)
    d.line((w * 0.3, by + bh + 22 * s, w * 0.7, by + bh + 22 * s), fill=(80, 80, 80), width=max(1, round(s)))
    f = font(FR, round(8.5 * s))
    for i, t in enumerate(("PDF 編輯・合併・轉換・OCR", "開源軟體 AGPL-3.0 · CGU_RC")):
        tw = d.textlength(t, font=f)
        d.text(((w - tw) / 2, h - (34 - i * 14) * s), t, font=f, fill=(168, 168, 168) if i == 0 else (130, 130, 130))
    return im.convert("RGB")

def small(w, h):
    s = w / 55
    im = Image.new("RGBA", (w, h), (255, 255, 255, 255)); d = ImageDraw.Draw(im)
    r = round(8 * s)
    d.rounded_rectangle((round(2 * s), round(2 * s), w - round(2 * s), h - round(2 * s)), radius=r, fill=DARK)
    lg = logo_white(round(28 * s)); im.alpha_composite(lg, ((w - lg.width) // 2, (h - lg.height) // 2 - round(3 * s)))
    d.rounded_rectangle((round(14 * s), h - round(11 * s), w - round(14 * s), h - round(8 * s)), radius=round(1.5 * s), fill=ORANGE)
    return im.convert("RGB")

BIG = [(164, 314), (192, 386), (246, 459), (273, 556), (328, 604), (383, 707), (410, 797)]
SMALL = [(55, 55), (64, 68), (83, 80), (92, 97), (110, 106), (119, 123), (138, 140)]
for w, h in BIG: big(w, h).save(os.path.join(OUT, f"wizard_{w}x{h}.bmp"))
for w, h in SMALL: small(w, h).save(os.path.join(OUT, f"small_{w}x{h}.bmp"))
print("ok")
