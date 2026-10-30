r"""
一次修改所有版本號，並重新產生安裝畫面圖片。

用法（在專案資料夾執行）：
    .\.venv\Scripts\python.exe set_version.py 2.4

會修改：
  pdf_tool.py              APP_VERSION = "V2.4"
  installer.iss            #define MyAppVersion "2.4"
  version_info.txt         filevers / prodvers (2, 4, 0, 0)、FileVersion / ProductVersion '2.4.0.0'
  installer/make_images.py VERSION = "V2.4"（並重新產生 installer/img/*.bmp）
"""
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def edit(rel, pattern, repl, count_expected=None):
    path = os.path.join(HERE, rel)
    raw = open(path, "rb").read()
    bom = raw.startswith(b"\xef\xbb\xbf")
    text = raw.decode("utf-8-sig")
    new, n = re.subn(pattern, repl, text, flags=re.M)
    if n == 0 or (count_expected and n != count_expected):
        sys.exit(f"[錯誤] {rel}：找不到要修改的版本號（符合 {n} 處）")
    with open(path, "wb") as f:
        f.write((b"\xef\xbb\xbf" if bom else b"") + new.encode("utf-8"))
    print(f"  ✓ {rel}（{n} 處）")


def main():
    if len(sys.argv) != 2 or not re.fullmatch(r"\d+\.\d+(\.\d+)?", sys.argv[1]):
        sys.exit("用法：python set_version.py 2.4   （或 2.4.1）")
    ver = sys.argv[1]
    parts = [int(x) for x in ver.split(".")] + [0, 0]
    a, b, c = parts[:3]
    four = f"{a}.{b}.{c}.0"
    tup = f"({a}, {b}, {c}, 0)"
    print(f"版本號改為 V{ver}：")
    edit("pdf_tool.py", r'^APP_VERSION = "V[^"]*"', f'APP_VERSION = "V{ver}"', 1)
    edit("installer.iss", r'#define MyAppVersion "[^"]*"', f'#define MyAppVersion "{ver}"', 1)
    edit("version_info.txt", r"filevers=\([^)]*\), prodvers=\([^)]*\)", f"filevers={tup}, prodvers={tup}", 1)
    edit("version_info.txt", r"'(FileVersion|ProductVersion)', '[^']*'", rf"'\1', '{four}'", 2)
    edit(os.path.join("installer", "make_images.py"), r'^VERSION = "V[^"]*"', f'VERSION = "V{ver}"', 1)
    print("重新產生安裝畫面圖片…")
    r = subprocess.run([sys.executable, os.path.join(HERE, "installer", "make_images.py")])
    if r.returncode != 0:
        sys.exit("[錯誤] 安裝畫面圖片產生失敗（版本號已修改）")
    print(f"\n完成！發布時請用：git tag v{ver}")


if __name__ == "__main__":
    main()
