# CGU PDF Tool（長庚大學PDF簡易編輯工具）

免費、離線、開源的 Windows PDF 編輯工具，給一般教職員日常使用：改字、合併、轉檔、簽名、蓋章、OCR，一個程式就搞定。

**Free code signing provided by [SignPath.io](https://about.signpath.io), certificate by [SignPath Foundation](https://signpath.org).**

## 功能

| 分類 | 功能 |
|---|---|
| 編輯 PDF | 修改原有文字、新增文字／圖片、螢光標記、方框、直線、手繪、便利貼、塗銷、刪除註解 |
| 超連結 | 連到本檔頁面或網址，可加淺藍底色提示 |
| 打勾 | 點框框打勾／再點取消，可框選一次勾整欄（✓ ✗ ●） |
| 簽名印章 | 匯入印章圖（自動去白底）或手寫簽名，可移動、縮放 |
| 數位簽章 | 以 .pfx 憑證簽署 PDF，可多人接力簽 |
| OCR | 繁中＋英文離線辨識 |
| 列印 | 可選頁碼範圍，橫式頁自動轉向 |
| 合併檔案 | 多個 PDF 依順序合併 |
| 組織頁面 | 合併多個檔案、插入檔案／空白頁、刪除、旋轉、拖曳排序（類似 Acrobat） |
| 格式轉換 | Word／圖片轉 PDF（Microsoft Word 或 LibreOffice）、PDF 轉 Word |
| 文字浮水印 | 支援中文，可調整樣式 |

## 下載

到 [Releases](../../releases) 下載最新的 `CGU_PDF_Tool_Setup_x.x.exe`（安裝版）。安裝不需要系統管理員權限。

Word 轉 PDF 需要電腦上有 Microsoft Word；沒有 Word 時會自動改用免費的 [LibreOffice](https://zh-tw.libreoffice.org/)。

## 從原始碼執行與建置

需要 Windows、Python 3.12（OCR 套件目前最高支援 3.12）。

```
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt pyinstaller
.\.venv\Scripts\python.exe pdf_tool.py
```

建置執行檔與安裝程式（需要 [Inno Setup 6](https://jrsoftware.org/isinfo.php)）：

```
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean CGU_PDF_Tool.spec
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```

正式發行版由 GitHub Actions（`.github/workflows/release.yml`）從本儲存庫的原始碼自動建置並簽章。

## 授權

本程式以 [GNU AGPL-3.0](LICENSE) 授權發布，可自由使用、修改與散布；散布修改版時須一併提供原始碼。
使用的第三方元件與授權列在程式的「關於」視窗中（PyQt6 — GPL-3.0、PyMuPDF — AGPL-3.0、RapidOCR／ONNX Runtime — Apache-2.0、pyHanko — MIT、Lucide 圖示 — ISC 等）。

## Code signing policy

Free code signing provided by [SignPath.io](https://about.signpath.io), certificate by [SignPath Foundation](https://signpath.org).

Windows release binaries (`CGU_PDF_Tool.exe` and the `CGU_PDF_Tool_Setup_*.exe` installer) are built from this repository by GitHub Actions and signed through SignPath. Only binaries built from this repository's source code are signed.

**Team roles**

- Committers and reviewers: [@RokeChen](https://github.com/RokeChen)
- Approvers: [@RokeChen](https://github.com/RokeChen)

Every release is manually approved before signing. All team members use multi-factor authentication for GitHub and SignPath.

## Privacy policy

This program will not transfer any information to other networked systems unless specifically requested by the user or the person installing or operating it.

All processing (PDF editing, OCR, conversion, signing) happens locally on the user's computer. The program only opens a web page when the user clicks a link, and Word conversion uses Microsoft Word or LibreOffice installed on the same computer.

（隱私說明：本程式所有處理都在使用者電腦上離線完成，不會將任何資料傳送到網路；只有在使用者自行點選連結時才會開啟網頁。）
