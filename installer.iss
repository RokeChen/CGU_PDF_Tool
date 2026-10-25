; 長庚大學PDF簡易編輯工具 安裝程式（Inno Setup 6.4+）
; 編譯：build_installer.bat → dist\CGU_PDF_Tool_Setup_<版本>.exe
; 精靈圖片由 installer\make_images.py 產生（installer\img\*.bmp）
#define MyAppName "長庚大學PDF簡易編輯工具"
#define MyAppVersion "2.1"
#define MyAppPublisher "CGU_RC"
#define MyAppURL "https://www.cgu.edu.tw"
#define MyAppExeName "CGU_PDF_Tool.exe"
#define MyProgId "CGU.PDFTool.pdf"
#define MyCapKey "Software\CGU_RC\CGU_PDF_Tool\Capabilities"

[Setup]
AppId={{6D3B1C2E-8F4A-4C1B-9E2D-0C6A1F2B8E44}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} V{#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppCopyright=Copyright © 2026 {#MyAppPublisher}（AGPL-3.0）
VersionInfoVersion={#MyAppVersion}.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
VersionInfoDescription={#MyAppName} 安裝程式
VersionInfoCopyright=Copyright © 2026 {#MyAppPublisher}
; 單一使用者安裝，不需系統管理員權限，也不再詢問安裝模式
PrivilegesRequired=lowest
DefaultDirName={autopf}\CGU_PDF_Tool
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableDirPage=auto
DisableWelcomePage=no
DisableReadyPage=no
ShowLanguageDialog=no
ChangesAssociations=yes
OutputDir=dist
OutputBaseFilename=CGU_PDF_Tool_Setup_{#MyAppVersion}
SetupIconFile=assets\app_icon.ico
LicenseFile=LICENSE
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
WizardStyle=modern
WizardImageFile=installer\img\wizard_164x314.bmp,installer\img\wizard_192x386.bmp,installer\img\wizard_246x459.bmp,installer\img\wizard_273x556.bmp,installer\img\wizard_328x604.bmp,installer\img\wizard_383x707.bmp,installer\img\wizard_410x797.bmp
WizardSmallImageFile=installer\img\small_55x55.bmp,installer\img\small_64x68.bmp,installer\img\small_83x80.bmp,installer\img\small_92x97.bmp,installer\img\small_110x106.bmp,installer\img\small_119x123.bmp,installer\img\small_138x140.bmp
WizardImageStretch=no
WizardImageBackColor=$1F1F1F
Compression=lzma2/max
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes

[Languages]
Name: "cht"; MessagesFile: "installer\ChineseTraditional.isl"

[Messages]
SetupAppTitle=安裝程式
SetupWindowTitle=安裝 %1
UninstallAppTitle=解除安裝
WelcomeLabel1=歡迎使用 [name] 安裝程式
WelcomeLabel2=這個精靈會在您的電腦上安裝 [name/ver]。%n%n本工具提供 PDF 文字修改、合併、頁面組織、格式轉換、文字浮水印與離線 OCR 辨識（繁中＋英文）。%n%n安裝只會套用到目前的 Windows 帳號，不需要系統管理員權限。建議先關閉其他應用程式再繼續。
ClickNext=按「下一步」繼續，或按「取消」離開安裝程式。
SelectTasksLabel2=請選擇安裝 [name] 時要一併執行的項目，然後按「下一步」。
ReadyLabel1=已準備好安裝 [name]。
ReadyLabel2a=按「安裝」開始安裝；若要確認或修改設定，請按「上一步」。
InstallingLabel=正在安裝 [name]，請稍候…
FinishedHeadingLabel=[name] 安裝完成
FinishedLabelNoIcons=[name] 已安裝到您的電腦。
FinishedLabel=[name] 已安裝到您的電腦，可從「開始」功能表或捷徑開啟。%n%n之後在 PDF 檔案上按右鍵 →「開啟檔案」，即可選擇本程式；若要設為預設 PDF 程式，可到 Windows「預設應用程式」設定。
ClickFinish=按「完成」結束安裝程式。
ConfirmUninstall=確定要完全移除 %1 及其所有元件嗎？
UninstalledAll=%1 已從您的電腦移除。
; ---- 正式翻譯檔缺少的訊息（Inno Setup 6.4+ 新增） ----
DownloadingLabel2=正在下載檔案…
ExtractingLabel=正在解壓縮檔案…
ButtonStopExtraction=停止解壓縮(&S)
StopExtraction=確定要停止解壓縮嗎？
ErrorExtractionAborted=已中止解壓縮
ErrorExtractionFailed=解壓縮失敗：%1
ArchiveIncorrectPassword=密碼不正確
ArchiveIsCorrupted=壓縮檔已損毀
ArchiveUnsupportedFormat=不支援的壓縮檔格式
RetryCancelSelectAction=請選擇動作
RetryCancelRetry=重試(&T)
RetryCancelCancel=取消
StatusDownloadFiles=正在下載檔案…
SourceVerificationFailed=來源檔案驗證失敗：%1
VerificationSignatureDoesntExist=找不到簽章檔「%1」
VerificationSignatureInvalid=簽章檔「%1」無效
VerificationKeyNotFound=簽章檔「%1」使用了未知的金鑰
VerificationFileNameIncorrect=檔案名稱不正確
VerificationFileTagIncorrect=檔案標記不正確
VerificationFileSizeIncorrect=檔案大小不正確
VerificationFileHashIncorrect=檔案雜湊值不正確
ErrorDownloading=下載檔案時發生錯誤：
ErrorExtracting=解壓縮檔案時發生錯誤：

[CustomMessages]
DesktopIcon=建立桌面捷徑
ShortcutsGroup=捷徑：
OtherGroup=其他：
OpenDefaultApps=安裝完成後開啟「預設應用程式」設定（可將本程式設為預設 PDF 程式）
LaunchNow=立即啟動 {#MyAppName}
AppDesc=PDF 編輯、合併、頁面組織、格式轉換、簽名與 OCR 辨識工具（開源軟體）
PdfDoc=PDF 文件
UninstallShortcut=解除安裝 {#MyAppName}

[Tasks]
Name: "desktopicon"; Description: "{cm:DesktopIcon}"; GroupDescription: "{cm:ShortcutsGroup}"; Flags: unchecked
Name: "defaultapps"; Description: "{cm:OpenDefaultApps}"; GroupDescription: "{cm:OtherGroup}"; Flags: unchecked

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\app_icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app_icon.ico"; Comment: "{cm:AppDesc}"
Name: "{group}\{cm:UninstallShortcut}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app_icon.ico"; Comment: "{cm:AppDesc}"; Tasks: desktopicon

[Registry]
; ProgID
Root: HKA; Subkey: "Software\Classes\{#MyProgId}"; ValueType: string; ValueName: ""; ValueData: "{cm:PdfDoc}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\{#MyProgId}"; ValueType: string; ValueName: "FriendlyTypeName"; ValueData: "{cm:PdfDoc}"
Root: HKA; Subkey: "Software\Classes\{#MyProgId}\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"
Root: HKA; Subkey: "Software\Classes\{#MyProgId}\shell\open"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "{#MyAppName}"
Root: HKA; Subkey: "Software\Classes\{#MyProgId}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""
; 「開啟檔案」清單
Root: HKA; Subkey: "Software\Classes\.pdf\OpenWithProgids"; ValueType: string; ValueName: "{#MyProgId}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".pdf"; ValueData: ""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""
; 預設應用程式（Default Apps）
Root: HKA; Subkey: "Software\CGU_RC\CGU_PDF_Tool"; Flags: uninsdeletekey
Root: HKA; Subkey: "{#MyCapKey}"; ValueType: string; ValueName: "ApplicationName"; ValueData: "{#MyAppName}"
Root: HKA; Subkey: "{#MyCapKey}"; ValueType: string; ValueName: "ApplicationDescription"; ValueData: "{cm:AppDesc}"
Root: HKA; Subkey: "{#MyCapKey}"; ValueType: string; ValueName: "ApplicationIcon"; ValueData: """{app}\{#MyAppExeName}"",0"
Root: HKA; Subkey: "{#MyCapKey}\FileAssociations"; ValueType: string; ValueName: ".pdf"; ValueData: "{#MyProgId}"
Root: HKA; Subkey: "Software\RegisteredApplications"; ValueType: string; ValueName: "CGU_PDF_Tool"; ValueData: "{#MyCapKey}"; Flags: uninsdeletevalue

[Run]
Filename: "{cmd}"; Parameters: "/c start """" ms-settings:defaultapps"; Flags: runhidden nowait; Tasks: defaultapps
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchNow}"; Flags: nowait postinstall skipifsilent
