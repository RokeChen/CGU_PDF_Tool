# SignPath 設定參考

在 SignPath 專案（project slug：`cgu-pdf-tool`）中建立兩個 Artifact configuration，
內容分別貼上下面兩段 XML。GitHub 上傳的產物是 ZIP，所以最外層是 `<zip-file>`。

**exe**（slug：`exe`）
```xml
<?xml version="1.0" encoding="utf-8"?>
<artifact-configuration xmlns="http://signpath.io/artifact-configuration/v1">
  <zip-file>
    <pe-file path="CGU_PDF_Tool.exe">
      <authenticode-sign />
    </pe-file>
  </zip-file>
</artifact-configuration>
```

**installer**（slug：`installer`）
```xml
<?xml version="1.0" encoding="utf-8"?>
<artifact-configuration xmlns="http://signpath.io/artifact-configuration/v1">
  <zip-file>
    <pe-file path="*.exe">
      <authenticode-sign />
    </pe-file>
  </zip-file>
</artifact-configuration>
```

Signing policy slug：`release-signing`（需人工核准）。
GitHub repo 的 Settings → Secrets and variables → Actions：
- Secret `SIGNPATH_API_TOKEN`
- Variable `SIGNPATH_ORGANIZATION_ID`（slug 若不同，也可設定 `SIGNPATH_PROJECT_SLUG`、`SIGNPATH_POLICY_SLUG`）
