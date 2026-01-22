# PasteMD
<p align="center">
  <img src="../../assets/icons/logo.png" alt="PasteMD" width="160" height="160">
</p>

<p align="center">
  <a href="https://github.com/RICHQAQ/PasteMD/releases">
    <img src="https://img.shields.io/github/v/release/RICHQAQ/PasteMD?sort=semver&label=Release&style=flat-square&logo=github" alt="Release">
  </a>
  <a href="https://github.com/RICHQAQ/PasteMD/releases">
    <img src="https://img.shields.io/github/downloads/RICHQAQ/PasteMD/total?label=Downloads&style=flat-square&logo=github" alt="Downloads">
  </a>
  <a href="../../LICENSE">
    <img src="https://img.shields.io/github/license/RICHQAQ/PasteMD?style=flat-square" alt="License">
  </a>
  <img src="https://img.shields.io/badge/Python-3.12%2B-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python 3.12+">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Word%20%7C%20WPS-5e8d36?style=flat-square&logo=windows&logoColor=white" alt="Platform">
</p>

<p align="center">
  <a href="README.en.md">English</a>
  |
  <a href="../../README.md">简体中文</a> 
  |
  <a href="README.ja.md">日本語</a>
</p>

> When writing papers or reports, do formulas copied from AI tools (like ChatGPT or DeepSeek) turn into garbled text in Word? Do Markdown tables fail to paste correctly into Excel? **PasteMD was built specifically to solve these problems.**
> 
> <img src="../../docs/gif/atri/igood.gif"
     alt="I am good"
     width="100">

PasteMD is a lightweight tray app that watches your clipboard, converts Markdown or HTML-rich text to DOCX through Pandoc, and pastes the result straight into the caret position of Word or WPS. It understands Markdown tables and can paste them directly into Excel with formatting preserved, and it recognizes HTML rich text (except math) copied from web pages.

---

## Feature Highlights

### Demo Videos

#### Markdown → Word/WPS

<p align="center">
  <img src="../../docs/gif/demo.gif" alt="Markdown to Word demo" width="600">
</p>

#### Copy AI web reply → Word/WPS
<p align="center">
  <img src="../../docs/gif/demo-html.gif" alt="HTML rich text demo" width="600">
</p>

#### Markdown tables → Excel
<p align="center">
  <img src="../../docs/gif/demo-excel.gif" alt="Markdown table to Excel demo" width="600">
</p>

#### Apply formatting presets
<p align="center">
  <img src="../../docs/gif/demo-chage_format.gif" alt="Formatting demo" width="600">
</p>

### Workflow Boosters

- Global hotkey (default `Ctrl+Shift+B`) to paste the latest Markdown/HTML clipboard snapshot as DOCX.
- Automatically recognizes Markdown tables, converts them to spreadsheets, and pastes into Excel while keeping bold/italic/code formats.
- Recognizes HTML rich text copied from web pages and converts/pastes into Word/WPS.
- Detects the foreground target app (Word, WPS, or Excel) and opens the correct program when needed.
- Tray menu for toggling features, viewing logs, reloading config, and checking for updates.
- Optional toast notifications and background logging for every conversion.

---

## AI Website Compatibility

The following table summarizes how well popular AI chat sites work with PasteMD when copying Markdown or direct HTML content.

| AI Service | Copy Markdown (no formulas) | Copy Markdown (with formulas) | Copy page content (no formulas) | Copy page content (with formulas) |
|------------|----------------------------|-------------------------------|---------------------------------|-----------------------------------|
| Kimi | ✅ Perfect | ✅ Perfect | ✅ Perfect | ⚠️ Formulas missing |
| DeepSeek | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| Tongyi Qianwen | ✅ Perfect | ✅ Perfect | ✅ Perfect | ⚠️ Formulas missing |
| Doubao* | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| ChatGLM/Zhipu | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| ChatGPT | ✅ Perfect | ⚠️ Rendered as code | ✅ Perfect | ✅ Perfect |
| Gemini | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| Grok | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| Claude | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |

_*Doubao requires granting clipboard read permissions in the browser before copying HTML content with formulas (set it via the lock icon near the URL bar)._

Legend:
- ✅ **Perfect** — formatting, styles, and formulas are kept as-is.
- ⚠️ **Rendered as code** — math formulas appear as raw LaTeX and must be rebuilt inside Word/WPS.
- ⚠️ **Formulas missing** — math formulas are removed; rebuild them manually with the equation editor.

Test description:
1. **Copy Markdown** — use the “Copy” button provided beneath most AI responses (typically Markdown, sometimes HTML).
2. **Copy page content** — manually select the AI reply and copy (HTML rich text).

---

## Getting Started

1. Download an executable from the [Releases page](https://github.com/RICHQAQ/PasteMD/releases/):
   - ~~**PasteMD_vx.x.x.exe** — portable build, requires Pandoc to be installed and accessible from `PATH`.~~ (no longer provided; please build from source if needed)
   - **PasteMD_pandoc-Setup.exe** — bundled installer that ships with Pandoc and works out of the box.
2. Open Word, WPS, or Excel and place the caret where you want to paste.
3. Copy Markdown or HTML-rich text, then press the global hotkey (`Ctrl+Shift+B` by default).
4. PasteMD will:
   - Send Markdown tables to Excel (when Excel is already open).
   - Convert regular Markdown/HTML to DOCX and insert it into Word/WPS.
5. A notification in the tray (and optional toast) confirms success or failure.

---

## Configuration

The first launch creates a `config.json` file in the user data directory (Windows: `%APPDATA%\\PasteMD\\config.json`， MacOS: `~/Library/Application Support/PasteMD/config.json`). Edit it directly, then use the tray menu item **“Reload config/hotkey”** to apply changes instantly.

```json
{
  "hotkey": "<ctrl>+<shift>+b",
  "pandoc_path": "pandoc",
  "reference_docx": null,
  "save_dir": "%USERPROFILE%\\Documents\\pastemd",
  "keep_file": false,
  "notify": true,
  "enable_excel": true,
  "excel_keep_format": true,
  "no_app_action": "open",
  "md_disable_first_para_indent": true,
  "html_disable_first_para_indent": true,
  "html_formatting": {
    "strikethrough_to_del": true
  },
  "move_cursor_to_end": true,
  "Keep_original_formula": false,
  "language": "zh-CN",
  "pandoc_request_headers": [
    "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
  ],
  "pandoc_filters": []
}
```

Key fields:

- `hotkey` — global shortcut syntax such as `<ctrl>+<alt>+v`.
- `pandoc_path` — executable name or absolute path for Pandoc.
- `reference_docx` — optional style template consumed by Pandoc.
- `save_dir` — directory used when generated DOCX files are kept.
- `keep_file` — store converted DOCX files to disk instead of deleting them.
- `notify` — show system notifications when conversions finish.
- `enable_excel` — detect Markdown tables and paste them into Excel automatically.
- `excel_keep_format` — attempt to preserve bold/italic/code styles inside Excel.
- `no_app_action` — action when no target app is detected. Values: `open` (auto open), `save` (save only), `clipboard` (copy file to clipboard), `none` (no action). Default: `open`.
- `md_disable_first_para_indent` / `html_disable_first_para_indent` — normalize the first paragraph style to body text.
- `html_formatting` — options for formatting HTML rich text before conversion.
  - `strikethrough_to_del` — convert strikethrough ~~ to `<del>` tags for proper rendering.
- `move_cursor_to_end` — move the caret to the end of the inserted result.
- `Keep_original_formula` — keep original math formulas (in LaTeX code form).
- `language` — UI language: `en-US`, `zh-CN`, or `ja-JP`.
- `pandoc_request_headers` — request headers passed to Pandoc as `--request-header` when fetching remote resources (e.g. images). Example: `["User-Agent: ...", "Referer: https://www.oschina.net/"]`. Set to `[]` to disable request headers.
- **`pandoc_filters`** — **✨ New feature** - Custom Pandoc Filter list. Add `.lua` scripts or executable file paths; filters execute in list order. Extends Pandoc conversion with custom format processing, special syntax transformation, etc. Default: empty list. Example: `["%APPDATA%\\npm\\mermaid-filter.cmd"]` for Mermaid diagram support.

---

## 🔧 Advanced: Custom Pandoc Filters

### What are Pandoc Filters?

Pandoc Filters are plugin programs that process document content during conversion. PasteMD supports configuring multiple filters that execute sequentially to extend functionality.

### Use Case Example: Mermaid Diagram Support

To use Mermaid diagrams in Markdown and convert them properly to Word, you can use [mermaid-filter](https://github.com/raghur/mermaid-filter).

**1. Install mermaid-filter**

```bash
npm install --global mermaid-filter
```

*Prerequisite: [Node.js](https://nodejs.org/) must be installed*

<details>
<summary>⚠️ <b>Troubleshooting: Chrome Download Failure</b></summary>

Installing mermaid-filter requires downloading Chromium browser. If automatic download fails, you can download it manually:

**Step 1: Find Required Chromium Version**

Check the file: `%APPDATA%\npm\node_modules\mermaid-filter\node_modules\puppeteer-core\lib\cjs\puppeteer\revisions.d.ts`

Find content like:
```typescript
chromium: "1108766";
```
Or in the error message, e.g.:

```bash
npm error Error: Download failed: server returned code 502. URL: https://npmmirror.com/mirrors/chromium-browser-snapshots/Win_x64/1108766/chrome-win.zip
```
Find version like `Win_x64/1108766`.

Note down this version number (e.g., `1108766`).

**Step 2: Download Chromium**

Based on the version number from Step 1, download the corresponding Chromium:

```
https://storage.googleapis.com/chromium-browser-snapshots/Win_x64/1108766/chrome-win.zip
```

(Replace `1108766` in the URL with your version number)

**Step 3: Extract to Designated Directory**

Extract the downloaded `chrome-win.zip` to:

```
%USERPROFILE%\.cache\puppeteer\chrome\win64-1108766\chrome-win
```

(Replace `1108766` in the path with your version number)

After extraction, `chrome.exe` should be located at:  
`%USERPROFILE%\.cache\puppeteer\chrome\win64-1108766\chrome-win\chrome.exe`

</details>

**2. Configure in PasteMD**

Option 1: Via Settings UI
- Open PasteMD Settings → Conversion Tab → Pandoc Filters
- Click "Add..." button
- Select filter file: `%APPDATA%\npm\mermaid-filter.cmd`
- Save settings

Option 2: Edit config file
```json
{
  "pandoc_filters": [
    "%APPDATA%\\npm\\mermaid-filter.cmd"
  ]
}
```

**3. Test It Out**

Copy the following Markdown and convert with PasteMD:

~~~markdown
```mermaid
graph LR
    A[Start] --> B[Process]
    B --> C[End]
```
~~~

The Mermaid diagram will be rendered as an image and inserted into Word.

### More Filter Resources

- [Official Pandoc Filters List](https://github.com/jgm/pandoc/wiki/Pandoc-Filters)
- [Lua Filters Documentation](https://pandoc.org/lua-filters.html)

---

## Tray Menu

- Show the current global hotkey (read-only).
- Enable/disable the hotkey.
- Toggle notifications, set the action when no target app is detected, and toggle moving the caret to the end after paste.
- Enable or disable Excel-specific features and formatting preservation.
- Toggle keeping generated DOCX files.
- HTML Formatting: toggle conversion of strikethrough ~~ to `<del>` tags for proper rendering.
- Keep_original_formula: Whether to preserve the original mathematical formula in its LaTeX code form.
- Open save directory, view logs, edit configuration, or reload hotkeys.
- Check for updates and view installed version.
- Quit PasteMD.

---

## Build From Source

Recommended environment: Python 3.12 (64-bit).

```bash
pip install -r requirements.txt
python main.py
```

Packaged build (PyInstaller):

```bash
pyinstaller --clean -F -w -n PasteMD
  --icon assets\icons\logo.ico
  --add-data "assets\icons;assets\icons"
  --add-data "pastemd\i18n\locales\*.json;pastemd\i18n\locales"
  --add-data "pastemd\lua;pastemd\lua"
  --hidden-import plyer.platforms.win.notification
  main.py
```

The compiled executable will be placed in `dist/PasteMD.exe`.

---

## ⭐ Star

Every star helps — thank you for sharing PasteMD with more users.

<img src="../../docs/gif/atri/likeyou.gif"
     alt="like you"
     width="150">

[![Star History Chart](https://api.star-history.com/svg?repos=RICHQAQ/PasteMD&type=date&legend=top-left)](https://www.star-history.com/#RICHQAQ/PasteMD&type=date&legend=top-left)

---

## ☕ Support & Donation


If PasteMD saves you time, consider buying the author a coffee — your support helps prioritize fixes, enhancements, and new integrations.

Also welcome to join the **PasteMD User Group** for discussion and support:

<div align="center">
  <img src="../../docs/img/qrcode.jpg" alt="PasteMD QQ Group QR Code" width="200" />
  <br>
  <sub>Scan to join the PasteMD QQ group</sub>
</div>

<img src="../../docs/gif/atri/flower.gif"
     alt="give you a flower"
     width="150">

| Alipay | WeChat |
| --- | --- |
| ![Alipay](../../docs/pay/Alipay.jpg) | ![WeChat](../../docs/pay/Weixinpay.png) |

---

## License

This project is released under the [MIT License](../../LICENSE).
Third-party licenses are listed in [THIRD_PARTY_NOTICES.md](../../THIRD_PARTY_NOTICES.md).
