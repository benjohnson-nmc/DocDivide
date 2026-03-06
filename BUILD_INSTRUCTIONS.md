# Build Instructions -- DocDivide Installer

## What you'll produce
`Output\DocDivideSetup.exe` -- a standard Windows installer (~80-120 MB) that:
- Installs to `Program Files\DocDivide`
- Creates a Start Menu shortcut
- Includes an optional desktop shortcut
- Bundles Python, all libraries, and Poppler (no prerequisites for end users)
- Includes an uninstaller

---

## Prerequisites (build machine only)

| Tool | Where to get it |
|------|----------------|
| Python 3.10+ (64-bit) | https://python.org |
| Poppler for Windows | https://github.com/oschwartz10612/poppler-windows/releases |
| PyInstaller | `pip install pyinstaller` |
| Inno Setup 6 | https://jrsoftware.org/isdl.php |

---

## Step-by-step

### 1. Set up your project folder

```
DocDivide/
  docdivide.py
  docdivide.spec
  installer.iss
  BUILD_INSTRUCTIONS.md
  requirements.txt
  icon.svg          <- convert to icon.ico before building
```

### 2. Convert icon.svg to icon.ico

1. Open icon.svg in a browser to preview it
2. Go to https://convertio.co/svg-ico/ and upload icon.svg
3. Download the result and save as `icon.ico` in the project folder
   (Ideally generate multi-size: 256x256, 64x64, 32x32, 16x16)

### 3. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 4. Download and configure Poppler

1. Download from: https://github.com/oschwartz10612/poppler-windows/releases
2. Extract to e.g. `C:\poppler`
3. Open `docdivide.spec` and update:
   ```python
   POPPLER_BIN = r"C:\poppler\Library\bin"
   ```

### 5. Build the executable

```bash
pyinstaller docdivide.spec
```

Test: run `dist\DocDivide\DocDivide.exe` before continuing.

### 6. Build the installer

1. Open Inno Setup Compiler
2. File > Open > select `installer.iss`
3. Update `#define AppPublisher` to your company name
4. Press F9 to compile
5. Output: `Output\DocDivideSetup.exe`

---

## Distributing to users

Share `DocDivideSetup.exe`. Users run it and click through the wizard.
Each user needs their own Anthropic API key from https://console.anthropic.com

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| poppler not found | Check POPPLER_BIN path in spec file |
| Antivirus flags exe | Code-sign with a certificate, or submit for AV whitelisting |
| App crashes silently | Set `console=True` in spec to see errors |
| ModuleNotFoundError | Add module to hiddenimports in spec |
| Large installer (~100MB) | Normal -- Python runtime is included |
