# SuperEbookMachine

**SuperEbookMachine** is a Windows desktop app for **batch converting ebooks (especially PDFs)** into Kindle-friendly formats while **preserving your folder structure**.

It’s a simple GUI wrapper around **Calibre’s `ebook-convert`** that adds:
- recursive scanning
- bulk conversion
- progress tracking
- parallel workers
- easy “Kindle App vs Physical Kindle” workflow presets

---

## Features

- 📁 **Recursive scanning** through subfolders
- 🧠 **Preserve folder structure** in the output directory
- 🎛️ **Choose input types** (`.pdf`, `.epub`, `.mobi`, etc.) or specify custom extensions
- 🎯 **Target presets**:
  - **Kindle App** → outputs **EPUB** + uses Send-to-Kindle
  - **Physical Kindle (USB)** → outputs **AZW3**
- ⚡ **Parallel conversion** (multi-worker)
- 📊 **Progress bar** with `done/total` + OK/SKIP/FAIL stats
- 🖥️ Runs as a standalone `.exe` (no Python needed for end users)

---

## Requirements

### Calibre is required
SuperEbookMachine does **not** bundle Calibre.

It depends on Calibre’s command-line tool:

- `ebook-convert`

Download Calibre (free):
https://calibre-ebook.com

> You do **not** need to open the Calibre GUI — SuperEbookMachine only calls `ebook-convert` in the background.

---

## Installation (Windows)

1. Install **Calibre**
2. Download and run:
   - `SuperEbookMachine.exe`

If the app can’t find `ebook-convert`, set the path manually (common location):


---

## Usage

1. **Input root folder**  
   Select the folder containing your library (including subfolders).

2. **Output root folder**  
   Select where converted files should be written.

3. **Target**
   - **Kindle App (Phone/Desktop/Cloud)**  
     - Output format: **EPUB**
     - Upload using **Send-to-Kindle**
   - **Physical Kindle (USB)**  
     - Output format: **AZW3**
     - Copy to `Kindle/documents/` via USB

4. **Scan input file types**
   - Check formats, or enable custom extensions (comma-separated)

5. Click **Start**

---

## Output behavior (folders preserved)

SuperEbookMachine mirrors the folder tree from input to output.

Example:

Input/
Papers/
Immunology/
paper1.pdf
paper2.pdf

Output/
Kindle/
Immunology/
paper1.epub
paper2.epub


Existing outputs are skipped unless **Overwrite** is enabled.

---

## Sending to Kindle (EPUB workflow)

For Kindle apps (iOS/Android/Desktop), the recommended workflow is:

1. Export as **EPUB**
2. Upload via **Send-to-Kindle**:
   https://www.amazon.com/sendtokindle

Amazon converts the EPUB in the cloud and syncs it across your Kindle devices/apps.

---

## Loading onto a physical Kindle (AZW3 workflow)

AZW3 is for physical Kindles via USB:

1. Connect Kindle via USB
2. Open the Kindle drive
3. Copy `.azw3` files into:


4. Eject Kindle

---

## Troubleshooting

### “ebook-convert not found”
- Install Calibre
- Or browse to the executable:
  `C:\Program Files\Calibre2\ebook-convert.exe`

### Some files fail to convert
This is normal for certain PDFs (scans, encryption, corrupted metadata, unusual layouts).
Failures are logged and the batch continues.

### Kindle App won’t open AZW3
Correct — Kindle apps generally require EPUB via Send-to-Kindle.
Use the **Kindle App** target preset.

---

## Development / Building an EXE

Packaged using PyInstaller:

```powershell
pyinstaller --onefile --windowed --name "SuperEbookMachine" SuperEbookMachine.py

dist\SuperEbookMachine.exe
```
## License

Provided as-is for personal use.

Calibre is a separate project with its own license and is not bundled with this app.
