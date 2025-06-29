# Office to PDF Right-Click Converter

A simple Windows utility that adds a "Convert to PDF" option to the right-click context menu. It allows for quick, one-click conversion of Microsoft Office files (.docx, .doc, .pptx) to PDF.

 <!-- You can create a gif and upload it to show a demo -->

## Features
- **One-Click Conversion:** Right-click any supported file and select "Convert to PDF".
- **Silent Operation:** No extra windows pop up during conversion.
- **Automatic Placement:** The new PDF is saved in the same folder as the original file.
- **Robust Logging:** If a conversion fails, a `conversion_log.txt` file is created to help diagnose the issue.

## Requirements
- Windows 10 or 11
- Microsoft Office (Word and/or PowerPoint) installed
- Python 3 installed (make sure to check the "Add Python to PATH" option during installation)

## Installation Guide

Follow these three steps carefully to set up the converter.

### Step 1: Download the Files
Clone this repository or download the files as a ZIP and extract them to a permanent location on your computer (e.g., `C:\Tools\PDFConverter`).
```bash
git clone https://github.com/YourUsername/YourRepoName.git
```

### Step 2: Install Required Python Library
This tool depends on the `pywin32` library to control Microsoft Office applications.

1. Open **Command Prompt** or **PowerShell**.
2. Run the following command to install the library:
   ```bash
   pip install pywin32
   ```

### Step 3: Run the Post-Installation Script (Crucial!)
After installing, you must register the library's components with Windows. This is a common source of errors if skipped.

1. Open **Command Prompt** or **PowerShell as an Administrator**.
2. Copy and paste the following command and press Enter. It automatically finds and runs the required script for you.
   ```powershell
   python -c "import sys, os; from pathlib import Path; p = Path(sys.executable).parent / 'Scripts' / 'pywin32_postinstall.py'; os.system(f'python {p} -install')"
   ```

### Step 4: Run the Installer
Now you can add the "Convert to PDF" option to your right-click menu.

1. Navigate to the folder where you saved the files.
2. **Right-click** on `install.bat`.
3. Select **"Run as administrator"**.
4. Follow the prompts in the window that appears.

The installation is complete! The option will now be available in your right-click menu.

## How to Uninstall
To remove the "Convert to PDF" option from your context menu, simply run the following command in Command Prompt:
```bash
REG DELETE "HKEY_CLASSES_ROOT\*\shell\PDFConverter" /f
```
