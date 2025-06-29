@echo off
REM =================================================================
REM        PDF Converter Right-Click Menu Installer
REM =================================================================
ECHO.
ECHO This script will add a "Convert to PDF" option to your right-click menu.
ECHO It needs to find your Python installation to work correctly.
ECHO.

REM --- Find the Python Executable ---
FOR /F "usebackq delims=" %%i IN (`where pythonw.exe`) DO (
    SET "PYTHON_PATH=%%i"
    GOTO :FoundPython
)

ECHO ERROR: Could not find pythonw.exe.
ECHO Please make sure Python is installed and you checked "Add Python to PATH" during installation.
ECHO.
pause
GOTO :EOF

:FoundPython
ECHO [+] Found Python at: %PYTHON_PATH%
ECHO.

REM --- Get the path of the converter script in the same directory ---
SET "SCRIPT_PATH=%~dp0convert_to_pdf.py"
IF NOT EXIST "%SCRIPT_PATH%" (
    ECHO ERROR: convert_to_pdf.py not found in the same folder as this installer.
    ECHO Please make sure both files are in the same directory.
    ECHO.
    pause
    GOTO :EOF
)
ECHO [+] Found script at: %SCRIPT_PATH%
ECHO.

REM --- Create the registry entries ---
ECHO [*] Adding keys to the Windows Registry...

REM Note: The command paths are escaped with double backslashes for registry compatibility.
SET "PYTHON_REG_PATH=%PYTHON_PATH:\=\\%"
SET "SCRIPT_REG_PATH=%SCRIPT_PATH:\=\\%"

REG ADD "HKEY_CLASSES_ROOT\*\shell\PDFConverter" /v "" /t REG_SZ /d "Convert to PDF" /f > NUL
REG ADD "HKEY_CLASSES_ROOT\*\shell\PDFConverter" /v "Icon" /t REG_SZ /d "imageres.dll,68" /f > NUL
REG ADD "HKEY_CLASSES_ROOT\*\shell\PDFConverter\command" /v "" /t REG_SZ /d "\"%PYTHON_REG_PATH%\" \"%SCRIPT_REG_PATH%\" \"%%1\"" /f > NUL

IF %ERRORLEVEL% EQU 0 (
    ECHO.
    ECHO =================================================================
    ECHO      SUCCESS! "Convert to PDF" has been added.
    ECHO =================================================================
    ECHO.
    ECHO You may need to restart File Explorer for the change to appear.
    ECHO You can close this window now.
) ELSE (
    ECHO.
    ECHO ERROR: Failed to write to the registry.
    ECHO Please make sure you ran this script "as an administrator".
)

ECHO.
pause
