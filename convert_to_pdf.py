# /convert_to_pdf.py

import os
import sys
import traceback
import datetime

# --- VERY FIRST THING: SETUP LOGGING ---
# This runs before any other imports that might fail, ensuring we can log import errors.
try:
    log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "conversion_log.txt")
except NameError:
    # If __file__ is not defined (e.g., in some environments), use current working directory.
    log_file_path = os.path.join(os.getcwd(), "conversion_log.txt")

def log_message(message_type, message):
    """Writes a message to the log file with a timestamp and error details if applicable."""
    try:
        with open(log_file_path, "a", encoding='utf-8') as f:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"--- {timestamp} [{message_type}] ---\n")
            f.write(str(message) + "\n")
            if message_type == "ERROR":
                # Write the full traceback for detailed debugging.
                traceback.print_exc(file=f)
            f.write("\n")
    except Exception:
        # Failsafe if logging itself fails. This is a last resort.
        pass

# --- START THE SCRIPT ---
log_message("INFO", f"Script started. Arguments received: {sys.argv}")

try:
    # --- NOW, try to import the risky library ---
    import win32com.client
    log_message("INFO", "Successfully imported win32com.client.")
except ImportError:
    log_message("ERROR", "Failed to import win32com.client. This is a critical error. Please ensure the 'pywin32' library is installed and its post-install script has been run as an administrator.")
    sys.exit() # Exit the script because nothing else will work.

# --- Conversion Functions ---
def convert_word_to_pdf(doc_path, pdf_path):
    """Automates MS Word to convert a .doc or .docx file to .pdf"""
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, FileFormat=17) # 17 = wdFormatPDF
        doc.Close()
        log_message("SUCCESS", f"Converted Word file: {os.path.basename(doc_path)}")
        return True
    except Exception as e:
        log_message("ERROR", f"Failed during Word conversion for file: {doc_path}\nError: {e}")
        return False
    finally:
        if word:
            word.Quit()

def convert_ppt_to_pdf(ppt_path, pdf_path):
    """Automates MS PowerPoint to convert a .pptx file to .pdf"""
    powerpoint = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
        presentation.SaveAs(pdf_path, FileFormat=32) # 32 = ppSaveAsPDF
        presentation.Close()
        log_message("SUCCESS", f"Converted PowerPoint file: {os.path.basename(ppt_path)}")
        return True
    except Exception as e:
        log_message("ERROR", f"Failed during PowerPoint conversion for file: {ppt_path}\nError: {e}")
        return False
    finally:
        if powerpoint:
            powerpoint.Quit()

# --- Main Logic ---
def main():
    if len(sys.argv) < 2:
        log_message("WARN", "Script was run without a file argument. To use, right-click on a supported file and select 'Convert to PDF'.")
        return

    file_to_convert = sys.argv[1]
    log_message("INFO", f"Processing file: {file_to_convert}")

    if not os.path.exists(file_to_convert):
        log_message("ERROR", f"Input file not found: {file_to_convert}")
        return
        
    base_name, extension = os.path.splitext(file_to_convert)

    supported_word_formats = ['.doc', '.docx']
    supported_ppt_formats = ['.ppt', '.pptx']

    if extension.lower() in supported_word_formats:
        pdf_path = f"{base_name}.pdf"
        convert_word_to_pdf(file_to_convert, pdf_path)
    elif extension.lower() in supported_ppt_formats:
        pdf_path = f"{base_name}.pdf"
        convert_ppt_to_pdf(file_to_convert, pdf_path)
    else:
        log_message("WARN", f"Unsupported file type skipped: {extension}")
        return

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log_message("FATAL", f"An unhandled exception occurred in the main execution block: {e}")
