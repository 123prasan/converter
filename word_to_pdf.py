import sys
import os
import subprocess
import platform

def log(message):
    print(message)
    sys.stdout.flush()

def convert_windows(input_path, output_path):
    """Windows specific conversion using Microsoft Word COM"""
    try:
        import pythoncom
        from docx2pdf import convert
        pythoncom.CoInitialize()
        convert(input_path, output_path)
        return True
    except Exception as e:
        log(f"Windows Error: {str(e)}")
        return False

def convert_linux(input_path, output_dir):
    """Linux specific conversion using LibreOffice (Headless)"""
    try:
        # LibreOffice typically uses 'libreoffice' or 'soffice' command
        command = 'soffice' 
        process = subprocess.run([
            command,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            input_path
        ], capture_output=True, text=True)
        
        return process.returncode == 0
    except Exception as e:
        log(f"Linux Error: {str(e)}")
        return False

def main():
    if len(sys.argv) < 3:
        sys.exit(1)

    input_path = os.path.abspath(sys.argv[1])
    output_path = os.path.abspath(sys.argv[2])
    output_dir = os.path.dirname(output_path)
    
    current_os = platform.system()
    log(f"Process: Environment detected as {current_os}")

    success = False
    if current_os == "Windows":
        log("Status: Utilizing Microsoft Word COM Engine...")
        success = convert_windows(input_path, output_path)
    else:
        log("Status: Utilizing LibreOffice Headless Engine...")
        success = convert_linux(input_path, output_dir)

    if success:
        log("Success: Conversion completed dynamically.")
        sys.exit(0)
    else:
        log("Fail: Engine failed to process document.")
        sys.exit(1)

if __name__ == "__main__":
    main()