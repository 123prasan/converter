import sys
import os
import subprocess
import platform

def log(message):
    """Sends real-time updates to Node.js backend"""
    print(message)
    sys.stdout.flush()

def convert_windows(input_path, output_path):
    """Windows specific conversion using Aspose.Cells"""
    try:
        import aspose.cells as cells
        log("Status: Utilizing Aspose High-Fidelity Engine...")
        
        workbook = cells.Workbook(input_path)
        save_options = cells.PdfSaveOptions()
        
        # Ensure professional layout: prevents columns from splitting across pages
        save_options.all_columns_in_one_page_per_sheet = True
        
        workbook.save(output_path, save_options)
        return True
    except ImportError:
        log("Error: Aspose.Cells not installed. Run 'pip install aspose-cells'")
        return False
    except Exception as e:
        log(f"Aspose Error: {str(e)}")
        return False

def convert_linux(input_path, output_dir):
    """Linux specific conversion using LibreOffice (soffice)"""
    try:
        log("Status: Utilizing LibreOffice Headless Engine (Linux)...")
        # LibreOffice command to convert calc files to pdf
        process = subprocess.run([
            'soffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            input_path
        ], capture_output=True, text=True)
        
        if process.returncode == 0:
            return True
        else:
            log(f"LibreOffice Error: {process.stderr}")
            return False
    except Exception as e:
        log(f"Linux Subprocess Error: {str(e)}")
        return False

def main():
    if len(sys.argv) < 3:
        log("Error: Missing file paths.")
        sys.exit(1)

    input_path = os.path.abspath(sys.argv[1])
    output_path = os.path.abspath(sys.argv[2])
    output_dir = os.path.dirname(output_path)
    
    current_os = platform.system()
    log(f"Process: Environment detected as {current_os}")

    success = False
    if current_os == "Windows":
        success = convert_windows(input_path, output_path)
    else:
        # LibreOffice automatically names the file. 
        # We ensure it lands in the right directory.
        success = convert_linux(input_path, output_dir)
        
        # Cleanup: LibreOffice might name it based on input filename, 
        # so we may need to rename it to match the requested output_path
        generated_name = os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
        generated_path = os.path.join(output_dir, generated_name)
        
        if success and os.path.exists(generated_path) and generated_path != output_path:
            if os.path.exists(output_path): os.remove(output_path)
            os.rename(generated_path, output_path)

    if success:
        log("Success: Excel converted to high-fidelity PDF.")
        sys.exit(0)
    else:
        log("Fail: Excel Engine could not process the file.")
        sys.exit(1)

if __name__ == "__main__":
    main()