import os
import sys
import fitz  # PyMuPDF

def log(message):
    """Sends real-time updates to Node.js backend"""
    print(message)
    sys.stdout.flush()

def main():
    # Args: script.py <output_path> <input_file_1> <input_file_2> ...
    if len(sys.argv) < 3:
        log("Error: At least two files are required for merging.")
        sys.exit(1)

    output_path = sys.argv[1]
    input_files = sys.argv[2:]

    log("Process: Initializing Merge Engine...")

    try:
        # 1. Create a blank destination PDF
        result = fitz.open()

        log(f"Process: Merging {len(input_files)} documents...")

        for i, file_path in enumerate(input_files):
            if not os.path.exists(file_path):
                continue
            
            log(f"Status: Stitching document {i+1}...")
            # 2. Open each input PDF and insert it into the result
            with fitz.open(file_path) as m_file:
                result.insert_pdf(m_file)

        # 3. Save the final merged document
        log("Process: Finalizing unified document stream...")
        result.save(output_path, garbage=4, deflate=True)
        result.close()

        log("Success: PDF Merge Complete.")
        sys.exit(0)

    except Exception as e:
        log(f"Critical Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()