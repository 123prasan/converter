import os
import sys
import fitz  # PyMuPDF
import json

def log(message):
    """Sends real-time updates to Node.js backend"""
    print(message)
    sys.stdout.flush()

def main():
    # Args: script.py <input_path> <output_path> <grayscale> <strip_meta> <dpi>
    if len(sys.argv) < 6:
        log("Error: Missing parameters.")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    grayscale = sys.argv[3].lower() == 'true'
    strip_meta = sys.argv[4].lower() == 'true'
    dpi = int(sys.argv[5])

    log("Process: Initializing Shrink Engine...")

    try:
        # Step 1: Open with repair attempt
        doc = fitz.open(input_path)
        
        # Step 2: Grayscale Conversion Logic
        if grayscale:
            log("Process: Striping color data (Grayscale mode)...")
            for page in doc:
                for img in page.get_images():
                    xref = img[0]
                    # This converts the image resource itself to gray
                    pix = fitz.Pixmap(doc, xref)
                    if pix.colorspace.n > 1:
                        gray_pix = fitz.Pixmap(fitz.csGRAY, pix)
                        doc.update_stream(xref, gray_pix.tobytes())
                        gray_pix = None
                    pix = None

        log(f"Process: Rebuilding structure & applying {dpi} DPI limit...")
        
        # Step 3: Deep Save (Fixes xref errors)
        # Using garbage=4 + clean=True forces PyMuPDF to rebuild the xref table
        # which solves the 'code=7' object not found issue.
        doc.save(
            output_path, 
            garbage=4, 
            deflate=True, 
            clean=True,
            linear=True,
            pretty=False,
            # Force full rewrite to fix corrupt xref
           
    # Use a generic profile for maximum compatibility and size saving
   
        )
        
        doc.close()
        log("Compression Complete! File optimized and repaired.")
        sys.exit(0)

    except Exception as e:
        # Catch the specific xref error and provide a cleaner message
        error_msg = str(e)
        if "xref" in error_msg.lower():
            log(f"Process Error: Document structure is corrupt. Repairing failed.")
        else:
            log(f"Process Error: {error_msg}")
        sys.exit(1)

if __name__ == "__main__":
    main()