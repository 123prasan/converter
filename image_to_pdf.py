import os
import sys
import img2pdf
from PIL import Image

# -------------------------
# REAL-TIME LOGGING HELPER
# -------------------------
def log(message):
    """Sends real-time updates to Node.js backend via stdout"""
    print(message)
    sys.stdout.flush()

# -------------------------
# MAIN CONVERSION ENGINE
# -------------------------
def main():
    # Expecting only 2 arguments now: <input_path> <output_path>
    if len(sys.argv) < 3:
        log("Error: Missing parameters for Image processing.")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    log(f"Job Started: Neural Image Engine")

    try:
        # 1. Validation & Pre-processing
        log("Process: Analyzing image dimensions and DPI...")
        
        # We use PIL to ensure the image is valid before img2pdf handles it
        with Image.open(input_path) as img:
            img_format = img.format
            log(f"Status: Detected {img_format} source format.")

        # 2. Conversion
        log("Process: Wrapping image into lossless PDF container...")
        
        # img2pdf is the fastest way to convert without losing quality
        with open(output_path, "wb") as f:
            f.write(img2pdf.convert(input_path))
        
        log("Conversion Complete!")
        sys.exit(0)

    except Exception as e:
        log(f"Critical Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()