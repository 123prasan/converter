import os
import sys
import easyocr
import cv2
import numpy as np
import logging

# Disable EasyOCR's internal verbose logging and progress bars
logging.getLogger('easyocr').setLevel(logging.ERROR)

def log(message):
    print(message)
    sys.stdout.flush()

def main():
    if len(sys.argv) < 3:
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    try:
        # Step 1: Initialize Reader (download_enabled=True is default, but we hide output)
        log("Process: Initializing Neural Core...")
        
        # verbose=False stops the |███---| progress bars from printing to stdout
        reader_kn = easyocr.Reader(['kn', 'en'], gpu=False, verbose=False)
        
        log("Status: Scanning Kannada & English layers...")
        img = cv2.imread(input_path)
        res_kn = reader_kn.readtext(img, detail=0, paragraph=True)

        log("Status: Scanning Hindi layers...")
        reader_hi = easyocr.Reader(['hi', 'en'], gpu=False, verbose=False)
        res_hi = reader_hi.readtext(img, detail=0, paragraph=True)
        
        # Combine results
        extracted = res_kn + [p for p in res_hi if p not in res_kn]
        final_text = "\n\n".join(extracted)

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(final_text)

        log("OCR Complete: Success")
        sys.exit(0)

    except Exception as e:
        log(f"AI Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()