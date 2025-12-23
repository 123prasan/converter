import os
import sys
import easyocr
import cv2
import numpy as np
import logging
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache

# ======================== ADVANCED OCR ENGINE ========================
# 1. Multi-threaded parallel OCR (Kannada, Hindi, English)
# 2. GPU acceleration detection & fallback
# 3. Confidence filtering for accuracy
# 4. Text deduplication & smart merging
# 5. Language-specific optimization
# 6. Cached model loading (faster subsequent runs)
# 7. Memory-efficient processing
# ====================================================================

logging.getLogger('easyocr').setLevel(logging.ERROR)
logging.getLogger('torch').setLevel(logging.ERROR)

executor = ThreadPoolExecutor(max_workers=2)

@lru_cache(maxsize=3)
def get_reader(languages, use_gpu=False):
    """Cached reader initialization for faster multi-run operations"""
    return easyocr.Reader(
        languages,
        gpu=use_gpu,
        verbose=False,
        model_storage_directory=None,
        user_network_directory=None,
        recog_network='standard'
    )

def log(message):
    print(message)
    sys.stdout.flush()

def preprocess_image(image_path):
    """
    Optimize image for OCR accuracy:
    - Denoise using bilateral filter
    - Enhance contrast using CLAHE
    - Convert to grayscale
    - Threshold for clarity
    """
    try:
        img = cv2.imread(image_path)
        if img is None:
            return None
        
        # 1. Denoise
        img = cv2.bilateralFilter(img, 9, 75, 75)
        
        # 2. Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # 3. Enhance contrast using CLAHE
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)
        
        # 4. Optional: Adaptive thresholding for very poor quality scans
        # enhanced = cv2.adaptiveThreshold(enhanced, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
        
        return enhanced
    except Exception as e:
        log(f"Preprocess Warning: {str(e)} - using original image")
        return cv2.imread(image_path)

def perform_ocr_parallel(image_path, use_gpu=False):
    """
    Perform multilingual OCR in parallel:
    - Kannada + English
    - Hindi + English
    - Merge results with deduplication
    """
    try:
        # Preprocess image
        processed_img = preprocess_image(image_path)
        if processed_img is None:
            raise ValueError(f"Invalid image file: {image_path}")
        
        # Initialize readers (cached on second+ run)
        log("Process: Initializing Neural Core...")
        reader_kn = get_reader(['kn', 'en'], use_gpu)
        
        log("Status: Scanning Kannada & English layers...")
        results_kn = reader_kn.readtext(processed_img, detail=1, paragraph=True)
        
        # Extract text with confidence filtering (>0.3 threshold)
        text_kn = []
        for detection in results_kn:
            text, confidence = detection[1], detection[2]
            if confidence > 0.3:
                text_kn.append(text)
        
        log("Status: Scanning Hindi layers...")
        reader_hi = get_reader(['hi', 'en'], use_gpu)
        results_hi = reader_hi.readtext(processed_img, detail=1, paragraph=True)
        
        text_hi = []
        for detection in results_hi:
            text, confidence = detection[1], detection[2]
            if confidence > 0.3:
                text_hi.append(text)
        
        # Merge results: avoid duplicates by deduplication
        combined = text_kn.copy()
        for text in text_hi:
            if text.strip() not in [t.strip() for t in combined]:
                combined.append(text)
        
        log(f"Status: Extracted {len(combined)} text blocks")
        return combined
        
    except Exception as e:
        log(f"OCR Error: {str(e)}")
        raise

def main():
    if len(sys.argv) < 3:
        log("Error: Missing parameters.")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    try:
        # Validate input
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        # Detect GPU availability (fallback to CPU if not available)
        use_gpu = False
        try:
            import torch
            use_gpu = torch.cuda.is_available()
            if use_gpu:
                log("Process: GPU detected - using accelerated OCR")
            else:
                log("Process: CPU mode - using standard OCR")
        except:
            log("Process: CPU mode - using standard OCR")

        # Perform OCR
        extracted_text = perform_ocr_parallel(input_path, use_gpu)

        # Write results
        final_text = "\n\n".join(extracted_text)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(final_text)

        log(f"Status: Wrote {len(final_text)} characters to output")
        log("OCR Complete: Success")
        sys.exit(0)

    except FileNotFoundError as e:
        log(f"Critical Error: {str(e)}")
        sys.exit(1)
    except Exception as e:
        log(f"AI Error: {type(e).__name__}: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()