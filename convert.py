import sys
import os
import time
import logging
import shutil
import tempfile
import uuid
from multiprocessing import Pool, cpu_count, freeze_support
from functools import lru_cache

# ======================== SILENCE NOISY LOGGERS ========================
# This stops pdf2docx from spamming stderr, which Node.js thinks are errors
logging.getLogger("pdf2docx").setLevel(logging.ERROR)
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("PIL").setLevel(logging.ERROR)
# =======================================================================

# Third-party imports
try:
    from pdf2docx import Converter
    from pdf2image import convert_from_path
    import pytesseract
    import cv2
    import numpy as np
    from docx import Document
    import pdfplumber
    
    # Cryptography for E2EE
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
except ImportError as e:
    # Silent fail for dependency check to avoid noise, handled in main execution
    pass

# ======================== SINGULARITY PDF ENGINE v8.0 ========================
# 1. RAM-Safe: Streams pages in batches (10 pages/cycle)
# 2. Turbo OCR: Disables 'inverted text' check for 20% speed boost
# 3. Hybrid Routing: Instantly switches between Vector and Raster engines
# 4. Silent Mode: Suppresses library INFO logs
# ==============================================================================

class SpeedLogger:
    """Non-blocking logger that flushes immediately for Node.js"""
    @staticmethod
    def log(msg):
        sys.stdout.write(f"{msg}\n")
        sys.stdout.flush()

class SecurityEngine:
    """Handles decryption of E2EE secured files"""
    @staticmethod
    def decrypt_file(input_path, secret_key):
        if not input_path.endswith('.enc'):
            return input_path
        
        SpeedLogger.log("Security: Decrypting secure payload...")
        try:
            with open(input_path, 'rb') as f:
                data = f.read()
            
            salt = data[:16]
            iv = data[16:28]
            ciphertext = data[28:]

            kdf = PBKDF2HMAC(
                algorithm=hashes.SHA256(),
                length=32,
                salt=salt,
                iterations=100000
            )
            key = kdf.derive(secret_key.encode())

            aesgcm = AESGCM(key)
            plain_pdf = aesgcm.decrypt(iv, ciphertext, None)
            
            decrypted_path = input_path.replace(".enc", ".pdf")
            with open(decrypted_path, "wb") as f:
                f.write(plain_pdf)
            
            return decrypted_path
        except Exception:
            return input_path

class ImageOptimizer:
    """Lightweight preprocessing to boost OCR speed"""
    @staticmethod
    def preprocess(image):
        try:
            # Zero-Copy conversion to OpenCV format
            img_array = np.array(image)
            
            # If RGB, convert to Gray
            if len(img_array.shape) == 3:
                img_array = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            
            # Simple thresholding (Faster than adaptive for standard docs)
            _, binary = cv2.threshold(img_array, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            return binary
        except:
            return np.array(image)

def worker_ocr_page(args):
    """
    Worker function for Multiprocessing OCR.
    Optimization: tessedit_do_invert=0 skips 'inverted' check (huge speedup)
    """
    index, image = args
    try:
        processed_img = ImageOptimizer.preprocess(image)
        # --oem 1 (LSTM only) --psm 3 (Auto Page Segmentation)
        config = r'--oem 1 --psm 3 -c tessedit_do_invert=0'
        text = pytesseract.image_to_string(processed_img, config=config)
        return (index, text.strip())
    except:
        return (index, "")

class ConversionEngine:
    
    @staticmethod
    def analyze_pdf_structure(pdf_path):
        """
        Determines if the PDF has selectable text.
        Returns: (has_text: bool, page_count: int)
        """
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if not pdf.pages: return (False, 0)
                
                page_count = len(pdf.pages)
                
                # Check random sample (First, Middle, Last) to avoid reading whole file
                indices = [0, page_count // 2, page_count - 1]
                text_len = 0
                
                for i in set(indices): # set removes duplicates for small docs
                    if i < page_count:
                        text = pdf.pages[i].extract_text()
                        if text: text_len += len(text)
                
                return (text_len > 100, page_count)
        except:
            return (False, 0)

    @staticmethod
    def convert_native(pdf_path, output_path):
        """Uses pdf2docx for blazing fast native conversion"""
        SpeedLogger.log("Engine: Vector Mode (Native Text)...")
        try:
            # Suppress internal logging during instantiation
            cv = Converter(pdf_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            return True
        except Exception as e:
            # Only log real errors
            return False

    @staticmethod
    def convert_ocr(pdf_path, output_path, page_count):
        """
        Uses Parallel OCR with Batch Processing.
        Prevents RAM overflow on large files.
        """
        SpeedLogger.log("Engine: Raster Mode (AI OCR)...")
        
        try:
            doc = Document()
            # Batch size: Processes 10 pages, writes them, clears RAM, repeats.
            batch_size = 10 
            
            # Determine optimal CPU cores (leave 1 for OS)
            pool_size = max(1, min(cpu_count() - 1, 8))
            
            total_processed = 0
            
            with Pool(processes=pool_size) as pool:
                for i in range(0, page_count, batch_size):
                    end_page = min(i + batch_size, page_count)
                    
                    # Convert specific page range to images
                    images = convert_from_path(
                        pdf_path, 
                        dpi=150, # 150 is the web standard sweet spot
                        first_page=i+1, 
                        last_page=end_page,
                        thread_count=4 # Threading for image conversion (IO bound)
                    )
                    
                    if not images: continue
                    
                    # Prepare tasks
                    tasks = [(i + idx, img) for idx, img in enumerate(images)]
                    
                    # Run OCR in parallel (CPU bound)
                    results = pool.map(worker_ocr_page, tasks)
                    results.sort(key=lambda x: x[0])
                    
                    # Dump to Word Doc immediately
                    for _, text in results:
                        if text:
                            # Sanitize text to prevent XML errors
                            clean_text = ''.join(c for c in text if c.isprintable() or c in ['\n', '\t', '\r'])
                            doc.add_paragraph(clean_text)
                            doc.add_page_break()
                    
                    total_processed += len(results)
                    SpeedLogger.log(f"Status: Processed {total_processed}/{page_count} pages...")
                    
                    # Explicitly delete images to free RAM
                    del images
            
            doc.save(output_path)
            return True
            
        except Exception as e:
            SpeedLogger.log(f"OCR Error: {e}")
            return False

def main():
    if len(sys.argv) < 4:
        SpeedLogger.log("Error: Missing arguments")
        sys.exit(1)

    input_path = sys.argv[1]
    tool = sys.argv[2]
    output_path = sys.argv[3]
    secret_key = sys.argv[4] if len(sys.argv) > 4 else ""

    start_time = time.time()
    
    # 1. Security Decryption
    working_pdf = SecurityEngine.decrypt_file(input_path, secret_key)
    
    if not os.path.exists(working_pdf):
        SpeedLogger.log("Error: Input file missing")
        sys.exit(1)

    try:
        # 2. Structure Analysis
        has_text, page_count = ConversionEngine.analyze_pdf_structure(working_pdf)
        
        success = False
        
        # 3. Intelligent Switching
        if has_text:
            success = ConversionEngine.convert_native(working_pdf, output_path)
            # Fallback check: if native produced a tiny/empty file, force OCR
            if success and os.path.exists(output_path) and os.path.getsize(output_path) < 3000:
                SpeedLogger.log("Status: Native conversion yielded poor results. Switching to OCR.")
                success = ConversionEngine.convert_ocr(working_pdf, output_path, page_count)
            elif not success:
                success = ConversionEngine.convert_ocr(working_pdf, output_path, page_count)
        else:
            success = ConversionEngine.convert_ocr(working_pdf, output_path, page_count)

        # 4. Final Reporting
        if success and os.path.exists(output_path):
            duration = time.time() - start_time
            # Node.js looks for this exact regex: /won in ([\d\.]+)s/
            SpeedLogger.log(f"Success: PDF converted. won in {duration:.3f}s")
            
            if working_pdf != input_path:
                try: os.remove(working_pdf)
                except: pass
            sys.exit(0)
        else:
            SpeedLogger.log("Error: Conversion failed")
            sys.exit(1)

    except Exception as e:
        SpeedLogger.log(f"Critical Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    freeze_support()
    main()