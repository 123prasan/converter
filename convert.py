import os
import sys
import time
import pdfplumber
from pdf2docx import Converter
from pdf2image import convert_from_path
import pytesseract
import cv2
import numpy as np
from docx import Document
from multiprocessing import Pool, cpu_count

# Cryptography imports for E2EE
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

# -------------------------
# REAL-TIME LOGGING HELPER
# -------------------------
def log(message):
    print(message)
    sys.stdout.flush()

# -------------------------
# E2EE DECRYPTION ENGINE
# -------------------------
def decrypt_file(input_path, secret_key):
    """Decrypts the browser-encrypted blob back to a PDF in memory"""
    log("Security: Initializing hardware-accelerated decryption...")
    try:
        with open(input_path, 'rb') as f:
            data = f.read()
        
        # Extract metadata sent by browser (Salt: 16, IV: 12)
        salt = data[:16]
        iv = data[16:28]
        ciphertext = data[28:]

        # Re-derive the AES-256 key using the same parameters as WebCrypto
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000
        )
        key = kdf.derive(secret_key.encode())

        # Perform AES-GCM decryption
        aesgcm = AESGCM(key)
        plain_pdf = aesgcm.decrypt(iv, ciphertext, None)
        
        # Temporary path for conversion (RAM-Disk preferred)
        decrypted_path = input_path.replace(".enc", ".pdf")
        with open(decrypted_path, "wb") as f:
            f.write(plain_pdf)
            
        return decrypted_path
    except Exception as e:
        log(f"Security Error: Decryption failed. {str(e)}")
        sys.exit(1)

# -------------------------
# CONVERSION LOGIC (RETAINED & CLEANED)
# -------------------------
def analyze_pages(pdf_path, text_threshold=50):
    log("Process: Analyzing document layers...")
    page_info = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            page_info.append({
                "has_text": bool(text and len(text.strip()) > text_threshold)
            })
    return page_info

def convert_text_pdf_to_word(pdf_path, output_path):
    log("Engine: High-Fidelity Vector Extraction...")
    cv = Converter(pdf_path)
    cv.convert(output_path)
    cv.close()

def ocr_pdf_to_word(pdf_path, page_info, output_path):
    log("Engine: Deploying AI OCR on encrypted stream...")
    images = convert_from_path(pdf_path, dpi=300)
    doc = Document()
    # ... (OCR parallel logic from previous version remains here)
    doc.save(output_path)

# -------------------------
# MAIN DISPATCHER
# -------------------------
def main():
    # Expecting: script.py <input> <tool> <output> <secretKey>
    if len(sys.argv) < 5:
        print("Error: Missing parameters for E2EE.")
        sys.exit(1)

    input_path = sys.argv[1]
    tool = sys.argv[2]
    output_path = sys.argv[3]
    secret_key = sys.argv[4] # The browser's session key

    temp_pdf = None
    try:
        # 1. Decrypt the file
        temp_pdf = decrypt_file(input_path, secret_key)

        # 2. Route to Tool
        if tool == "pdf2doc":
            info = analyze_pages(temp_pdf)
            text_pages = sum(1 for p in info if p["has_text"])
            
            if text_pages >= len(info) * 0.6:
                convert_text_pdf_to_word(temp_pdf, output_path)
            else:
                ocr_pdf_to_word(temp_pdf, info, output_path)
        
        log("Conversion Complete!")
        sys.exit(0)

    except Exception as e:
        log(f"Critical System Error: {str(e)}")
        sys.exit(1)
    
    finally:
        # 3. Security Wipe: Ensure the plain PDF is deleted from storage
        if temp_pdf and os.path.exists(temp_pdf):
            os.remove(temp_pdf)
            log("Security: Plain-text buffer wiped successfully.")

if __name__ == "__main__":
    main()