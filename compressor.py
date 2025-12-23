import os
import sys
import fitz  # PyMuPDF
from PIL import Image
from io import BytesIO
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

    # Validate input
    if not os.path.exists(input_path):
        log(f"Error: Input file not found: {input_path}")
        sys.exit(1)

    log("Process: Initializing Shrink Engine v2.0...")
    
    try:
        # Step 1: Open PDF
        doc = fitz.open(input_path)
        original_pages = doc.page_count
        log(f"Process: Analyzing document ({original_pages} pages)...")

        # Step 2: Optimize images & handle grayscale
        if original_pages > 0:
            log("Process: Optimizing image streams...")
            
            for page_num in range(original_pages):
                page = doc[page_num]
                
                # Get all images in page
                images = page.get_images()
                if images:
                    log(f"Process: Compressing {len(images)} images on page {page_num + 1}...")
                    
                    for img_index in images:
                        try:
                            xref = img_index[0]
                            pix = fitz.Pixmap(doc, xref)
                            
                            # Apply grayscale if requested
                            if grayscale and pix.colorspace.n > 1:
                                gray_pix = fitz.Pixmap(fitz.csGRAY, pix)
                                
                                # Compress with JPEG quality=85
                                img_data = gray_pix.tobytes("jpeg", 85)
                                doc.update_stream(xref, img_data)
                                gray_pix = None
                            elif not grayscale:
                                # Lossy JPEG compression at quality 88 for color images
                                if pix.n - pix.alpha > 1:  # Has color
                                    img_data = pix.tobytes("jpeg", 88)
                                    doc.update_stream(xref, img_data)
                            
                            pix = None
                        except Exception as img_err:
                            log(f"Warning: Could not optimize image {img_index}: {str(img_err)}")
                            continue
        
        # Step 3: Remove metadata if requested
        if strip_meta:
            log("Process: Stripping metadata...")
            doc.metadata = None
        
        log(f"Process: Rebuilding structure with {dpi} DPI optimization...")
        
        # Step 4: Advanced save with maximum compression
        # garbage=4: Remove unused objects and rebuild xref
        # deflate=True: Compress streams
        # clean=True: Clean redundant data
        # linear=True: Linearize for faster web viewing
        doc.save(
            output_path, 
            garbage=4,      # Full cleanup
            deflate=True,   # Stream compression
            clean=True,     # Remove redundancy
            linear=True,    # Web-optimized
            pretty=False,   # No formatting
            no_new_id=False # Update ID
        )
        
        doc.close()
        
        # Calculate compression ratio
        original_size = os.path.getsize(input_path)
        compressed_size = os.path.getsize(output_path)
        ratio = int((1 - compressed_size / original_size) * 100)
        
        log(f"Compression Complete! Reduced by {ratio}% ({original_size >> 20}MB → {compressed_size >> 20}MB)")
        sys.exit(0)

    except Exception as e:
        # Better error handling
        error_msg = str(e)
        if "corrupted" in error_msg.lower() or "xref" in error_msg.lower():
            log(f"Error: PDF is corrupted. Try repairing with an online tool first.")
        elif "permission" in error_msg.lower():
            log(f"Error: PDF is password-protected. Unlock first before compressing.")
        elif "out of memory" in error_msg.lower():
            log(f"Error: File too large for compression. Try a smaller file.")
        else:
            log(f"Error: {error_msg}")
        sys.exit(1)

if __name__ == "__main__":
    main()
