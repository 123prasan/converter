import os
import sys
import zipfile
from pdf2image import convert_from_path

def log(message):
    print(message)
    sys.stdout.flush()

def main():
    # Args: script.py <input_path> <output_folder> <dpi>
    if len(sys.argv) < 4:
        sys.exit(1)

    input_path = sys.argv[1]
    output_dir = sys.argv[2]
    dpi = int(sys.argv[3])
    
    log(f"Process: Initializing Rendering Engine at {dpi} DPI...")

    try:
        # 1. Convert PDF pages to list of PIL images
        # thread_count=4 speeds up rendering on multi-core CPUs
        pages = convert_from_path(input_path, dpi=dpi, thread_count=4)
        
        image_paths = []
        log(f"Process: Rasterizing {len(pages)} pages...")

        for i, page in enumerate(pages):
            img_name = f"page_{i+1}.jpg"
            img_path = os.path.join(output_dir, img_name)
            
            # Save as high-quality JPG
            page.save(img_path, 'JPEG', quality=95)
            image_paths.append(img_path)
            log(f"Status: Rendered page {i+1}/{len(pages)}")

        # 2. If multi-page, wrap in a ZIP for the user
        zip_name = f"converted_images_{i+1}.zip"
        zip_path = os.path.join(output_dir, zip_name)
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in image_paths:
                zipf.write(file, os.path.basename(file))
                os.remove(file) # Clean up individual JPEGs

        log(f"Success: ZIP archive created.")
        # We print the final filename so Node.js knows what to serve
        print(f"RESULT_FILE:{zip_name}")
        sys.exit(0)

    except Exception as e:
        log(f"Critical Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()