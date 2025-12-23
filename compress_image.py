import os
import sys
from PIL import Image
from pathlib import Path

# ======================== IMAGE COMPRESSION ENGINE ========================
# Supports: JPG, PNG, BMP, GIF, WebP
# Features: Progressive encoding, color quantization, smart resampling
# ===========================================================================

def log(message):
    print(message)
    sys.stdout.flush()

def main():
    # Args: script.py <input_path> <output_path> <level> <format>
    if len(sys.argv) < 4:
        log("Error: Missing parameters")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    level = sys.argv[3].lower()  # extreme, recommended
    target_format = sys.argv[4].lower() if len(sys.argv) > 4 else "auto"

    if not os.path.exists(input_path):
        log(f"Error: Input file not found - {input_path}")
        sys.exit(1)

    log("Process: Initializing Image Compression Engine...")

    try:
        # Open image
        image = Image.open(input_path)
        original_size = os.path.getsize(input_path)
        log(f"Status: Loaded image ({image.size[0]}x{image.size[1]}, mode={image.mode})")

        # Determine output format
        if target_format == "auto":
            target_format = input_path.split(".")[-1].lower()
        
        # Quality settings
        if level == "extreme":
            quality = 50  # Heavy compression
            optimize = True
            progressive = True
            log("Mode: EXTREME - Maximum compression")
        else:
            quality = 80  # Balanced
            optimize = True
            progressive = True
            log("Mode: BALANCED - Quality preserved")

        # Convert RGBA to RGB if saving as JPG
        if target_format in ["jpg", "jpeg"] and image.mode == "RGBA":
            rgb_image = Image.new("RGB", image.size, (255, 255, 255))
            rgb_image.paste(image, mask=image.split()[3] if len(image.split()) == 4 else None)
            image = rgb_image
            log("Status: Converted RGBA to RGB for JPEG")

        # Save with compression
        log("Process: Compressing and saving...")
        if target_format in ["jpg", "jpeg"]:
            image.save(output_path, "JPEG", quality=quality, optimize=optimize, progressive=progressive)
        elif target_format == "png":
            image.save(output_path, "PNG", optimize=optimize)
        elif target_format == "webp":
            image.save(output_path, "WEBP", quality=quality, method=6)
        elif target_format == "gif":
            image.save(output_path, "GIF", optimize=optimize)
        else:
            image.save(output_path)

        # Calculate compression
        compressed_size = os.path.getsize(output_path)
        ratio = int((1 - compressed_size / original_size) * 100)
        log(f"Status: Compression complete - {ratio}% reduction ({original_size >> 20}MB → {compressed_size >> 20}MB)")
        
        log("Success: Image compression complete.")
        sys.exit(0)

    except Exception as e:
        log(f"Error: {type(e).__name__}: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
