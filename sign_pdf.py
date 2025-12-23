import sys
import os

# Create a local debug log to catch errors that stderr might miss
def debug_log(text):
    with open("python_debug.log", "a") as f:
        f.write(text + "\n")

try:
    import fitz  # PyMuPDF
except ImportError:
    debug_log("CRITICAL: PyMuPDF (fitz) is not installed. Run 'pip install pymupdf'")
    print("Error: PyMuPDF not installed")
    sys.exit(7)

def log(message):
    print(message)
    sys.stdout.flush()

def main():
    try:
        if len(sys.argv) < 11:
            sys.exit(1)

        input_pdf = os.path.abspath(sys.argv[1])
        output_pdf = os.path.abspath(sys.argv[2])
        name = sys.argv[3]
        page_range = sys.argv[4]
        margin_x = float(sys.argv[5])
        margin_y = float(sys.argv[6])
        font_size = float(sys.argv[7])
        color_hex = sys.argv[8].lstrip('#')
        font_style = sys.argv[9]
        opacity = float(sys.argv[10])

        # Convert hex color to RGB (0-1 range)
        rgb = tuple(int(color_hex[i:i+2], 16) / 255 for i in (0, 2, 4))
        
        # Map font styles to PyMuPDF fonts
        font_map = {
            "script": "tibi",      # Times Bold Italic for elegant script look
            "modern": "tibo",      # Times Bold
            "classic": "tiit",     # Times Italic
            "formal": "tiro"       # Times Regular
        }
        selected_font = font_map.get(font_style, "tibi")

        if not os.path.exists(input_pdf):
            sys.exit(1)

        doc = fitz.open(input_pdf)
        total_pages = len(doc)
        
        target_indices = []
        if page_range.lower() == "all":
            target_indices = list(range(total_pages))
        elif "-" in page_range:
            parts = page_range.split("-")
            target_indices = list(range(int(parts[0])-1, min(int(parts[1]), total_pages)))
        else:
            target_indices = [int(p.strip()) - 1 for p in page_range.split(",") if p.strip().isdigit()]

        for p_idx in target_indices:
            if 0 <= p_idx < total_pages:
                page = doc[p_idx]
                page_rect = page.rect
                
                # Calculate text dimensions
                text_len = fitz.get_text_length(name, fontname=selected_font, fontsize=font_size)
                
                # Position calculation (convert from bottom-left to PyMuPDF coordinates)
                safe_x = max(10, min(margin_x, page_rect.width - text_len - 10))
                safe_y = max(font_size + 10, min(page_rect.height - margin_y, page_rect.height - 10))

                # Insert signature with opacity
                page.insert_text(fitz.Point(safe_x, safe_y), name, 
                                fontsize=font_size, 
                                color=rgb, 
                                fontname=selected_font, 
                                overlay=True)

        # Save with optimization
        doc.save(output_pdf, garbage=4, deflate=True)
        doc.close()
        
        log("Success: PDF Signed with Style")
        sys.exit(0)

    except Exception as e:
        debug_log(f"Execution Error: {str(e)}")
        log(f"Error: {str(e)}")
        sys.exit(7)

if __name__ == "__main__":
    main()