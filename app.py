import sys
import os
from convert import is_text_pdf, convert_text_pdf, ocr_pdf_advanced

def main():
    if len(sys.argv) < 3:
        print("Usage: python convert.py <pdf_path> <tool>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    tool = sys.argv[2].lower()

    # Set output path based on tool
    output_name = "output_text.docx" if tool != "ocr" else "output_ocr_advanced.docx"
    output_path = os.path.join(os.getcwd(), output_name)

    os.makedirs("outputs", exist_ok=True)

    try:
        if tool == "ocr":
            ocr_pdf_advanced(pdf_path, output_path)
        else:
            convert_text_pdf(pdf_path, output_path)

        print(output_path)  # Node can capture this
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
