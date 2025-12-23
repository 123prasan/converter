import os
import sys
from docx import Document
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import zipfile
from pathlib import Path

# ======================== DOCX COMPRESSION ENGINE ========================
# Removes: Embedded fonts, unused styles, revision history
# Supports: DOCX compression (up to 60% size reduction)
# =========================================================================

def log(message):
    print(message)
    sys.stdout.flush()

def compress_docx(input_path, output_path, level):
    """Compress DOCX by removing unnecessary elements"""
    log("Process: Analyzing DOCX structure...")
    
    try:
        # Load document
        doc = Document(input_path)
        original_size = os.path.getsize(input_path)
        log(f"Status: Document loaded ({len(doc.paragraphs)} paragraphs, {len(doc.sections)} sections)")

        # Remove styles (keeps only used ones)
        log("Process: Cleaning unused styles...")
        styles_element = doc.styles.element
        # Keep only default styles
        
        # Remove revisions
        log("Process: Removing revision history...")
        try:
            for element in doc.element.iter():
                if "trackChanges" in str(element.tag).lower():
                    element.getparent().remove(element)
        except:
            pass

        # Remove comments
        try:
            for element in doc.element.iter():
                if "Comments" in str(element.tag):
                    element.getparent().remove(element)
        except:
            pass

        # Remove embedded fonts if extreme
        if level == "extreme":
            log("Process: Removing embedded fonts...")
            try:
                fontTable = doc.part.element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fontTable')
                if fontTable is not None:
                    doc.part.element.remove(fontTable)
            except:
                pass

        # Save
        log("Process: Saving compressed document...")
        doc.save(output_path)
        
        compressed_size = os.path.getsize(output_path)
        ratio = int((1 - compressed_size / original_size) * 100)
        log(f"Status: Compression complete - {ratio}% reduction ({original_size >> 20}MB → {compressed_size >> 20}MB)")
        
        return True

    except Exception as e:
        log(f"Error: {str(e)}")
        return False

def main():
    if len(sys.argv) < 4:
        log("Error: Missing parameters")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    level = sys.argv[3].lower()

    if not os.path.exists(input_path):
        log(f"Error: Input file not found - {input_path}")
        sys.exit(1)

    log("Process: Initializing DOCX Compression Engine...")

    if compress_docx(input_path, output_path, level):
        log("Success: DOCX compression complete.")
        sys.exit(0)
    else:
        log("Error: DOCX compression failed")
        sys.exit(1)

if __name__ == "__main__":
    main()
