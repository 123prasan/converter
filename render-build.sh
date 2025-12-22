#!/usr/bin/env bash
apt-get update
# Install LibreOffice for Word/PPT/Excel conversion
apt-get install -y libreoffice
# Install Poppler for PDF to Image conversion
apt-get install -y poppler-utils
# Install Tesseract for OCR
apt-get install -y tesseract-ocr
# Install libGL for OpenCV (Required by EasyOCR)
apt-get install -y libgl1