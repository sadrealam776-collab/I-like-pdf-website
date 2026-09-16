#!/usr/bin/env bash
# Exit on error
set -o errexit

# Install system-level Tesseract-OCR for Linux
apt-get update
apt-get install -y tesseract-ocr

# Install Python dependencies
pip install -r requirements.txt
