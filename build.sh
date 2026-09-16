#!/usr/bin/env bash
set -o errexit

# Upgrade pip, setuptools, and wheel to support modern pre-built binary wheels
python -m pip install --upgrade pip setuptools wheel

# Install system-level Tesseract-OCR for Linux
apt-get update && apt-get install -y tesseract-ocr

# Install Python dependencies
pip install -r requirements.txt
