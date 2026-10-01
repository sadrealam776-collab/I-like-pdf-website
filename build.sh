#!/usr/bin/env bash
set -o errexit

# Upgrade pip, setuptools, and wheel
python -m pip install --upgrade pip setuptools wheel

# Install system-level Tesseract-OCR
apt-get update && apt-get install -y tesseract-ocr

# Force pip to use pre-built binary wheels only (prevents C++ compilation failure)
pip install --only-binary=:all: -r requirements.txt
