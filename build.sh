#!/usr/bin/env bash
set -o errexit

# Upgrade core packaging tools
python -m pip install --upgrade pip setuptools wheel

# Install Linux system packages for OCR
apt-get update && apt-get install -y tesseract-ocr

# Install packages using pre-built binary wheels only
pip install --only-binary=:all: -r requirements.txt
