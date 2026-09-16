#!/usr/bin/env bash
set -o errexit

# Upgrade pip, setuptools, and wheel
python -m pip install --upgrade pip setuptools wheel

# Install system dependencies
apt-get update && apt-get install -y tesseract-ocr

# Install Python requirements with binary preference
pip install --prefer-binary -r requirements.txt
