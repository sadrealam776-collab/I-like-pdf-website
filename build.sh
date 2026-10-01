#!/usr/bin/env bash
set -o errexit

# Upgrade core packaging tools
python -m pip install --upgrade pip setuptools wheel

# Install Python dependencies normally
pip install -r requirements.txt
