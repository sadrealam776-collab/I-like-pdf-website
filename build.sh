#!/usr/bin/env bash
set -o errexit

# Upgrade core packaging tools
python -m pip install --upgrade pip setuptools wheel

# Install dependencies using pre-built binary wheels only (never build from source)
pip install --only-binary=:all: -r requirements.txt
