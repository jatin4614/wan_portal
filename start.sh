#!/bin/bash
# Quick start script for the Vajr Comn State portal
cd "$(dirname "$0")"

if [ ! -d "venv" ]; then
  echo "Creating virtual environment..."
  python3 -m venv venv
fi

source venv/bin/activate
pip install -q -r requirements.txt

# Upload cap: set to 0 to disable, or any number of MB.
# Defaults to 16 GB if unset.
export MAX_UPLOAD_MB="${MAX_UPLOAD_MB:-16384}"

echo "Starting Vajr Comn State on http://0.0.0.0:5000 (upload cap: $MAX_UPLOAD_MB MB)"
python app.py
