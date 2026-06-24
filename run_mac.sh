#!/bin/bash
set -e
cd "$(dirname "$0")"
python3 -m venv venv
source venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install --no-cache-dir -r requirements.txt
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501
