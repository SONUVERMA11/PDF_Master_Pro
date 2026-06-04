#!/bin/bash
cd /home/kaal/Downloads/PDF_Master_Pro

echo "Setting up virtual environment..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
fi

source venv/bin/activate

echo "Installing dependencies..."
pip install -r requirements.txt

echo "Starting PDF Master Pro..."
streamlit run app.py
