#!/bin/bash
# Word Document Translator - Mac/Linux Launcher
# Run this script: chmod +x run.sh && ./run.sh

echo "========================================"
echo "  Word Document Translator"
echo "  English to Simplified Chinese"
echo "========================================"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo "Please install Python 3.8+ from https://www.python.org/downloads/"
    exit 1
fi

# Check if venv exists, create if not
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "ERROR: Failed to create virtual environment."
        exit 1
    fi
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Install/update requirements
echo "Checking dependencies..."
pip install -q -r requirements.txt
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install dependencies."
    exit 1
fi

echo ""
echo "Starting the application..."
echo "The browser will open automatically."
echo "Press Ctrl+C to stop the server."
echo ""

# Run Streamlit
streamlit run app.py

