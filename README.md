# ShotDORO Video Analyzer

A video analysis tool that detects shooting events and adds timing markers to CSV files using OCR-based ammunition counting.

## Quick Start

### Prerequisites
- Python 3.7+
- Tesseract OCR: Download from [tesseract-ocr/tesseract](https://github.com/tesseract-ocr/tesseract)

### Installation
```bash
pip install opencv-python pytesseract pillow pandas pyinstaller
```

### Usage

1. **🎬 Load Video** - Select your video file
2. **🖼️ Select Frame** - Choose from 10 random sample frames for region setup
3. **🎯 Set Regions** - Select ammunition counter areas (total and current ammo)
4. **🚀 Start Analysis** - Begin automatic shooting detection
5. **📋 Load CSV** - Select your existing CSV file for ping addition
6. **➕ Add Shot Pings** - Merge detected shooting times with CSV data
7. **Done!** - Export CSV with shooting markers

### Building EXE
```bash
pyinstaller --onefile --windowed --name="ShotDORO" main.py
```

### Features
- OCR-based ammunition detection
- Random frame sampling for optimal region selection
- Advanced region selection with zoom/pan controls
- Automatic CSV time column detection
- Shooting event timing with customizable tolerance

### Output
The program adds a "shot_time" column to your CSV with "O" markers at detected shooting moments.

---
*Minimum region selection: 10x10 pixels*