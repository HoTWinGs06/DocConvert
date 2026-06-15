# DocConvert — PDF ↔ Word Converter

A private, minimalist, distraction-free document converter web application. 

## Features
- **Cross-Platform Architecture**: Uses native Microsoft Word COM engine on Windows for 100% layout fidelity, and automatically falls back to LibreOffice headless + `pdf2docx` on Linux environments.
- **Privacy First**: Files are uploaded with unique UUID names and are permanently deleted after 1 hour by a background daemon.
- **Sleek UI**: Single-page design with Inter typography, drag-and-drop support, upload progress tracking, and automatic file extension detection.
- **Cloud Ready**: Configured for 1-click free deployment using Docker.

---

## Local Development (Windows)

1. Make sure Python 3.10+ and Microsoft Word are installed.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the development server:
   ```bash
   python app.py
   ```
4. Open your browser and go to `http://127.0.0.1:5000`.
