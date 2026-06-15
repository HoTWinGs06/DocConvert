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

---

## How to Upload to your GitHub

Initialize your remote repository on GitHub and push this code:

1. Go to [github.com/new](https://github.com/new) and create a repository named `DocConvert`.
2. Connect and push your local commits by running these commands in your project folder:
   ```bash
   git remote add origin https://github.com/HoTWinGs06/DocConvert.git
   git push -u origin main
   ```

---

## How to Deploy for Free (Render.com)

Render offers a fully-featured **Free Tier** for Docker-based Web Services.

1. Create a free account at [Render.com](https://render.com).
2. From the Dashboard, click **New +** and select **Web Service**.
3. Connect your GitHub account and select your `DocConvert` repository.
4. Render will automatically read the `Dockerfile` in the project. Configure the settings:
   - **Name**: `docconvert`
   - **Region**: Choose the closest one to you
   - **Branch**: `main`
   - **Runtime**: `Docker` (automatically selected)
   - **Instance Type**: **Free** ($0/month)
5. Click **Deploy Web Service**. 

Render will automatically pull the code, install LibreOffice, resolve dependencies, and launch your private converter web application!
