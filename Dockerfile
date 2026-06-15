# Use official lightweight Python image
FROM python:3.10-slim

# Install system dependencies (LibreOffice for Word to PDF, and libraries for PyMuPDF/OpenCV)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables
ENV PYTHONUNBUFFERED=1
ENV PORT=5000

# Set working directory
WORKDIR /app

# Copy requirements and install dependencies
COPY requirements.txt .
# Filter out Windows-only pywin32 library
RUN grep -v "pywin32" requirements.txt > requirements-linux.txt
RUN pip install --no-cache-dir -r requirements-linux.txt
RUN pip install --no-cache-dir gunicorn

# Copy project files
COPY . .

# Create uploads directory
RUN mkdir -p uploads

# Expose port
EXPOSE 5000

# Run the application with gunicorn using the PORT env variable
CMD gunicorn --bind 0.0.0.0:$PORT app:app
