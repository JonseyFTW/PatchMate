FROM python:3.8-slim

WORKDIR /app

# Install system dependencies including those needed for Playwright
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    gcc \
    python3-dev \
    wget \
    gnupg \
    ca-certificates \
    libnss3 \
    libnspr4 \
    libdbus-1-3 \
    libatk1.0-0 \
    libdrm2 \
    libxkbcommon0 \
    libgtk-3-0 \
    libatspi2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Install Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright browsers (this is the missing step!)
RUN playwright install chromium
RUN playwright install-deps

# Copy application files
COPY app.py config.py ./
COPY templates ./templates/

# Explicitly create required directories
RUN mkdir -p /app/uploads /data/shared/patching

# Set environment variables
ENV PYTHONUNBUFFERED=1

# Run the application
CMD ["python", "app.py"]