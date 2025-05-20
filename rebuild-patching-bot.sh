#!/bin/bash
# rebuild-patching-bot.sh - Script to rebuild and restart the patching bot
OLLAMA_HOST="host.docker.internal"  
OLLAMA_PORT="11434"

echo "Stopping any running containers..."
docker stop patching-bot 2>/dev/null || true
docker rm patching-bot 2>/dev/null || true

echo "Building new image with Excel support..."
docker build -f Dockerfile -t patching-bot:latest .

echo "Creating required host directories..."
mkdir -p watch_dir
mkdir -p uploads

echo "Cleaning up any incorrect directories..."
# Remove any directories with semicolons in the name
find . -name "*\;*" -type d -exec rm -rf {} \; 2>/dev/null || true

# Convert Windows-style paths to Docker-compatible paths
# This helps prevent the semicolon issue on Windows
WATCH_DIR_PATH=$(cd watch_dir && pwd)
UPLOADS_PATH=$(cd uploads && pwd)

echo "Using paths:"
echo " - Watch directory: $WATCH_DIR_PATH"
echo " - Uploads directory: $UPLOADS_PATH"

echo "Starting new container..."
docker run -d \
  --name patching-bot \
  -p 5000:5000 \
  -v "$WATCH_DIR_PATH:/data/shared/patching" \
  -v "$UPLOADS_PATH:/app/uploads" \
  -e OLLAMA_URL="http://${OLLAMA_HOST}:${OLLAMA_PORT}/api/generate" \
  -e AI_PROVIDER="ollama" \
  -e OLLAMA_MODEL="mistral" \
  --add-host=host.docker.internal:host-gateway \
  patching-bot:latest

echo "Container started. Follow logs with: docker logs -f patching-bot"
echo "The dashboard is available at: http://localhost:5000"