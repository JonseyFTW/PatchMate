#!/bin/bash
# run-with-ollama.sh - Run the patching bot with proper Ollama configuration

# Replace this with your actual Ollama IP or hostname
OLLAMA_HOST="host.docker.internal"  # This is typically the Docker host IP
OLLAMA_PORT="11434"

echo "Starting container with Ollama at ${OLLAMA_HOST}:${OLLAMA_PORT}..."
docker run -d \
  --name patching-bot \
  -p 5000:5000 \
  -v $(pwd)/watch_dir:/data/shared/patching \
  -v $(pwd)/uploads:/app/uploads \
  -e OLLAMA_URL="http://${OLLAMA_HOST}:${OLLAMA_PORT}/api/generate" \
  -e AI_PROVIDER="ollama" \
  -e OLLAMA_MODEL="mistral" \
  --add-host=host.docker.internal:host-gateway \
  patching-bot:latest

echo "Container started. Follow logs with: docker logs -f patching-bot"
echo "The dashboard is available at: http://localhost:5000"