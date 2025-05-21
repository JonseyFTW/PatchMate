# config.py - Configuration settings for the application

import os

# Application settings
DEBUG = True
HOST = '0.0.0.0'
PORT = 5000
SECRET_KEY = 'q9K2TKsJLZF9RQasmoi4baZraoxS2Bj'

# File monitoring settings
AUTO_START_MONITORING = False
WATCH_DIRECTORY = '/data/shared/patching'
UPLOAD_DIRECTORY = 'uploads'

# Azure settings
AZURE_SUBSCRIPTION_ID = '164esdf-377c-4fdf-859d-61bsdff07df'
AZURE_RESOURCE_GROUP = 'Azure-Arc'
AZURE_AUTOMATION_ACCOUNT = 'Azure-Arc-Automation'
AZURE_CLIENT_ID = 'be908e77-ab14-4c7f-9960-dd853364d58c'
AZURE_CLIENT_SECRET = 'cwo8Q~d9Xi~_dylnDEsdfXAEdJ.c-F'
AZURE_TENANT_ID = '1a68cf95-dsdfcf728efdf'
AZURE_WEBHOOK_URL = 'https://cdcdsfc9-ee03-452b-a5c5-392f1720ad.webhook.eus.azure-automation.net/webhooks?token=q9K2TKsJLZ%2fF9RQasmoisdfBjCNO8nICVwG0%3d'

# AI settings
AI_PROVIDER = os.getenv('AI_PROVIDER', 'vllm')  # 'openai' or 'ollama' or 'vllm'
OPENAI_API_KEY = os.getenv('OPENAI_API_KEY', 'your-openai-api-key')
OPENAI_MODEL = os.getenv('OPENAI_MODEL', 'gpt-3.5-turbo')

OLLAMA_URL = os.getenv('OLLAMA_URL', 'http://host.docker.internal:11434/api/generate')
OLLAMA_MODEL = os.getenv('OLLAMA_MODEL', 'mistral')

VLLM_CHAT_COMPLETIONS_URL = os.getenv('VLLM_CHAT_COMPLETIONS_URL', 'https://vllm.contosoresources.com/v1/chat/completions')
VLLM_MODELS_URL = os.getenv('VLLM_MODELS_URL', 'https://vllm.contosoresources.com/v1/models')
VLLM_MODEL = os.getenv('VLLM_MODEL', 'Qwen3-30B-A3B-FP8') # Specify your default VLLM model
VLLM_API_KEY = os.getenv('VLLM_API_KEY', '123456789') # Add if your VLLM setup requires an API key
VLLM_VERIFY_SSL = False


# Email settings
EMAIL_FROM = 'patching@contosoresources.com'
EMAIL_TO = 'cjones@contosoresources.com'
SMTP_SERVER = 'webmail.contosoresources.com'
SMTP_PORT = 25
SMTP_USERNAME = None
SMTP_PASSWORD = None