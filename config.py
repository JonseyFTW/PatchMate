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
AZURE_SUBSCRIPTION_ID = '164e94dfgd-61b0487f07df'
AZURE_RESOURCE_GROUP = 'Azure-Arc'
AZURE_AUTOMATION_ACCOUNT = 'Azure-Arc-Automation'
AZURE_CLIENT_ID = 'be908dfgdd853364d58c'
AZURE_CLIENT_SECRET = 'cwg9Xi~_dylndfgjtbXAEdJ.c-F'
AZURE_TENANT_ID = '1a68gf-8f7b-44ecfgefdf'
AZURE_WEBHOOK_URL = 'https://cdcagdf5c1720ad.webhook.eus.azure-automation.net/webhooks?token=qgK2dsJLZ%2fF9RQasmoifgBjCNO8nICVwG0%3d'

# AI settings
AI_PROVIDER = os.getenv('AI_PROVIDER', 'ollama')  # 'openai' or 'ollama'
OPENAI_API_KEY = os.getenv('OPENAI_API_KEY', 'your-openai-api-key')
OPENAI_MODEL = os.getenv('OPENAI_MODEL', 'gpt-3.5-turbo')
OLLAMA_URL = os.getenv('OLLAMA_URL', 'http://host.docker.internal:11434/api/generate')
OLLAMA_MODEL = os.getenv('OLLAMA_MODEL', 'mistral')
# Email settings
EMAIL_FROM = 'patching@contoso.com'
EMAIL_TO = 'cjones@contoso.com'
SMTP_SERVER = 'webmail.contoso.com'
SMTP_PORT = 25
SMTP_USERNAME = None
SMTP_PASSWORD = None