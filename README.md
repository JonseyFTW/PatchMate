# PatchMate

PatchMate automates Windows patch status collection using Azure Automation runbooks. Server results are aggregated and an email summary is generated.

## Features
- Watches a directory for CSV or Excel files containing server names.
- Launches Azure Automation runbooks to gather patch status and diagnostics.
- Searches the web for articles referencing failed KB updates and includes helpful links in the report (prioritising Microsoft and Reddit sources). Searches are performed using Playwright for more reliable results.
- Sends a professional looking report via email in both plain text and HTML formats.

## Configuration
Edit `config.py` to update Azure, email and AI provider settings. Install Python dependencies from `requirements.txt` and run `app.py`.
