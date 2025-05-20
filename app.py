# app.py - Main application file with Excel and CSV support

import os
import csv
import time
import json
import re
import sys
import logging
import smtplib
import threading
import requests
import traceback
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from flask import Flask, render_template, jsonify, request, redirect, url_for
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from watchdog.observers.polling import PollingObserver
from datetime import datetime
import pandas as pd
import openai
import openpyxl

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("patching_bot.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Flask application
app = Flask(__name__)
app.config.from_pyfile('config.py')
app.config['TEMPLATES_AUTO_RELOAD'] = True

# Global variables
ACTIVE_JOBS = {}
COMPLETED_JOBS = {}
SERVER_RESULTS = []
MONITOR_ACTIVE = False
observer = None

def initialize_app_state():
    """Initialize or reset application state variables"""
    global ACTIVE_JOBS, COMPLETED_JOBS, SERVER_RESULTS, MONITOR_ACTIVE, observer
    
    # Clear all dictionaries and lists
    ACTIVE_JOBS = {}
    COMPLETED_JOBS = {}
    SERVER_RESULTS = []
    MONITOR_ACTIVE = False
    observer = None
    
    logger.info("Application state initialized")

# Initialize app state when module is loaded
initialize_app_state()


# --- File Monitoring ---

class EnhancedFileHandler(FileSystemEventHandler):
    """Enhanced file handler that supports both CSV and Excel files and multiple event types"""
    
    def process_file(self, file_path):
        """Common logic to process detected files"""
        if not os.path.isdir(file_path):
            file_ext = os.path.splitext(file_path)[1].lower()
            
            # Log verbose information about detected file
            logger.info(f"File detected: {file_path}")
            logger.info(f"File extension: {file_ext}")
            
            # Process CSV and Excel files
            if file_ext in ['.csv', '.xlsx', '.xls']:
                logger.info(f"Processing supported file: {file_path}")
                try:
                    process_machine_file(file_path)
                except Exception as e:
                    logger.error(f"Error processing file {file_path}: {str(e)}")
            else:
                logger.info(f"Ignoring unsupported file type: {file_ext}")
    
    def on_created(self, event):
        """Handle file creation events"""
        logger.debug(f"Creation event detected: {event.src_path}")
        self.process_file(event.src_path)
    
    def on_modified(self, event):
        """Handle file modification events"""
        logger.debug(f"Modification event detected: {event.src_path}")
        # Only process if it's a file (not a directory) and exists
        if not event.is_directory and os.path.exists(event.src_path):
            # Check if this file has been processed before
            # This helps prevent duplicate processing of the same file
            file_path = event.src_path
            if not hasattr(self, '_processed_files'):
                self._processed_files = set()
            
            if file_path not in self._processed_files:
                self._processed_files.add(file_path)
                self.process_file(file_path)
    
    def on_moved(self, event):
        """Handle file moved events (often triggered when files are saved by external programs)"""
        logger.debug(f"Move event detected: {event.dest_path}")
        self.process_file(event.dest_path)

def normalize_path(path):
    """Normalize a file path to prevent issues with semicolons and other special characters"""
    # Remove problematic characters that might create multiple directories
    clean_path = path.replace(';', '').replace('\\', '/').strip()
    
    # Convert Windows paths to Unix format
    if ':' in clean_path:
        # Handle Windows paths like C:\path\to\dir
        parts = clean_path.split(':')
        if len(parts) > 1:
            clean_path = parts[1]  # Take the part after the drive letter
    
    # Ensure no trailing slash
    clean_path = clean_path.rstrip('/')
    
    logger.info(f"Normalized path from '{path}' to '{clean_path}'")
    return clean_path

# Update the ensure_directory_exists function
def ensure_directory_exists(path):
    """Safely create a directory if it doesn't exist"""
    # Normalize the path first
    path = normalize_path(path)
    
    if not os.path.exists(path):
        logger.info(f"Creating directory: {path}")
        try:
            os.makedirs(path, exist_ok=True)
            return True
        except Exception as e:
            logger.error(f"Failed to create directory {path}: {str(e)}")
            return False
    else:
        logger.info(f"Directory already exists: {path}")
        return True
    
def process_machine_file(file_path):
    """Process a file containing machine names (supports CSV and Excel)"""
    logger.info(f"Processing machine file: {file_path}")
    
    try:
        # Determine file type
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # Read the file into a pandas DataFrame based on file type
        if file_ext == '.csv':
            logger.info(f"Reading CSV file: {file_path}")
            try:
                df = pd.read_csv(file_path, encoding='utf-8-sig')  # Handle BOM if present
            except Exception as e:
                logger.warning(f"Error reading with utf-8-sig: {str(e)}, trying other encodings")
                df = pd.read_csv(file_path, encoding='latin1')  # Try another encoding
        elif file_ext in ['.xlsx', '.xls']:
            logger.info(f"Reading Excel file: {file_path}")
            df = pd.read_excel(file_path)
        else:
            logger.error(f"Unsupported file extension: {file_ext}")
            return
        
        # Log column names for debugging
        logger.info(f"Columns found in file: {df.columns.tolist()}")
        
        # Find machine name column using flexible pattern matching
        machine_col = find_machine_name_column(df)
        
        if not machine_col:
            logger.error("Could not find a column containing machine names")
            return
        
        logger.info(f"Using column '{machine_col}' for machine names")
        
        # Extract machine names
        machines = df[machine_col].dropna().tolist()
        logger.info(f"Found {len(machines)} machines in file")
        
        # Log the first few machine names for verification
        if machines:
            logger.info(f"Sample machine names: {machines[:min(5, len(machines))]}")
        
        # Process each machine
        for machine in machines:
            machine = str(machine).strip()  # Convert to string and strip whitespace
            if machine:
                # Start a thread for each machine to process in parallel
                threading.Thread(
                    target=process_machine,
                    args=(machine,),
                    daemon=True
                ).start()
                logger.info(f"Started processing thread for machine: {machine}")
            else:
                logger.warning("Skipping empty machine name")
    
    except Exception as e:
        logger.error(f"Error processing file {file_path}: {str(e)}")
        logger.exception("Detailed traceback:")
        raise

def start_file_monitoring():
    global MONITOR_ACTIVE, observer
    
    if MONITOR_ACTIVE:
        return {"status": "already_running"}
    
    path = normalize_path(app.config['WATCH_DIRECTORY'])
    logger.info(f"Starting file monitoring on {path}")
    
    try:
        # Verify the directory exists
        if not os.path.exists(path):
            logger.info(f"Creating directory: {path}")
            os.makedirs(path, exist_ok=True)
        
        # List contents of the directory for debugging
        try:
            contents = os.listdir(path)
            logger.info(f"Contents of {path}: {contents}")
        except Exception as e:
            logger.warning(f"Could not list directory contents: {str(e)}")
        
        # Use the enhanced file handler with additional logging
        event_handler = EnhancedFileHandler()
        
        # Use PollingObserver for more reliable monitoring with Docker volumes
        from watchdog.observers.polling import PollingObserver
        observer = PollingObserver(timeout=1)  # Poll every 1 second
        
        observer.schedule(event_handler, path, recursive=False)
        observer.start()
        MONITOR_ACTIVE = True
        
        logger.info(f"File monitoring started on {path} (using polling observer)")
        
        # Start a periodic directory check thread for additional reliability
        threading.Thread(
            target=periodic_directory_check,
            args=(path,),
            daemon=True
        ).start()
        
        return {"status": "started"}
    except Exception as e:
        logger.error(f"Failed to start monitoring: {str(e)}")
        logger.exception("Detailed traceback:")
        return {"status": "error", "message": str(e)}

def periodic_directory_check(directory_path):
    """Periodically check the directory for new files as a backup mechanism"""
    last_check_files = set()
    
    while MONITOR_ACTIVE:
        try:
            # Get current files in directory
            current_files = set()
            for filename in os.listdir(directory_path):
                full_path = os.path.join(directory_path, filename)
                if os.path.isfile(full_path):
                    current_files.add(full_path)
            
            # Find new files since last check
            new_files = current_files - last_check_files
            if new_files:
                logger.info(f"Periodic check found {len(new_files)} new files: {new_files}")
                for file_path in new_files:
                    # Process each new file
                    try:
                        file_ext = os.path.splitext(file_path)[1].lower()
                        if file_ext in ['.csv', '.xlsx', '.xls']:
                            logger.info(f"Processing found file: {file_path}")
                            process_machine_file(file_path)
                    except Exception as e:
                        logger.error(f"Error processing found file {file_path}: {str(e)}")
            
            # Update last check files
            last_check_files = current_files
            
        except Exception as e:
            logger.error(f"Error in periodic directory check: {str(e)}")
        
        # Sleep for 30 seconds before next check
        time.sleep(30)
    
def stop_file_monitoring():
    global MONITOR_ACTIVE, observer
    
    if not MONITOR_ACTIVE:
        return {"status": "not_running"}
    
    try:
        observer.stop()
        observer.join()
        MONITOR_ACTIVE = False
        logger.info("File monitoring stopped")
        return {"status": "stopped"}
    except Exception as e:
        logger.error(f"Failed to stop monitoring: {str(e)}")
        return {"status": "error", "message": str(e)}

# --- File Processing ---

def find_machine_name_column(df):
    """Find the column that most likely contains machine names using flexible pattern matching"""
    # Common patterns for machine name columns
    patterns = [
        r'machine.*name',
        r'computer.*name',
        r'host.*name',
        r'server.*name',
        r'machine',
        r'computer',
        r'host',
        r'server',
        r'name'
    ]
    
    # Check for exact matches first (with case insensitivity)
    for col in df.columns:
        col_lower = col.lower()
        if 'machine' in col_lower and 'name' in col_lower:
            return col
    
    # Try regex patterns
    for pattern in patterns:
        for col in df.columns:
            if re.search(pattern, col.lower()):
                return col
    
    # If no matches found and we have a small number of columns, use the first text column
    if len(df.columns) <= 3:
        for col in df.columns:
            # Check if column contains text data that could be server names
            if df[col].dtype == 'object' and not df[col].empty:
                # Check if values look like machine names (alphanumeric with possible dots or hyphens)
                sample = df[col].dropna().iloc[0] if not df[col].dropna().empty else ''
                if isinstance(sample, str) and re.match(r'^[a-zA-Z0-9\.\-]+$', sample):
                    return col
    
    # Fallback: If we have only one column, use it regardless of name
    if len(df.columns) == 1:
        return df.columns[0]
    
    # No suitable column found
    return None

def process_machine(machine_name):
    """Process updates for a specific machine"""
    logger.info(f"Processing machine: {machine_name}")
    job_id = None
    
    try:
        # Run the Azure Automation runbook
        job_id = run_azure_runbook(machine_name)
        
        if not job_id:
            logger.error(f"Failed to start runbook for {machine_name}")
            # Add to completed jobs with error status
            error_id = f"error-{machine_name}-{int(time.time())}"
            COMPLETED_JOBS[error_id] = {
                "machine": machine_name,
                "status": "error",
                "error": "Failed to start runbook",
                "completion_time": datetime.now().isoformat()
            }
            return
        
        # Add job to active jobs
        ACTIVE_JOBS[job_id] = {
            "machine": machine_name,
            "start_time": datetime.now().isoformat(),
            "status": "running"
        }
        
        # Poll job status until complete
        job_output = poll_job_status(job_id)
        
        # Process the results
        if job_output:
            # Parse the job output
            results = parse_job_output(job_output, machine_name)
            
            # Store the results
            COMPLETED_JOBS[job_id] = {
                "machine": machine_name,
                "status": "completed",
                "completion_time": datetime.now().isoformat(),
                "results": results
            }
            
            # Add to server results
            SERVER_RESULTS.append(results)
            
            logger.info(f"Completed processing for {machine_name}: {results['totalUpdates']} updates, {len(results['failedUpdates'])} failed")
        else:
            logger.error(f"Failed to get output for job {job_id} (machine: {machine_name})")
            COMPLETED_JOBS[job_id] = {
                "machine": machine_name,
                "status": "failed",
                "error": "Failed to get job output",
                "completion_time": datetime.now().isoformat()
            }
        
        # Remove from active jobs
        if job_id in ACTIVE_JOBS:
            del ACTIVE_JOBS[job_id]
            
        # Check if all machines are processed
        check_all_complete()
    
    except Exception as e:
        logger.error(f"Error processing machine {machine_name}: {str(e)}")
        logger.exception("Traceback:")
        
        # Ensure we track the failure in completed jobs
        job_key = job_id if job_id else f"error-{machine_name}-{int(time.time())}"
        COMPLETED_JOBS[job_key] = {
            "machine": machine_name,
            "status": "error",
            "error": str(e),
            "completion_time": datetime.now().isoformat()
        }
        
        # Remove from active jobs if it was there
        if job_id and job_id in ACTIVE_JOBS:
            del ACTIVE_JOBS[job_id]
            
        # Still check if all machines are processed
        check_all_complete()

def update_requirements():
    """Install required packages for Excel support if they're not already installed"""
    try:
        import openpyxl
        import xlrd
        logger.info("Excel support libraries already installed")
    except ImportError:
        logger.info("Installing Excel support libraries")
        import subprocess
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "xlrd"])
            logger.info("Successfully installed Excel support libraries")
        except Exception as e:
            logger.error(f"Failed to install Excel support libraries: {str(e)}")
            logger.error("You may need to manually install them: pip install openpyxl xlrd")

# --- Azure Integration ---

def process_machine(machine_name):
    """Process updates for a specific machine"""
    logger.info(f"Processing machine: {machine_name}")
    
    try:
        # Run the Azure Automation runbook
        job_id = run_azure_runbook(machine_name)
        
        if not job_id:
            logger.error(f"Failed to start runbook for {machine_name}")
            return
        
        # Add job to active jobs
        ACTIVE_JOBS[job_id] = {
            "machine": machine_name,
            "start_time": datetime.now().isoformat(),
            "status": "running"
        }
        
        # Poll job status until complete
        job_output = poll_job_status(job_id)
        
        if not job_output:
            logger.error(f"Failed to get output for job {job_id} (machine: {machine_name})")
            COMPLETED_JOBS[job_id] = {
                "machine": machine_name,
                "status": "failed",
                "error": "Failed to get job output"
            }
            return
        
        # Parse the job output
        results = parse_job_output(job_output, machine_name)
        
        # Store the results
        COMPLETED_JOBS[job_id] = {
            "machine": machine_name,
            "status": "completed",
            "completion_time": datetime.now().isoformat(),
            "results": results
        }
        
        # Add to server results
        SERVER_RESULTS.append(results)
        
        # Remove from active jobs
        if job_id in ACTIVE_JOBS:
            del ACTIVE_JOBS[job_id]
        
        # Check if all machines are processed
        check_all_complete()
    
    except Exception as e:
        logger.error(f"Error processing machine {machine_name}: {str(e)}")
        if job_id:
            COMPLETED_JOBS[job_id] = {
                "machine": machine_name,
                "status": "error",
                "error": str(e)
            }

def run_azure_runbook(machine_name):
    """Run the Azure Automation runbook and return the job ID"""
    logger.info(f"Running Azure runbook for {machine_name}")
    
    try:
        webhook_url = app.config['AZURE_WEBHOOK_URL']
        
        payload = {
            "ComputerName": machine_name
        }
        
        response = requests.post(
            webhook_url,
            json=payload,
            headers={"Content-Type": "application/json"}
        )
        
        if response.status_code == 202:
            # Extract the job ID from the response
            result = response.json()
            job_id = result.get('JobIds', [None])[0]
            logger.info(f"Runbook started for {machine_name}, Job ID: {job_id}")
            return job_id
        else:
            logger.error(f"Failed to start runbook: {response.status_code} - {response.text}")
            return None
    
    except Exception as e:
        logger.error(f"Error running Azure runbook: {str(e)}")
        return None

# --- Update these two functions in app.py ---

def poll_job_status(job_id):
    """Poll the job status until complete and return the job output"""
    logger.info(f"Polling job status for job ID: {job_id}")
    
    subscription_id = app.config['AZURE_SUBSCRIPTION_ID']
    resource_group = app.config['AZURE_RESOURCE_GROUP']
    automation_account = app.config['AZURE_AUTOMATION_ACCOUNT']
    
    status_url = f"https://management.azure.com/subscriptions/{subscription_id}/resourceGroups/{resource_group}/providers/Microsoft.Automation/automationAccounts/{automation_account}/jobs/{job_id}?api-version=2019-06-01"
    output_url = f"https://management.azure.com/subscriptions/{subscription_id}/resourceGroups/{resource_group}/providers/Microsoft.Automation/automationAccounts/{automation_account}/jobs/{job_id}/output?api-version=2019-06-01"
    
    # Get the access token
    access_token = get_azure_token()
    
    if not access_token:
        logger.error("Failed to get Azure access token")
        return None
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    # Poll status until job is complete
    max_attempts = 30
    attempt = 0
    
    while attempt < max_attempts:
        try:
            response = requests.get(status_url, headers=headers)
            
            if response.status_code == 200:
                job_status = response.json()
                status = job_status.get('properties', {}).get('status')
                
                if status == 'Completed':
                    logger.info(f"Job {job_id} completed")
                    break
                elif status in ['Failed', 'Suspended', 'Stopped']:
                    logger.error(f"Job {job_id} ended with status: {status}")
                    return None
            else:
                logger.error(f"Failed to get job status: {response.status_code} - {response.text}")
            
            # Wait before next attempt
            time.sleep(15)
            attempt += 1
        
        except Exception as e:
            logger.error(f"Error polling job status: {str(e)}")
            attempt += 1
    
    if attempt >= max_attempts:
        logger.error(f"Max polling attempts reached for job {job_id}")
        return None
    
    # Get job output - IMPORTANT: don't try to parse as JSON, return raw text
    try:
        response = requests.get(output_url, headers=headers)
        
        if response.status_code == 200:
            # We can see from logs that we're getting plain text, not JSON
            # So we'll get the raw text response rather than trying to parse JSON
            raw_output = response.text
            
            # Log the first part of the output for debugging
            logger.debug(f"Raw output from job (first 200 chars): {raw_output[:200]}")
            
            # Check if output contains PowerShell data
            if "ComputerName" in raw_output and "Date" in raw_output:
                # This appears to be valid PowerShell output
                return raw_output
            else:
                # If we don't see the expected content, try JSON parsing as fallback
                try:
                    json_output = response.json()
                    return json_output.get('value', '')
                except:
                    # Return whatever we got if JSON parsing fails
                    return raw_output
        else:
            logger.error(f"Failed to get job output: {response.status_code} - {response.text}")
            return None
    
    except Exception as e:
        logger.error(f"Error getting job output: {str(e)}")
        return None

def get_azure_token():
    """Get an Azure access token"""
    client_id = app.config['AZURE_CLIENT_ID']
    client_secret = app.config['AZURE_CLIENT_SECRET']
    tenant_id = app.config['AZURE_TENANT_ID']
    
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/token"
    
    payload = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "resource": "https://management.azure.com/"
    }
    
    try:
        response = requests.post(token_url, data=payload)
        
        if response.status_code == 200:
            return response.json().get('access_token')
        else:
            logger.error(f"Failed to get token: {response.status_code} - {response.text}")
            return None
    
    except Exception as e:
        logger.error(f"Error getting Azure token: {str(e)}")
        return None

# --- Output Parsing ---

def parse_job_output(output, server_name):
    """Parse the job output and extract update information with detailed logging"""
    logger.info(f"Parsing job output for {server_name}")
    
    # Default structure for results
    results = {
        "serverName": server_name,
        "allUpdates": [],
        "failedUpdates": [],
        "totalUpdates": 0,
        "hasFailures": False
    }
    
    try:
        if not output:
            logger.warning(f"No output to parse for {server_name}")
            return results
        
        # Log a sample of the output for debugging
        output_length = len(output)
        logger.info(f"Output length: {output_length} characters")
        
        sample_size = min(1000, output_length)
        sample = output[:sample_size]
        logger.info(f"Output sample (first {sample_size} chars):\n{sample}...")
        
        # Get server name from output if present
        if "Received ComputerName:" in output:
            server_match = re.search(r"Received ComputerName:\s*([^\n]+)", output)
            if server_match and server_match.group(1).strip():
                results["serverName"] = server_match.group(1).strip()
                logger.info(f"Found server name in output: {results['serverName']}")

        # Look for failure indicators directly in the raw output
        failure_terms = ["Failed", "Error", "Failure", "failed", "error", "failure"]
        raw_failure_indicators = [term for term in failure_terms if term in output]
        
        if raw_failure_indicators:
            logger.info(f"Found potential failure indicators in raw output: {raw_failure_indicators}")
            
            # Extract some context around failures for debugging
            for term in raw_failure_indicators:
                matches = re.finditer(term, output)
                for match in matches:
                    pos = match.start()
                    start = max(0, pos - 100)
                    end = min(output_length, pos + 100)
                    context = output[start:end].replace('\n', ' ')
                    logger.info(f"Failure context: '...{context}...'")

        # Parse updates using a simpler approach - split by date entries
        updates = []
        
        # Split the output into blocks based on Date headers
        blocks = re.split(r"\nDate\s+:", output)
        
        logger.info(f"Found {len(blocks)} update blocks in output")
        
        # Process each block (skip the first one which is the header)
        for i, block in enumerate(blocks[1:] if len(blocks) > 1 else []):
            # Reinsert the "Date:" prefix that got removed by the split
            block_text = "Date    : " + block.strip()
            
            try:
                # Extract fields using simple pattern matching
                date_match = re.search(r"Date\s+:\s+([^\n]+)", block_text)
                operation_match = re.search(r"Operation\s+:\s+([^\n]+)", block_text)
                status_match = re.search(r"Status\s+:\s+([^\n]+)", block_text)
                update_match = re.search(r"Update\s+:\s+([^\n]+)", block_text)
                title_match = re.search(r"Title\s+:\s+([^\n]+)", block_text)
                
                if date_match and status_match:  # Require at least date and status
                    update_info = {
                        "date": date_match.group(1).strip(),
                        "operation": operation_match.group(1).strip() if operation_match else "",
                        "status": status_match.group(1).strip(),
                        "update": update_match.group(1).strip() if update_match else "",
                        "title": title_match.group(1).strip() if title_match else ""
                    }
                    
                    # Log each parsed update
                    logger.info(f"Parsed update block {i+1}: {json.dumps(update_info)}")
                    
                    # Detect failures and log them prominently
                    if "Failed" in update_info.get("status", ""):
                        logger.warning(f"FAILURE DETECTED in update: {json.dumps(update_info)}")
                    
                    updates.append(update_info)
                else:
                    logger.warning(f"Skipping block {i+1} - missing required fields")
                    logger.debug(f"Incomplete block content: {block_text[:200]}...")
            except Exception as e:
                logger.error(f"Error parsing update block {i+1}: {str(e)}")
                logger.debug(f"Problematic block content: {block_text[:200]}...")
                continue
        
        # Store all updates
        results["allUpdates"] = updates
        results["totalUpdates"] = len(updates)
        
        # Find failed updates
        failed_updates = [u for u in updates if "Failed" in u.get("status", "")]
        results["failedUpdates"] = failed_updates
        results["hasFailures"] = len(failed_updates) > 0
        
        # Log the final results summary
        logger.info(f"Parsed {len(updates)} updates for {server_name}, {len(failed_updates)} failed")
        if failed_updates:
            logger.warning(f"Server {server_name} has {len(failed_updates)} failed updates")
            for i, failed in enumerate(failed_updates):
                logger.warning(f"  Failed update {i+1}: {json.dumps(failed)}")
        
        return results
    
    except Exception as e:
        logger.error(f"Error parsing job output: {str(e)}")
        logger.exception("Detailed traceback:")
        return results  # Return default results

# --- Results Processing ---

def check_all_complete():
    """Check if all machines are processed and generate a report"""
    if ACTIVE_JOBS:
        # Still have active jobs
        return
    
    if not SERVER_RESULTS:
        # No results to process
        return
    
    logger.info("All machines processed, generating report")
    
    try:
        # Aggregate results
        aggregate_results = {
            "serverResults": SERVER_RESULTS,
            "totalServers": len(SERVER_RESULTS),
            "serversWithFailures": len([s for s in SERVER_RESULTS if s.get("hasFailures", False)]),
            "totalFailedUpdates": sum(len(s.get("failedUpdates", [])) for s in SERVER_RESULTS)
        }
        
        # Start a separate thread for report generation and email sending
        # This prevents timeouts from blocking the main application
        threading.Thread(
            target=generate_and_send_report,
            args=(aggregate_results,),
            daemon=True
        ).start()
        
        logger.info("Report generation started in background thread")
    
    except Exception as e:
        logger.error(f"Error in check_all_complete: {str(e)}")

def generate_and_send_report(aggregate_results):
    """Generate a report and send email in a separate thread"""
    try:
        # Generate report using AI
        logger.info("Generating AI report in background thread")
        report = generate_ai_report(aggregate_results)
        
        # Send email notification
        logger.info("Sending email notification with generated report")
        send_email_notification(report)
        
        # Clear results for next run
        SERVER_RESULTS.clear()
        
        # Log the completion
        logger.info("Report generated and email sent")
        
    except Exception as e:
        logger.error(f"Error in generate_and_send_report: {str(e)}")
        logger.exception("Detailed traceback:")
        
        # Try to send a fallback email if AI report generation failed
        try:
            fallback_report = generate_template_report(aggregate_results)
            fallback_report = "ERROR: AI report generation failed. Using template report instead.\n\n" + fallback_report
            send_email_notification(fallback_report)
            logger.info("Fallback report sent due to error in AI report generation")
        except Exception as email_error:
            logger.error(f"Failed to send fallback email: {str(email_error)}")

# --- AI Integration ---

def generate_ai_report(results):
    """Generate a report using AI with better fallback handling and detailed logging"""
    logger.info("Generating AI report")
    
    try:
        # Prepare the prompt
        prompt = "You are a system administrator assistant analyzing Windows update information.\n\n"
        prompt += "Review the following update data from all servers:\n\n"
        
        # Check if we have valid server results
        if not results or not results.get("serverResults"):
            logger.warning("No server results to generate report")
            return generate_template_report(results)
        
        # Count servers with failures for validation
        servers_with_failures = 0
        total_failed_updates = 0
        
        for server in results.get("serverResults", []):
            has_failures = server.get("hasFailures", False)
            if has_failures:
                servers_with_failures += 1
                total_failed_updates += len(server.get("failedUpdates", []))
            
            has_failures_str = "Yes" if has_failures else "No"
            prompt += f"## Server: {server.get('serverName', 'Unknown')}\n"
            prompt += f"Failed Updates: {has_failures_str}\n\n"
            
            if has_failures:
                prompt += "The following updates failed:\n\n"
                for update in server.get("failedUpdates", []):
                    prompt += f"- Date: {update.get('date', 'Unknown')}\n"
                    prompt += f"  - Update: {update.get('update', 'Unknown')}\n"
                    prompt += f"  - Title: {update.get('title', 'Unknown')}\n\n"
            else:
                prompt += "No updates failed on this server.\n\n"
        
        # Add explicit summary statistics to make it clearer to the AI
        prompt += "\n## Summary Statistics:\n"
        prompt += f"- Total Servers: {len(results.get('serverResults', []))}\n"
        prompt += f"- Servers With Failures: {servers_with_failures}\n"
        prompt += f"- Total Failed Updates: {total_failed_updates}\n\n"
        
        prompt += "\nPlease provide a concise summary of all Windows update results, highlighting any failed updates across all servers. "
        prompt += "Include the total number of servers scanned, how many had failed updates, and provide recommendations on what actions should be taken next."
        prompt += "\nIMPORTANT: If any servers had failures, clearly identify them and never state that no failures were detected."
        
        # Log the prompt length for debugging
        logger.info(f"AI Prompt (length: {len(prompt)}):\n{prompt}")
        
        # Choose AI integration method based on configuration
        ai_provider = app.config.get('AI_PROVIDER', 'template').lower()
        logger.info(f"Using AI provider: {ai_provider}")
        
        if ai_provider == 'openai':
            return generate_openai_report(prompt)
        elif ai_provider == 'ollama':
            # THIS IS THE IMPORTANT CHANGE - Check if your generate_ollama_report takes 1 or 2 parameters
            # Option 1: If your generate_ollama_report takes only 1 parameter
            return generate_ollama_report(prompt)
            
            # Option 2: If you've updated generate_ollama_report to take 2 parameters
            # return generate_ollama_report(prompt, servers_with_failures > 0)
        else:
            # Fallback to a simple template if no AI
            logger.info(f"Using template report (AI provider '{ai_provider}' not configured)")
            return generate_template_report(results)
    
    except Exception as e:
        logger.error(f"Error generating AI report: {str(e)}")
        logger.exception("Detailed traceback:")
        return generate_template_report(results)

def generate_ollama_report(prompt):
    """Generate a report using Ollama with enhanced logging and validation"""
    try:
        # Check if the prompt indicates there are failures
        has_failures_flag = "Servers With Failures: 0" not in prompt and "Failed Updates: Yes" in prompt
        
        ollama_url = app.config.get('OLLAMA_URL', 'http://host.docker.internal:11434/api/generate')
        ollama_model = app.config.get('OLLAMA_MODEL', 'mistral')
        
        logger.info(f"Connecting to Ollama at: {ollama_url}")
        logger.info(f"Using model: {ollama_model}")
        
        # Log request parameters
        request_body = {
            "model": ollama_model,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": 0.2,  # Lower temperature for more consistent responses
                "top_p": 0.9
            }
        }
        logger.info(f"Ollama request parameters: model={ollama_model}, prompt_length={len(prompt)}")
        
        # Add longer timeout for large prompts
        start_time = time.time()
        logger.info(f"Sending request to Ollama using model '{ollama_model}'...")
        
        response = requests.post(
            ollama_url,
            json=request_body,
            timeout=1200 # 2 minute timeout
        )
        
        request_time = time.time() - start_time
        logger.info(f"Ollama request completed in {request_time:.2f} seconds")
        
        # Log the response status
        logger.info(f"Ollama response status code: {response.status_code}")
        
        if response.status_code == 200:
            # Parse the JSON response
            try:
                json_response = response.json()
                result = json_response.get('response', '')
                logger.info(f"Successfully generated report with model '{ollama_model}' (length: {len(result)})")
                
                # Validate the response content if we know there are failures
                if has_failures_flag:
                    # Check if the response correctly mentions failures
                    failure_terms = ["fail", "failed", "failure", "unsuccessful", "error", "issues"]
                    found_failure_mention = any(term in result.lower() for term in failure_terms)
                    
                    if not found_failure_mention:
                        logger.warning(f"VALIDATION FAILED: Model '{ollama_model}' didn't mention failures when they exist!")
                        logger.warning("Adding a correction to the report...")
                        
                        # Add a correction notice
                        result = (
                            f"⚠️ CORRECTION: The system detected failed updates, but the AI summary "
                            f"from model '{ollama_model}' did not properly acknowledge them. Please check the detailed logs. ⚠️\n\n"
                        ) + result
                
                return result
            except Exception as e:
                logger.error(f"Error parsing Ollama JSON response: {str(e)}")
                return generate_template_report({"serverResults": []})
        else:
            logger.error(f"Ollama error with model '{ollama_model}': {response.status_code}")
            
            # Handle common errors specifically
            if response.status_code == 404:
                error_message = f"Model '{ollama_model}' not found on Ollama server"
                logger.error(error_message)
                return f"ERROR: {error_message}\n\n" + generate_template_report({"serverResults": []})
            elif response.status_code == 500:
                error_message = f"Internal server error from Ollama with model '{ollama_model}'"
                logger.error(error_message)
                return f"ERROR: {error_message}\n\n" + generate_template_report({"serverResults": []})
            
            return generate_template_report({"serverResults": []})
    
    except Exception as e:
        logger.error(f"Error with Ollama: {str(e)}")
        logger.exception("Detailed traceback:")
        # Fall back to template report
        return f"ERROR connecting to Ollama: {str(e)}\n\n" + generate_template_report({"serverResults": []})

def generate_openai_report(prompt):
    """Generate a report using OpenAI"""
    try:
        openai.api_key = app.config['OPENAI_API_KEY']
        
        response = openai.chat.completions.create(
            model=app.config.get('OPENAI_MODEL', 'gpt-3.5-turbo'),
            messages=[
                {"role": "system", "content": "You are a system administrator expert analyzing Windows updates."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=1000
        )
        
        return response.choices[0].message.content
    
    except Exception as e:
        logger.error(f"Error with OpenAI: {str(e)}")
        return f"Error generating AI report: {str(e)}"

def generate_ollama_report(prompt):
    """Generate a report using Ollama with enhanced logging and validation"""
    try:
        # Check if the prompt indicates there are failures
        has_failures_flag = "Servers With Failures: 0" not in prompt and "Failed Updates: Yes" in prompt
        
        ollama_url = app.config.get('OLLAMA_URL', 'http://host.docker.internal:11434/api/generate')
        ollama_model = app.config.get('OLLAMA_MODEL', 'mistral')
        
        logger.info(f"Connecting to Ollama at: {ollama_url}")
        logger.info(f"Using model: {ollama_model}")
        
        # Log request parameters
        request_body = {
            "model": ollama_model,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": 0.2,  # Lower temperature for more consistent responses
                "top_p": 0.9
            }
        }
        logger.info(f"Ollama request parameters: model={ollama_model}, prompt_length={len(prompt)}")
        
        # Add longer timeout for large prompts
        start_time = time.time()
        logger.info(f"Sending request to Ollama using model '{ollama_model}'...")
        
        response = requests.post(
            ollama_url,
            json=request_body,
            timeout=1200  # 2 minute timeout
        )
        
        request_time = time.time() - start_time
        logger.info(f"Ollama request completed in {request_time:.2f} seconds")
        
        # Log the response status
        logger.info(f"Ollama response status code: {response.status_code}")
        
        if response.status_code == 200:
            # Parse the JSON response
            try:
                json_response = response.json()
                result = json_response.get('response', '')
                logger.info(f"Successfully generated report with model '{ollama_model}' (length: {len(result)})")
                
                # Validate the response content if we know there are failures
                if has_failures_flag:
                    # Check if the response correctly mentions failures
                    failure_terms = ["fail", "failed", "failure", "unsuccessful", "error", "issues"]
                    found_failure_mention = any(term in result.lower() for term in failure_terms)
                    
                    if not found_failure_mention:
                        logger.warning(f"VALIDATION FAILED: Model '{ollama_model}' didn't mention failures when they exist!")
                        logger.warning("Adding a correction to the report...")
                        
                        # Add a correction notice
                        result = (
                            f"⚠️ CORRECTION: The system detected failed updates, but the AI summary "
                            f"from model '{ollama_model}' did not properly acknowledge them. Please check the detailed logs. ⚠️\n\n"
                        ) + result
                
                return result
            except Exception as e:
                logger.error(f"Error parsing Ollama JSON response: {str(e)}")
                return generate_template_report({"serverResults": []})
        else:
            logger.error(f"Ollama error with model '{ollama_model}': {response.status_code}")
            
            # Handle common errors specifically
            if response.status_code == 404:
                error_message = f"Model '{ollama_model}' not found on Ollama server"
                logger.error(error_message)
                return f"ERROR: {error_message}\n\n" + generate_template_report({"serverResults": []})
            elif response.status_code == 500:
                error_message = f"Internal server error from Ollama with model '{ollama_model}'"
                logger.error(error_message)
                return f"ERROR: {error_message}\n\n" + generate_template_report({"serverResults": []})
            
            return generate_template_report({"serverResults": []})
    
    except Exception as e:
        logger.error(f"Error with Ollama: {str(e)}")
        logger.exception("Detailed traceback:")
        # Fall back to template report
        return f"ERROR connecting to Ollama: {str(e)}\n\n" + generate_template_report({"serverResults": []})

def generate_template_report(results):
    """Generate a basic report without AI"""
    total_servers = results.get("totalServers", 0)
    servers_with_failures = results.get("serversWithFailures", 0)
    total_failed_updates = results.get("totalFailedUpdates", 0)
    
    report = f"Windows Update Report\n\n"
    report += f"Total Servers Scanned: {total_servers}\n"
    report += f"Servers With Failed Updates: {servers_with_failures}\n"
    report += f"Total Failed Updates: {total_failed_updates}\n\n"
    
    if servers_with_failures > 0:
        report += "Servers with Failed Updates:\n\n"
        
        for server in results.get("serverResults", []):
            if server.get("hasFailures", False):
                report += f"Server: {server.get('serverName', 'Unknown')}\n"
                report += f"Failed Updates: {len(server.get('failedUpdates', []))}\n\n"
                
                for update in server.get("failedUpdates", []):
                    report += f"- Date: {update.get('date', 'Unknown')}\n"
                    report += f"  Update: {update.get('update', 'Unknown')}\n"
                    report += f"  Title: {update.get('title', 'Unknown')}\n\n"
    else:
        report += "No failures detected across any servers.\n"
    
    report += "\nRecommendation: "
    if servers_with_failures > 0:
        report += "Review the failed updates and schedule remediation."
    else:
        report += "No action required. All updates were successfully applied."
    
    return report

# --- Email Notification ---

def send_email_notification(report):
    """Send an email notification with the update report"""
    logger.info("Sending email notification")
    
    try:
        from_email = app.config['EMAIL_FROM']
        to_email = app.config['EMAIL_TO']
        smtp_server = app.config['SMTP_SERVER']
        smtp_port = app.config['SMTP_PORT']
        smtp_username = app.config.get('SMTP_USERNAME')
        smtp_password = app.config.get('SMTP_PASSWORD')
        
        # Create the email
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = 'Failed Updates Report - All Servers'
        
        # Add report to email
        msg.attach(MIMEText(report, 'plain'))
        
        # Send the email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            # No STARTTLS or authentication needed for internal mail server
            server.send_message(msg)
        
        logger.info("Email notification sent")
    
    except Exception as e:
        logger.error(f"Error sending email: {str(e)}")

# --- Web Interface Routes ---

@app.route('/')
def index():
    """Render the main dashboard page"""
    return render_template(
        'index.html',
        monitor_active=MONITOR_ACTIVE,
        active_jobs=ACTIVE_JOBS,
        completed_jobs=COMPLETED_JOBS,
        server_results=SERVER_RESULTS,
        config=app.config
    )

@app.route('/api/start-monitoring', methods=['POST'])
def api_start_monitoring():
    result = start_file_monitoring()
    return jsonify(result)

@app.route('/api/stop-monitoring', methods=['POST'])
def api_stop_monitoring():
    result = stop_file_monitoring()
    return jsonify(result)

@app.route('/api/process-csv', methods=['POST'])
def api_process_csv():
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "No file part"})
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({"status": "error", "message": "No selected file"})
    
    # Get file extension
    file_ext = os.path.splitext(file.filename)[1].lower()
    
    # Check if file type is supported
    if file_ext not in ['.csv', '.xlsx', '.xls']:
        return jsonify({
            "status": "error", 
            "message": f"Unsupported file type: {file_ext}. Please upload CSV or Excel files."
        })
    
    try:
        upload_dir = app.config['UPLOAD_DIRECTORY']
        # Safely ensure upload directory exists
        if not ensure_directory_exists(upload_dir):
            return jsonify({"status": "error", "message": f"Failed to create upload directory: {upload_dir}"})
            
        file_path = os.path.join(upload_dir, file.filename)
        file.save(file_path)
        
        logger.info(f"File saved to {file_path}, processing...")
        process_machine_file(file_path)
        
        return jsonify({"status": "success", "message": f"File processed: {file.filename}"})
    
    except Exception as e:
        logger.error(f"Error processing uploaded file: {str(e)}")
        logger.exception("Detailed traceback:")
        return jsonify({"status": "error", "message": str(e)})

@app.route('/api/status')
def api_status():
    """Enhanced API status endpoint that provides detailed job information"""
    # Ensure global dictionaries are initialized
    global ACTIVE_JOBS, COMPLETED_JOBS
    
    if ACTIVE_JOBS is None:
        ACTIVE_JOBS = {}
    if COMPLETED_JOBS is None:
        COMPLETED_JOBS = {}
    
    # Format active jobs data for the UI
    active_jobs_data = []
    for job_id, job_info in ACTIVE_JOBS.items():
        active_jobs_data.append({
            "id": job_id,
            "machine": job_info.get("machine", "Unknown"),
            "start_time": job_info.get("start_time", ""),
            "status": job_info.get("status", "running")
        })
    
    # Format completed jobs data for the UI
    completed_jobs_data = []
    for job_id, job_info in COMPLETED_JOBS.items():
        job_data = {
            "id": job_id,
            "machine": job_info.get("machine", "Unknown"),
            "completion_time": job_info.get("completion_time", ""),
            "status": job_info.get("status", "unknown"),
            "error": job_info.get("error", "") if job_info.get("status") in ["error", "failed"] else ""
        }
        
        # Add update counts if available
        if job_info.get("status") == "completed" and "results" in job_info:
            results = job_info.get("results", {})
            job_data["updates"] = results.get("totalUpdates", 0)
            job_data["failed"] = len(results.get("failedUpdates", []))
            job_data["results"] = results
        else:
            job_data["updates"] = 0
            job_data["failed"] = 0
            
        completed_jobs_data.append(job_data)
    
    # Log what we're returning for debugging
    logger.debug(f"Status API returning: {len(active_jobs_data)} active jobs, {len(completed_jobs_data)} completed jobs")
    
    return jsonify({
        "monitor_active": MONITOR_ACTIVE,
        "active_jobs": len(active_jobs_data),
        "active_jobs_data": active_jobs_data,
        "completed_jobs": len(completed_jobs_data),
        "completed_jobs_data": completed_jobs_data,
        "server_results": len(SERVER_RESULTS) if SERVER_RESULTS else 0
    })

# Add a reset endpoint to clear data if needed
@app.route('/api/reset', methods=['POST'])
def api_reset():
    """Reset application state"""
    initialize_app_state()
    return jsonify({"status": "success", "message": "Application state reset"})

# Add these routes to app.py

@app.route('/api/ollama-models')
def api_ollama_models():
    """Get available models from Ollama server"""
    try:
        # Get base URL from config (without the /api/generate endpoint)
        ollama_url = app.config.get('OLLAMA_URL', 'http://host.docker.internal:11434/api/generate')
        base_url = ollama_url.rsplit('/api/', 1)[0]

        # Query Ollama API for available models
        models_url = f"{base_url}/api/tags"
        logger.info(f"Querying Ollama models at: {models_url}")
        
        response = requests.get(models_url, timeout=10)
        
        if response.status_code == 200:
            models_data = response.json()
            logger.info(f"Found {len(models_data.get('models', []))} models in Ollama")
            
            # Format the results
            models = []
            for model in models_data.get('models', []):
                models.append({
                    'name': model.get('name'),
                    'modified_at': model.get('modified_at'),
                    'size': model.get('size')
                })
                
            return jsonify({
                "status": "success",
                "current_model": app.config.get('OLLAMA_MODEL', 'mistral'),
                "models": models
            })
        else:
            logger.error(f"Error querying Ollama models: {response.status_code} - {response.text}")
            return jsonify({
                "status": "error", 
                "message": f"Failed to get models: {response.status_code}",
                "current_model": app.config.get('OLLAMA_MODEL', 'mistral')
            })
            
    except Exception as e:
        logger.error(f"Error getting Ollama models: {str(e)}")
        return jsonify({
            "status": "error",
            "message": str(e),
            "current_model": app.config.get('OLLAMA_MODEL', 'mistral')
        })

@app.route('/api/set-ollama-model', methods=['POST'])
def api_set_ollama_model():
    """Set the Ollama model to use for report generation"""
    try:
        data = request.json
        model_name = data.get('model')
        
        if not model_name:
            return jsonify({"status": "error", "message": "No model name provided"})
        
        # Update the config
        app.config['OLLAMA_MODEL'] = model_name
        logger.info(f"Changed Ollama model to: {model_name}")
        
        return jsonify({
            "status": "success",
            "message": f"Model changed to {model_name}",
            "current_model": model_name
        })
        
    except Exception as e:
        logger.error(f"Error setting Ollama model: {str(e)}")
        return jsonify({
            "status": "error",
            "message": str(e)
        })
    
@app.route('/api/clear-completed', methods=['POST'])
def api_clear_completed():
    COMPLETED_JOBS.clear()
    return jsonify({"status": "success"})

@app.route('/api/debug', methods=['GET'])
def api_debug():
    """Debug endpoint to see internal state and test AI reporting"""
    try:
        # Gather debug info
        debug_info = {
            "app_version": "1.1.0",
            "monitor_active": MONITOR_ACTIVE,
            "active_jobs_count": len(ACTIVE_JOBS),
            "completed_jobs_count": len(COMPLETED_JOBS),
            "server_results_count": len(SERVER_RESULTS),
            "active_jobs": ACTIVE_JOBS,
            "completed_jobs": COMPLETED_JOBS,
            "server_results": SERVER_RESULTS
        }
        
        # Check if user wants to test AI report generation
        test_ai = request.args.get('test_ai', 'false').lower() == 'true'
        
        if test_ai and SERVER_RESULTS:
            # Generate a test report
            aggregate_results = {
                "serverResults": SERVER_RESULTS,
                "totalServers": len(SERVER_RESULTS),
                "serversWithFailures": len([s for s in SERVER_RESULTS if s.get("hasFailures", False)]),
                "totalFailedUpdates": sum(len(s.get("failedUpdates", [])) for s in SERVER_RESULTS)
            }
            
            report = generate_ai_report(aggregate_results)
            debug_info["test_ai_report"] = report
            
        return jsonify(debug_info)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})
    
@app.route('/api/reload-template')
def reload_template():
    """Force reload template and restart application"""
    try:
        # Clear template cache
        app.jinja_env.cache = {}
        return jsonify({"status": "success", "message": "Template cache cleared, refresh the page"})
    except Exception as e:
        logger.error(f"Error clearing template cache: {str(e)}")
        return jsonify({"status": "error", "message": str(e)})
    
@app.route('/api/system-info')
def api_system_info():
    try:
        info = {
            "app_version": "1.1.0",  # Update this with your version
            "python_version": sys.version,
            "os_info": os.name,
            "watch_directory": app.config['WATCH_DIRECTORY'],
            "watch_directory_exists": os.path.exists(app.config['WATCH_DIRECTORY']),
            "upload_directory": app.config['UPLOAD_DIRECTORY'],
            "upload_directory_exists": os.path.exists(app.config['UPLOAD_DIRECTORY']),
            "monitor_active": MONITOR_ACTIVE,
            "active_jobs": len(ACTIVE_JOBS),
            "completed_jobs": len(COMPLETED_JOBS),
            "results_count": len(SERVER_RESULTS),
            "installed_packages": {
                "pandas": pd.__version__,
                "flask": Flask.__version__,
                "watchdog": "Installed"  # You could import watchdog to get the version
            }
        }
        
        # Check for Excel support
        try:
            import openpyxl
            info["excel_support"] = f"Installed (openpyxl {openpyxl.__version__})"
        except ImportError:
            info["excel_support"] = "Not installed"
        
        return jsonify(info)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

# --- Main Entry Point ---

if __name__ == '__main__':
    # Initialize the app state to ensure clean startup
    initialize_app_state()
    
    # Try to ensure Excel support is available
    try:
        update_requirements()
    except Exception as e:
        logger.warning(f"Unable to automatically install Excel support: {str(e)}")
    
    # Start file monitoring if configured to do so
    if app.config.get('AUTO_START_MONITORING', False):
        start_file_monitoring()
    
    # Start the Flask application
    app.run(
        host=app.config.get('HOST', '0.0.0.0'),
        port=app.config.get('PORT', 5000),
        debug=app.config.get('DEBUG', False)
    )