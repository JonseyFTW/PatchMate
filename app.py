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
from werkzeug.utils import secure_filename # For secure file uploads
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from watchdog.observers.polling import PollingObserver
from datetime import datetime
import pandas as pd
import openai
# import openpyxl # Imported dynamically in update_requirements

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
# Using a simple dictionary to hold app state to allow for easier reset and management
APP_STATE = {
    "active_jobs": {},
    "completed_jobs": {},
    "server_results": [],
    "monitor_active": False,
    "observer": None,
    "processed_files_on_startup": set() # For FileSystemEventHandler
}

def initialize_app_state():
    """Initialize or reset application state variables"""
    global APP_STATE
    APP_STATE["active_jobs"] = {}
    APP_STATE["completed_jobs"] = {}
    APP_STATE["server_results"] = []
    APP_STATE["monitor_active"] = False
    if APP_STATE["observer"]:
        try:
            APP_STATE["observer"].stop()
            APP_STATE["observer"].join()
        except Exception as e:
            logger.warning(f"Could not stop existing observer during reset: {e}")
        APP_STATE["observer"] = None
    APP_STATE["processed_files_on_startup"] = set() # Reset this as well
    logger.info("Application state initialized/reset")

# Initialize app state when module is loaded
initialize_app_state()


# --- File Monitoring ---

class EnhancedFileHandler(FileSystemEventHandler):
    """Enhanced file handler that supports both CSV and Excel files and multiple event types"""
    
    def __init__(self):
        super().__init__()
        self._processed_this_session = APP_STATE["processed_files_on_startup"]

    def process_file(self, file_path):
        """Common logic to process detected files"""
        if not os.path.isdir(file_path):
            file_ext = os.path.splitext(file_path)[1].lower()
            
            logger.info(f"File detected: {file_path}")
            logger.info(f"File extension: {file_ext}")
            
            if file_ext in ['.csv', '.xlsx', '.xls']:
                if file_path in self._processed_this_session:
                    logger.info(f"File {file_path} already processed in this session. Skipping.")
                    return

                logger.info(f"Processing supported file: {file_path}")
                try:
                    process_machine_file(file_path)
                    self._processed_this_session.add(file_path) 
                except Exception as e:
                    logger.error(f"Error processing file {file_path}: {str(e)}")
            else:
                logger.info(f"Ignoring unsupported file type: {file_ext}")
    
    def on_created(self, event):
        logger.debug(f"Creation event detected: {event.src_path}")
        self.process_file(event.src_path)
    
    def on_modified(self, event):
        logger.debug(f"Modification event detected: {event.src_path}")
        if not event.is_directory and os.path.exists(event.src_path):
            self.process_file(event.src_path) 
    
    def on_moved(self, event):
        logger.debug(f"Move event detected: {event.dest_path}")
        self.process_file(event.dest_path)

def normalize_path_conservative(path_str):
    if not isinstance(path_str, str):
        logger.warning(f"normalize_path_conservative received non-string input: {type(path_str)}. Returning as is.")
        return path_str
    normalized = path_str.replace('\\', '/')
    normalized = normalized.strip()
    if len(normalized) > 1 and normalized.endswith('/'):
        normalized = normalized[:-1]
    logger.debug(f"Normalized path from '{path_str}' to '{normalized}' (conservative)")
    return normalized


def ensure_directory_exists(path):
    path_to_check = path 
    if not os.path.exists(path_to_check):
        logger.info(f"Creating directory: {path_to_check}")
        try:
            os.makedirs(path_to_check, exist_ok=True)
            return True
        except Exception as e:
            logger.error(f"Failed to create directory {path_to_check}: {str(e)}")
            return False
    else:
        logger.info(f"Directory already exists: {path_to_check}")
        return True
    
def process_machine_file(file_path):
    logger.info(f"Processing machine file: {file_path}")
    
    try:
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.csv':
            logger.info(f"Reading CSV file: {file_path}")
            try:
                df = pd.read_csv(file_path, encoding='utf-8-sig')
            except UnicodeDecodeError:
                logger.warning(f"UTF-8-SIG decoding failed for {file_path}, trying latin1.")
                df = pd.read_csv(file_path, encoding='latin1')
            except Exception as e:
                logger.error(f"Error reading CSV {file_path}: {str(e)}")
                raise 
        elif file_ext in ['.xlsx', '.xls']:
            logger.info(f"Reading Excel file: {file_path}")
            try:
                df = pd.read_excel(file_path)
            except Exception as e:
                logger.error(f"Error reading Excel file {file_path}: {str(e)}")
                raise 
        else:
            logger.error(f"Unsupported file extension: {file_ext}")
            return
        
        logger.info(f"Columns found in file: {df.columns.tolist()}")
        machine_col = find_machine_name_column(df)
        
        if not machine_col:
            logger.error("Could not find a column containing machine names in the file.")
            return
        
        logger.info(f"Using column '{machine_col}' for machine names")
        machines = df[machine_col].dropna().astype(str).tolist() 
        logger.info(f"Found {len(machines)} machines in file")
        
        if machines:
            logger.info(f"Sample machine names: {machines[:min(5, len(machines))]}")
        
        for machine_raw in machines:
            machine = str(machine_raw).strip() 
            if machine: 
                threading.Thread(
                    target=process_machine,
                    args=(machine,),
                    daemon=True
                ).start()
                logger.info(f"Started processing thread for machine: {machine}")
            else:
                logger.warning("Skipping empty machine name found in file.")
    
    except Exception as e:
        logger.error(f"Critical error in process_machine_file for {file_path}: {str(e)}")
        logger.exception("Detailed traceback for process_machine_file:")

def start_file_monitoring():
    if APP_STATE["monitor_active"]:
        return {"status": "already_running"}
    
    watch_dir = app.config['WATCH_DIRECTORY']
    logger.info(f"Attempting to start file monitoring on: {watch_dir}")
    
    try:
        if not os.path.exists(watch_dir):
            logger.info(f"Watch directory {watch_dir} does not exist. Attempting to create.")
            if not ensure_directory_exists(watch_dir): 
                 logger.error(f"Failed to create watch directory {watch_dir}. Monitoring cannot start.")
                 return {"status": "error", "message": f"Failed to create watch directory: {watch_dir}"}

        try:
            contents = os.listdir(watch_dir)
            logger.info(f"Contents of {watch_dir}: {contents}")
        except Exception as e:
            logger.warning(f"Could not list directory contents for {watch_dir}: {str(e)}")
        
        event_handler = EnhancedFileHandler()
        APP_STATE["observer"] = PollingObserver(timeout=5) 
        APP_STATE["observer"].schedule(event_handler, watch_dir, recursive=False)
        APP_STATE["observer"].start()
        APP_STATE["monitor_active"] = True
        
        logger.info(f"File monitoring started on {watch_dir} (using polling observer)")
        threading.Thread(target=periodic_directory_check, args=(watch_dir,), daemon=True).start()
        return {"status": "started"}
    except Exception as e:
        logger.error(f"Failed to start monitoring on {watch_dir}: {str(e)}")
        logger.exception("Detailed traceback for start_file_monitoring:")
        return {"status": "error", "message": str(e)}

def periodic_directory_check(directory_path):
    last_check_files = set()
    try:
        if os.path.exists(directory_path):
            for filename in os.listdir(directory_path):
                full_path = os.path.join(directory_path, filename)
                if os.path.isfile(full_path):
                    last_check_files.add(full_path)
            logger.info(f"Periodic check initialized with {len(last_check_files)} files in {directory_path}")
    except Exception as e:
        logger.error(f"Error during initial population in periodic_directory_check for {directory_path}: {str(e)}")

    while APP_STATE["monitor_active"]:
        try:
            current_files = set()
            if not os.path.exists(directory_path):
                logger.warning(f"Periodic check: Directory {directory_path} not found. Skipping check.")
                time.sleep(30) 
                continue

            for filename in os.listdir(directory_path):
                full_path = os.path.join(directory_path, filename)
                if os.path.isfile(full_path):
                    current_files.add(full_path)
            
            new_files = current_files - last_check_files
            if new_files:
                logger.info(f"Periodic check found {len(new_files)} new files: {new_files}")
                for file_path in new_files:
                    if file_path in APP_STATE["processed_files_on_startup"]:
                        logger.info(f"Periodic check: File {file_path} already handled by Watchdog. Skipping.")
                        continue
                    try:
                        file_ext = os.path.splitext(file_path)[1].lower()
                        if file_ext in ['.csv', '.xlsx', '.xls']:
                            logger.info(f"Periodic check processing found file: {file_path}")
                            process_machine_file(file_path)
                            APP_STATE["processed_files_on_startup"].add(file_path) 
                    except Exception as e:
                        logger.error(f"Error processing found file {file_path} in periodic check: {str(e)}")
            last_check_files = current_files
        except Exception as e:
            logger.error(f"Error in periodic directory check loop: {str(e)}")
        time.sleep(30) 
    logger.info("Periodic directory check thread stopped.")
    
def stop_file_monitoring():
    if not APP_STATE["monitor_active"] or not APP_STATE["observer"]:
        return {"status": "not_running"}
    try:
        APP_STATE["observer"].stop()
        APP_STATE["observer"].join() 
        APP_STATE["monitor_active"] = False
        APP_STATE["observer"] = None 
        logger.info("File monitoring stopped")
        return {"status": "stopped"}
    except Exception as e:
        logger.error(f"Failed to stop monitoring: {str(e)}")
        return {"status": "error", "message": str(e)}

# --- File Processing ---

def find_machine_name_column(df):
    preferred_patterns = [
        r'^machine\s*name$', r'^computer\s*name$', r'^host\s*name$', r'^server\s*name$'
    ]
    for col in df.columns:
        col_str = str(col) 
        for pattern in preferred_patterns:
            if re.search(pattern, col_str.lower().strip()):
                return col
    patterns = [
        r'machine', r'computer', r'host', r'server', r'name'
    ]
    for col in df.columns:
        col_str = str(col)
        if 'machine' in col_str.lower() and 'name' in col_str.lower(): 
            return col
    for pattern in patterns:
        for col in df.columns:
            col_str = str(col)
            if re.search(pattern, col_str.lower().strip()):
                return col
    if len(df.columns) <= 3:
        for col in df.columns:
            if df[col].dtype == 'object' and not df[col].empty:
                samples = df[col].dropna()
                if not samples.empty:
                    match_count = 0
                    for sample_val in samples.head(5): 
                        if isinstance(sample_val, str) and re.match(r'^[a-zA-Z0-9.\-]+$', sample_val.strip()):
                            match_count += 1
                    if match_count > 0 : 
                        logger.info(f"Column '{col}' selected by content format heuristic.")
                        return col
    if len(df.columns) == 1:
        return df.columns[0]
    logger.warning("Could not definitively identify machine name column.")
    return None


def process_machine(machine_name):
    logger.info(f"Processing machine: {machine_name}")
    job_id = None 
    
    try:
        job_id = run_azure_runbook(machine_name)
        
        if not job_id:
            logger.error(f"Failed to start runbook for {machine_name}. No Job ID received.")
            error_id = f"error_{machine_name.replace('.', '_')}_{int(time.time())}"
            APP_STATE["completed_jobs"][error_id] = {
                "machine": machine_name, "status": "error",
                "error": "Failed to start runbook (no Job ID)",
                "completion_time": datetime.now().isoformat()
            }
            check_all_complete() 
            return
        
        APP_STATE["active_jobs"][job_id] = {
            "machine": machine_name, "start_time": datetime.now().isoformat(),
            "status": "running"
        }
        
        job_output_text = poll_job_status(job_id) 
        
        if job_id in APP_STATE["active_jobs"]:
            del APP_STATE["active_jobs"][job_id]

        if job_output_text is not None: 
            results = parse_job_output(job_output_text, machine_name) 
            APP_STATE["completed_jobs"][job_id] = {
                "machine": results.get("serverName", machine_name), 
                "status": results.get("executionStatus", "completed").lower(), 
                "completion_time": datetime.now().isoformat(), 
                "results": results 
            }
            APP_STATE["server_results"].append(results) 
            logger.info(f"Completed processing for {results.get('serverName', machine_name)}: {results.get('totalUpdates', 0)} updates, {len(results.get('failedUpdates', []))} failed. Execution Status: {results.get('executionStatus')}")
        else: 
            logger.error(f"Failed to get output or job failed for job {job_id} (machine: {machine_name}).")
            APP_STATE["completed_jobs"][job_id] = {
                "machine": machine_name, "status": "failed",
                "error": "Job failed, timed out, or no output received from Azure.",
                "completion_time": datetime.now().isoformat(),
                "results": {"serverName": machine_name, "executionStatus": "Failed", "errorMessage": "Job polling failed or timed out."} 
            }
        
        check_all_complete() 
    
    except Exception as e:
        logger.error(f"Unhandled error in process_machine for {machine_name}: {str(e)}")
        logger.exception(f"Traceback for process_machine {machine_name}:")
        
        if job_id and job_id in APP_STATE["active_jobs"]:
            del APP_STATE["active_jobs"][job_id]

        job_key_for_error = job_id if job_id else f"error_{machine_name.replace('.', '_')}_{int(time.time())}"
        APP_STATE["completed_jobs"][job_key_for_error] = {
            "machine": machine_name, "status": "error",
            "error": f"Unhandled exception: {str(e)}",
            "completion_time": datetime.now().isoformat(),
            "results": {"serverName": machine_name, "executionStatus": "Error", "errorMessage": str(e)} 
        }
        check_all_complete()


def update_requirements():
    missing_packages = []
    try:
        import openpyxl
        logger.debug("openpyxl found.")
    except ImportError:
        missing_packages.append("openpyxl")
    try:
        import xlrd
        logger.debug("xlrd found.")
    except ImportError:
        missing_packages.append("xlrd")

    if missing_packages:
        logger.info(f"Attempting to install missing Excel support libraries: {', '.join(missing_packages)}")
        import subprocess
        try:
            args = [sys.executable, "-m", "pip", "install"] + missing_packages
            subprocess.check_call(args)
            logger.info(f"Successfully installed {', '.join(missing_packages)}")
            if "openpyxl" in missing_packages:
                global openpyxl 
                import openpyxl
            if "xlrd" in missing_packages:
                global xlrd
                import xlrd
        except subprocess.CalledProcessError as e:
            logger.error(f"Failed to install Excel support libraries using pip: {str(e)}")
            logger.error("Please install them manually: pip install openpyxl xlrd")
        except Exception as e:
            logger.error(f"An unexpected error occurred during pip install: {str(e)}")
    else:
        logger.info("Excel support libraries (openpyxl, xlrd) are already installed.")


# --- Azure Integration ---

def run_azure_runbook(machine_name):
    logger.info(f"Running Azure runbook for {machine_name}")
    try:
        webhook_url = app.config['AZURE_WEBHOOK_URL']
        payload = {"ComputerName": machine_name}
        response = requests.post(webhook_url, json=payload, headers={"Content-Type": "application/json"}, timeout=30)
        
        if 200 <= response.status_code < 300:
            try:
                result = response.json()
                job_ids = result.get('JobIds') 
                if job_ids and isinstance(job_ids, list) and len(job_ids) > 0:
                    job_id = job_ids[0]
                    logger.info(f"Runbook started for {machine_name}, Job ID: {job_id}")
                    return job_id
                else:
                    logger.error(f"Runbook triggered for {machine_name}, but no JobIds returned in response. Response: {response.text}")
                    return None
            except json.JSONDecodeError:
                logger.error(f"Failed to decode JSON response from webhook for {machine_name}. Status: {response.status_code}, Response: {response.text}")
                return None
        else:
            logger.error(f"Failed to start runbook for {machine_name}: {response.status_code} - {response.text}")
            return None
    except requests.exceptions.Timeout:
        logger.error(f"Timeout triggering Azure runbook for {machine_name}.")
        return None
    except requests.exceptions.RequestException as e: 
        logger.error(f"Error running Azure runbook for {machine_name}: {str(e)}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error in run_azure_runbook for {machine_name}: {str(e)}")
        return None


def poll_job_status(job_id):
    logger.info(f"Polling job status for job ID: {job_id}")
    
    access_token = get_azure_token()
    if not access_token:
        logger.error(f"Failed to get Azure access token for polling job {job_id}. Cannot proceed.")
        return None 
        
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    
    subscription_id = app.config['AZURE_SUBSCRIPTION_ID']
    resource_group = app.config['AZURE_RESOURCE_GROUP']
    automation_account = app.config['AZURE_AUTOMATION_ACCOUNT']
    
    status_url = (f"https://management.azure.com/subscriptions/{subscription_id}"
                  f"/resourceGroups/{resource_group}/providers/Microsoft.Automation"
                  f"/automationAccounts/{automation_account}/jobs/{job_id}?api-version=2019-06-01")
    output_url = (f"https://management.azure.com/subscriptions/{subscription_id}"
                  f"/resourceGroups/{resource_group}/providers/Microsoft.Automation"
                  f"/automationAccounts/{automation_account}/jobs/{job_id}/output?api-version=2019-06-01")
    
    max_attempts = 30  
    poll_interval = 20 
    attempt = 0
    
    while attempt < max_attempts:
        logger.info(f"Polling attempt {attempt + 1}/{max_attempts} for job {job_id}...")
        try:
            response = requests.get(status_url, headers=headers, timeout=30) 
            response.raise_for_status() 
            
            job_status_data = response.json()
            status = job_status_data.get('properties', {}).get('status')
            logger.info(f"Job {job_id} current status: {status}")

            if status == 'Completed':
                logger.info(f"Job {job_id} completed. Fetching output.")
                try:
                    output_response = requests.get(output_url, headers=headers, timeout=60) 
                    output_response.raise_for_status()
                    raw_output = output_response.text 
                    logger.debug(f"Raw output received for job {job_id} (first 200 chars): {raw_output[:200]}")
                    return raw_output 
                except requests.exceptions.HTTPError as e_out:
                    logger.error(f"HTTP error fetching output for completed job {job_id}: {e_out.response.status_code} - {e_out.response.text}")
                    return "" 
                except requests.exceptions.RequestException as e_out:
                    logger.error(f"Request error fetching output for job {job_id}: {str(e_out)}")
                    return "" 
                except Exception as e_out:
                    logger.error(f"Unexpected error fetching output for job {job_id}: {str(e_out)}")
                    return ""

            elif status in ['Failed', 'Suspended', 'Stopped']:
                logger.error(f"Job {job_id} ended with unrecoverable status: {status}. Details: {job_status_data.get('properties', {}).get('exception', 'No exception details')}")
                return None 

        except requests.exceptions.HTTPError as e_stat:
            logger.error(f"HTTP error polling status for job {job_id}: {e_stat.response.status_code} - {e_stat.response.text}")
        except requests.exceptions.Timeout:
            logger.warning(f"Timeout polling status for job {job_id} on attempt {attempt + 1}.")
        except requests.exceptions.RequestException as e_stat: 
            logger.error(f"Request error polling status for job {job_id}: {str(e_stat)}")
        except json.JSONDecodeError:
            logger.error(f"Failed to decode JSON status response for job {job_id}. Response text: {response.text if 'response' in locals() else 'N/A'}")
        except Exception as e_stat: 
            logger.error(f"Unexpected error during job status poll for {job_id}: {str(e_stat)}")

        attempt += 1
        if attempt < max_attempts: 
            time.sleep(poll_interval)

    logger.error(f"Max polling attempts ({max_attempts}) reached for job {job_id}. Assuming job timed out or failed to complete in expected time.")
    return None 

def get_azure_token():
    client_id = app.config['AZURE_CLIENT_ID']
    client_secret = app.config['AZURE_CLIENT_SECRET']
    tenant_id = app.config['AZURE_TENANT_ID']
    
    if not all([client_id, client_secret, tenant_id]):
        logger.error("Azure client ID, secret, or tenant ID is not configured. Cannot obtain token.")
        return None

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/token"
    payload = {
        "grant_type": "client_credentials", "client_id": client_id,
        "client_secret": client_secret, "resource": "https://management.azure.com/"
    }
    try:
        response = requests.post(token_url, data=payload, timeout=20)
        response.raise_for_status() 
        
        token_data = response.json()
        access_token = token_data.get('access_token')
        if not access_token:
            logger.error(f"Access token not found in response from Azure. Response: {token_data}")
            return None
        return access_token
    except requests.exceptions.HTTPError as e:
        logger.error(f"Failed to get Azure token (HTTP error): {e.response.status_code} - {e.response.text}")
        return None
    except requests.exceptions.Timeout:
        logger.error("Timeout attempting to get Azure token.")
        return None
    except requests.exceptions.RequestException as e:
        logger.error(f"Error getting Azure token (RequestException): {str(e)}")
        return None
    except json.JSONDecodeError:
        logger.error(f"Failed to decode JSON response when getting Azure token. Response: {response.text if 'response' in locals() else 'N/A'}")
        return None
    except Exception as e: 
        logger.error(f"Unexpected error getting Azure token: {str(e)}")
        return None

# --- Output Parsing ---

def parse_job_output(job_output_text, server_name_fallback):
    logger.info(f"Attempting to parse job output for server fallback: {server_name_fallback}")
    
    parsed_results = {
        "serverName": server_name_fallback,
        "executionStatus": "ParsingError", 
        "errorMessage": "Failed to parse job output.",
        "updateHistory": [],
        "diagnosticChecks": {},
        "timestampUTC": datetime.utcnow().isoformat(), 
        "rawOutputExcerpt": (job_output_text[:500] + "..." if len(job_output_text) > 500 else job_output_text) if job_output_text else "No output text received"
    }

    if not job_output_text or not job_output_text.strip():
        logger.warning(f"No job output text to parse for {server_name_fallback}.")
        parsed_results["errorMessage"] = "No job output text received from runbook."
        return parsed_results

    try:
        data = json.loads(job_output_text)
        
        parsed_results["serverName"] = data.get("ComputerName", server_name_fallback)
        parsed_results["executionStatus"] = data.get("ExecutionStatus", "UnknownStatusFromScript")
        parsed_results["errorMessage"] = data.get("ErrorMessage") 
        parsed_results["timestampUTC"] = data.get("TimestampUTC", parsed_results["timestampUTC"]) 
        
        update_history_from_ps = data.get("UpdateHistory")
        if isinstance(update_history_from_ps, list):
            parsed_results["updateHistory"] = update_history_from_ps
        elif update_history_from_ps is not None: 
            logger.warning(f"UpdateHistory from PowerShell for {parsed_results['serverName']} was not a list: {type(update_history_from_ps)}")
            parsed_results["updateHistory"] = [] 
        
        diagnostic_checks_from_ps = data.get("DiagnosticChecks")
        if isinstance(diagnostic_checks_from_ps, dict):
            parsed_results["diagnosticChecks"] = diagnostic_checks_from_ps
        elif diagnostic_checks_from_ps is not None:
            logger.warning(f"DiagnosticChecks from PowerShell for {parsed_results['serverName']} was not a dict: {type(diagnostic_checks_from_ps)}")
            parsed_results["diagnosticChecks"] = {}

        parsed_results["allUpdates"] = parsed_results["updateHistory"] 
        parsed_results["totalUpdates"] = len(parsed_results["updateHistory"])
        parsed_results["failedUpdates"] = [
            upd for upd in parsed_results["updateHistory"] 
            if isinstance(upd, dict) and "fail" in upd.get("Status", "").lower()
        ]
        parsed_results["hasFailures"] = len(parsed_results["failedUpdates"]) > 0 or \
                                       "fail" in parsed_results["executionStatus"].lower() or \
                                       "error" in parsed_results["executionStatus"].lower()

        logger.info(f"Successfully parsed structured job output for {parsed_results['serverName']}. Execution Status: {parsed_results['executionStatus']}")
        if parsed_results["errorMessage"]:
            logger.warning(f"Error message from script for {parsed_results['serverName']}: {parsed_results['errorMessage']}")

    except json.JSONDecodeError as e:
        logger.error(f"Failed to decode JSON from job output for {server_name_fallback}: {e}")
        parsed_results["errorMessage"] = f"Output was not valid JSON: {e}. Raw output excerpt in 'rawOutputExcerpt'."
    except Exception as e:
        logger.error(f"Unexpected error parsing job output for {server_name_fallback}: {e}")
        logger.exception("Detailed traceback for parse_job_output error:")
        parsed_results["errorMessage"] = f"General parsing error: {e}. Raw output excerpt in 'rawOutputExcerpt'."

    return parsed_results


# --- Results Processing ---

def check_all_complete():
    if APP_STATE["active_jobs"]:
        logger.info(f"{len(APP_STATE['active_jobs'])} job(s) still active. Report generation deferred.")
        return

    if not APP_STATE["server_results"]:
        logger.info("All jobs complete, but no server results to report for this batch.")
        return 
    
    logger.info("All jobs complete for the current batch and server results are present. Aggregating and generating report.")
    try:
        current_server_results = list(APP_STATE["server_results"]) 
        
        aggregate_results = {
            "serverResults": current_server_results, # This is the list of parsed_results dicts
            "totalServersProcessed": len(current_server_results),
            # Other aggregate stats can be calculated here if needed by generate_ai_report directly
        }
        
        threading.Thread(target=generate_and_send_report, args=(aggregate_results,), daemon=True).start()
        logger.info("Report generation and sending initiated in a background thread.")
        
        APP_STATE["server_results"].clear()
        logger.info("Live server_results cleared after initiating report for current batch. Completed_jobs persist for UI.")

    except Exception as e:
        logger.error(f"Error in check_all_complete while preparing for report generation: {str(e)}")
        logger.exception("Detailed traceback for check_all_complete error:")
        APP_STATE["server_results"].clear()
        logger.warning("Cleared server_results after error in check_all_complete.")


def generate_and_send_report(aggregate_results_snapshot):
    report_content = "Error: Report generation failed." 
    servers_requiring_attention_count = 0 # Default
    try:
        logger.info("Background thread: Generating AI report.")
        # generate_ai_report now returns a tuple: (report_string, servers_requiring_attention_count)
        report_content, servers_requiring_attention_count = generate_ai_report(aggregate_results_snapshot) 
        
        logger.info(f"Background thread: AI Report generated. Servers requiring attention: {servers_requiring_attention_count}. Attempting to send email.")
        send_email_notification(report_content, servers_requiring_attention_count) # Pass the count
        logger.info("Background thread: Email notification sent successfully.")
    except Exception as e:
        logger.error(f"Background thread: Error during report generation or email sending: {str(e)}")
        logger.exception("Detailed traceback for generate_and_send_report background thread:")
        try:
            # Recalculate sra_count for template based on its own logic if AI failed early
            # For simplicity, we'll use the sra_count from the failed AI attempt if available,
            # or re-evaluate for the template.
            # The template report will internally filter and count.
            sra_count_for_fallback_subject = servers_requiring_attention_count # Use count from AI attempt if available
            
            if "Error: Report generation failed." in report_content or not report_content.strip():
                logger.info("Background thread: AI report generation seems to have failed, using template for fallback email.")
                # Template report will calculate its own sra_count for its content.
                # We need a reliable sra_count for the subject line of the fallback email.
                # Let's re-calculate based on the input snapshot for the subject.
                sra_count_for_fallback_subject = 0
                if aggregate_results_snapshot and aggregate_results_snapshot.get("serverResults"):
                    for sr_item in aggregate_results_snapshot["serverResults"]:
                        if sr_item.get("failedUpdates") and len(sr_item.get("failedUpdates")) > 0:
                            sra_count_for_fallback_subject += 1
                
                fallback_report_content = "ERROR: AI report generation failed or an error occurred during the main report process. A template-based summary is provided below.\n\n"
                fallback_report_content += generate_template_report(aggregate_results_snapshot) 
            else:
                fallback_report_content = "NOTICE: An error occurred after the initial report generation. The report content below might be incomplete or reflect the state before the error.\n\n" + report_content
            
            send_email_notification(fallback_report_content, sra_count_for_fallback_subject) # Use the determined count
            logger.info("Background thread: Fallback email notification sent.")
        except Exception as email_error:
            logger.error(f"Background thread: Failed to send fallback email notification: {str(email_error)}")
    finally:
        logger.info("Background thread: generate_and_send_report task finished.")


# --- AI Integration ---
def generate_ai_report(results_snapshot):
    logger.info("Generating AI report with new structure...")
    server_results_data = results_snapshot.get("serverResults", [])

    # Filter servers for the report: only those with failed updates
    servers_for_detailed_report = []
    all_server_data_with_flags = [] # For broader context if needed by AI for conclusion

    for sr_item in server_results_data:
        has_failed_updates = sr_item.get("failedUpdates") and len(sr_item.get("failedUpdates")) > 0
        
        # Check for any diagnostic issues on this server
        has_diagnostic_issues = False
        diagnostics = sr_item.get("diagnosticChecks", {})
        if "low" in diagnostics.get("DiskC", {}).get("Status", "OK").lower() or \
           diagnostics.get("PendingReboot", {}).get("IsPending") == True:
            has_diagnostic_issues = True
        else:
            for service_name, service_info in diagnostics.get("Services", {}).items():
                if isinstance(service_info, dict) and service_info.get("Status") != "Running" and service_info.get("Status") != "NotFound":
                    has_diagnostic_issues = True
                    break
            if not has_diagnostic_issues and ("issues" in diagnostics.get("ArcConnectivity", {}).get("Status", "OK").lower() or \
                                           "issuesfound" in diagnostics.get("CBSLog", {}).get("Status", "NoObviousErrors").lower()):
                has_diagnostic_issues = True
        
        all_server_data_with_flags.append({**sr_item, "_has_failed_updates": has_failed_updates, "_has_diagnostic_issues": has_diagnostic_issues})

        if has_failed_updates:
            servers_for_detailed_report.append(sr_item)

    total_servers_evaluated = len(server_results_data)
    servers_requiring_attention_count = len(servers_for_detailed_report)

    if not server_results_data:
        logger.warning("No server results data in snapshot to generate AI report from. Using template.")
        # Ensure template also adheres to new structure and returns count
        return generate_template_report(results_snapshot), 0 

    prompt = "You are a system administrator assistant. Generate a 'Server Patch Status and Health Check Report'.\n"
    prompt += "Adhere strictly to the following section titles and formatting style inspired by the user's example.\n\n"

    prompt += "Executive Summary\n"
    prompt += f"- Metric: Total Servers Evaluated, Count: {total_servers_evaluated}\n"
    prompt += f"- Metric: Servers Requiring Attention (due to failed updates), Count: {servers_requiring_attention_count}\n\n"

    prompt += "Detailed Issue Overview\n"
    if not servers_for_detailed_report:
        prompt += "All servers successfully applied updates. No servers require detailing for failed updates in this section.\n\n"
    else:
        for i, server_data in enumerate(servers_for_detailed_report):
            server_name = server_data.get("serverName", "Unknown Server")
            diagnostics = server_data.get("diagnosticChecks", {})
            failed_updates = server_data.get("failedUpdates", []) # Should have items

            prompt += f"{i + 1}. Server: {server_name}\n"
            
            # Failed Updates (Primary reason for inclusion)
            prompt += "    Failed Updates:\n"
            for upd in failed_updates:
                prompt += f"        - KB: {upd.get('UpdateKB', 'N/A')}, Title: {upd.get('Title', 'N/A')}, Status: {upd.get('Status', 'N/A')}\n"

            # Accompanying Diagnostic Issues for these servers
            disk_c_info = diagnostics.get("DiskC", {})
            if "low" in disk_c_info.get("Status", "OK").lower():
                prompt += f"    Disk Space Alert:\n        Drive C: {disk_c_info.get('Details', 'N/A')}\n"

            pending_reboot_info = diagnostics.get("PendingReboot", {})
            if pending_reboot_info.get("IsPending") == True:
                prompt += f"    Pending Actions:\n        Pending reboot ({pending_reboot_info.get('Details', 'N/A')})\n"
            
            service_issues_texts = []
            service_checks = diagnostics.get("Services", {})
            for svc_name, svc_info in service_checks.items():
                if isinstance(svc_info, dict):
                    s_status = svc_info.get("Status", "N/A")
                    s_start_type = svc_info.get("StartType", "N/A")
                    if s_status != "Running" and s_status != "NotFound":
                        service_issues_texts.append(f"{svc_name} {s_status.lower()} (Startup Type: {s_start_type})")
            
            arc_details = diagnostics.get("ArcConnectivity", {}).get("Details", "")
            if "issues" in diagnostics.get("ArcConnectivity", {}).get("Status", "OK").lower():
                 service_issues_texts.append(f"Azure Arc Connectivity: {arc_details if arc_details else diagnostics.get('ArcConnectivity')['Status']}")


            if service_issues_texts:
                prompt += "    Service Issues:\n"
                for issue_text in service_issues_texts:
                    prompt += f"        - {issue_text}\n"
            prompt += "\n" # Newline after each server in detailed overview
        prompt += "\n"


    prompt += "Actionable Recommendations\n"
    if not servers_for_detailed_report:
        prompt += "No specific recommendations as no servers reported failed updates requiring detailed attention.\n\n"
    else:
        prompt += "Based on the 'Detailed Issue Overview', provide numbered, categorized recommendations. For each, state the 'Immediate Action'.\n"
        prompt += "Example categories: Failed Update Remediation (specify server), Disk Management (specify server), Pending Reboot Resolution (specify server), Service Management (specify services and startup type changes but if a service is stopped and set to Manual as the start type, ignore since that wouldn't be running constantly anyways. If one is stopped but set to automatic that is an issue), Azure Arc Agent Verification (specify server). We do not use GuestConfigAgent service so don't add that into any checks\n"
        # Provide hints for recommendations based on the issues found in servers_for_detailed_report
        recommendation_hints = []
        for server_data in servers_for_detailed_report:
            s_name = server_data.get("serverName")
            recommendation_hints.append(f"For {s_name}: Address {len(server_data.get('failedUpdates',[]))} failed update(s).")
            # Add hints for other issues if present on these servers
            # (Disk, Reboot, Services, Arc as done in the detailed overview prompt construction)
        prompt += "Consider these points for recommendations: " + "; ".join(recommendation_hints) + "\n\n"


    prompt += "Conclusion\n"
    if servers_requiring_attention_count == 0:
        prompt += "All evaluated servers have successfully applied recent patches. "
        any_diag_issues_overall = any(s['_has_diagnostic_issues'] for s in all_server_data_with_flags if not s['_has_failed_updates'])
        if any_diag_issues_overall:
            prompt += "Some servers may have minor diagnostic alerts (not detailed here as updates were successful) that can be reviewed through system logs for routine maintenance.\n"
        else:
            prompt += "No critical issues or failed updates were detected that require immediate attention.\n"
    else:
        prompt += "Summarize the situation for the servers detailed in the 'Detailed Issue Overview'. State that these servers require attention for failed updates and any listed diagnostic issues. "
        prompt += "Comment on the patch status of other servers not detailed (if applicable, e.g., 'Other X servers successfully applied all patches.'). "
        prompt += "Emphasize the need to address the identified problems promptly to ensure operational stability and compliance.\n"
    
    prompt += "\nIMPORTANT: Ensure the output strictly follows the section titles: 'Executive Summary', 'Detailed Issue Overview', 'Actionable Recommendations', 'Conclusion'. Use formatting (like numbered lists, bullet points, bolding for emphasis)"

    logger.info(f"AI Prompt (length: {len(prompt)} chars). Preview (first 500):\n{prompt[:500]}...")
    
    ai_provider = app.config.get('AI_PROVIDER', 'template').lower() 
    logger.info(f"Using AI provider: {ai_provider}")
    
    report_text = ""
    if ai_provider == 'openai':
        report_text = generate_openai_report(prompt, servers_requiring_attention_count > 0)
    elif ai_provider == 'ollama':
        report_text = generate_ollama_report(prompt, servers_requiring_attention_count > 0)
    elif ai_provider == 'vllm':
        report_text = generate_vllm_report(prompt, servers_requiring_attention_count > 0)
    else:
        logger.info(f"AI provider '{ai_provider}' not configured or unknown. Using template report.")
        report_text = generate_template_report(results_snapshot) # Template report will also return string only now
    
    # If AI failed and returned template, sra_count might not be from AI logic.
    # The sra_count returned should be the one used for filtering for this report.
    return report_text, servers_requiring_attention_count


def _validate_ai_response(ai_generated_text, model_name, provider_name, actual_has_failed_updates): 
    if actual_has_failed_updates: 
        issue_keywords = ["fail", "failed", "failure", "unsuccessful", "error", "issue", "problem", "warning", "critical", "unreachable", "low space", "not running", "remediation", "address"]
        if not any(keyword in ai_generated_text.lower() for keyword in issue_keywords):
            warning_message = (
                f"⚠️ AI VALIDATION NOTICE: Failed updates were present, but the AI summary from "
                f"{provider_name} model '{model_name}' might not have explicitly acknowledged them or recommended actions. "
                f"Please review detailed logs and the AI output carefully. ⚠️\n\n"
            )
            logger.warning(f"AI Validation: {provider_name} model '{model_name}' response seems to miss issue mentions when failed updates exist.")
            return warning_message + ai_generated_text
    return ai_generated_text

def generate_ollama_report(prompt, actual_has_failed_updates): 
    ollama_url = app.config.get('OLLAMA_URL', 'http://host.docker.internal:11434/api/generate')
    ollama_model = app.config.get('OLLAMA_MODEL', 'mistral')
    logger.info(f"Connecting to Ollama at: {ollama_url} with model: {ollama_model}")
    
    request_body = {
        "model": ollama_model, "prompt": prompt, "stream": False,
        "options": {"temperature": 0.2, "top_p": 0.9} 
    }
    logger.debug(f"Ollama request body: {json.dumps(request_body, indent=2)}") 

    try:
        start_time = time.time()
        response = requests.post(ollama_url, json=request_body, timeout=180) 
        request_time = time.time() - start_time
        logger.info(f"Ollama request to model '{ollama_model}' completed in {request_time:.2f}s. Status: {response.status_code}")
        
        response.raise_for_status() 
        
        json_response = response.json()
        result = json_response.get('response', '')
        if not result.strip():
            logger.warning(f"Ollama model '{ollama_model}' returned an empty response.")
            return "ERROR: Ollama returned an empty response.\n\n" + generate_template_report({"serverResults": []}) 

        logger.info(f"Ollama report generated (length: {len(result)}). First 100 chars: {result[:100]}")
        return _validate_ai_response(result, ollama_model, "Ollama", actual_has_failed_updates) 

    except requests.exceptions.HTTPError as e:
        logger.error(f"Ollama HTTP error (model '{ollama_model}'): {e.response.status_code} - {e.response.text[:500]}")
        error_message = f"Ollama API error (model: {ollama_model}, status: {e.response.status_code}). "
        if e.response.status_code == 404: error_message += f"Model '{ollama_model}' not found or API endpoint incorrect."
        else: error_message += "Check Ollama server logs and connectivity."
        return f"ERROR: {error_message}\n\n" + generate_template_report({"serverResults": []})
    except requests.exceptions.Timeout:
        logger.error(f"Timeout connecting to Ollama (model '{ollama_model}') at {ollama_url}.")
        return f"ERROR: Timeout connecting to Ollama AI service.\n\n" + generate_template_report({"serverResults": []})
    except requests.exceptions.RequestException as e:
        logger.error(f"Error connecting to Ollama (model '{ollama_model}'): {str(e)}")
        return f"ERROR: Cannot connect to Ollama AI service ({str(e)}).\n\n" + generate_template_report({"serverResults": []})
    except json.JSONDecodeError as e:
        logger.error(f"Error parsing Ollama JSON response (model '{ollama_model}'): {str(e)}. Response text: {response.text[:500] if 'response' in locals() else 'N/A'}")
        return f"ERROR: Malformed response from Ollama.\n\n" + generate_template_report({"serverResults": []})
    except Exception as e:
        logger.error(f"General error with Ollama (model '{ollama_model}'): {str(e)}")
        logger.exception("Detailed traceback for Ollama error:")
        return f"ERROR: An unexpected error occurred with Ollama AI service.\n\n" + generate_template_report({"serverResults": []})


def generate_openai_report(prompt, actual_has_failed_updates): 
    openai.api_key = app.config.get('OPENAI_API_KEY')
    openai_model = app.config.get('OPENAI_MODEL', 'gpt-3.5-turbo')

    if not openai.api_key or openai.api_key == 'your-openai-api-key':
        logger.error("OpenAI API key is not configured.")
        return "ERROR: OpenAI API key not configured.\n\n" + generate_template_report({"serverResults": []})

    logger.info(f"Generating report with OpenAI model: {openai_model}")
    try:
        start_time = time.time()
        client = openai.OpenAI(api_key=openai.api_key)
        response = client.chat.completions.create(
            model=openai_model,
            messages=[
                {"role": "system", "content": "You are a system administrator expert. Generate a 'Server Patch Status and Health Check Report' following the user's detailed structure and instructions precisely."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2, # Slightly lower for more deterministic structured output
            max_tokens=2500 
        )
        request_time = time.time() - start_time
        logger.info(f"OpenAI request completed in {request_time:.2f}s.")

        report_content = response.choices[0].message.content
        if not report_content.strip():
             logger.warning(f"OpenAI model '{openai_model}' returned an empty response.")
             return "ERROR: OpenAI returned an empty response.\n\n" + generate_template_report({"serverResults": []})

        logger.info(f"OpenAI report generated (length: {len(report_content)}). First 100 chars: {report_content[:100]}")
        return _validate_ai_response(report_content, openai_model, "OpenAI", actual_has_failed_updates) 

    except openai.APIConnectionError as e:
        logger.error(f"OpenAI API Connection Error: {str(e)}")
        return f"ERROR: Failed to connect to OpenAI API: {str(e)}\n\n" + generate_template_report({"serverResults": []})
    except openai.RateLimitError as e:
        logger.error(f"OpenAI Rate Limit Error: {str(e)}")
        return f"ERROR: OpenAI API rate limit exceeded: {str(e)}\n\n" + generate_template_report({"serverResults": []})
    except openai.APIStatusError as e: 
        logger.error(f"OpenAI API Status Error: Status {e.status_code}, Response: {e.response}")
        return f"ERROR: OpenAI API returned an error (Status {e.status_code}): {e.message}\n\n" + generate_template_report({"serverResults": []})
    except Exception as e: 
        logger.error(f"Unexpected error with OpenAI: {str(e)}")
        logger.exception("Detailed traceback for OpenAI error:")
        return f"ERROR: Error generating report with OpenAI: {str(e)}\n\n" + generate_template_report({"serverResults": []})


def generate_vllm_report(prompt, actual_has_failed_updates): 
    logger.info("Attempting to generate report using VLLM")
    vllm_url = app.config.get('VLLM_CHAT_COMPLETIONS_URL')
    vllm_model = app.config.get('VLLM_MODEL')
    vllm_api_key = app.config.get('VLLM_API_KEY') 
    vllm_verify_ssl = app.config.get('VLLM_VERIFY_SSL', True) 

    if not vllm_url or not vllm_model:
        logger.error("VLLM URL or model not configured.")
        return "ERROR: VLLM URL or model not configured.\n\n" + generate_template_report({"serverResults": []})

    headers = {"Content-Type": "application/json"}
    if vllm_api_key:
        headers["Authorization"] = f"Bearer {vllm_api_key}"
    
    payload = {
        "model": vllm_model,
        "messages": [
            {"role": "system", "content": "You are a system administrator expert. Generate a 'Server Patch Status and Health Check Report' following the user's detailed structure and instructions precisely."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2, 
        "max_tokens": 2500 
    }
    
    logger.info(f"Sending request to VLLM at {vllm_url} with model {vllm_model}. SSL Verify: {vllm_verify_ssl}")
    logger.debug(f"VLLM request payload: {json.dumps(payload, indent=2)}")
    try:
        start_time = time.time()
        response = requests.post(vllm_url, headers=headers, json=payload, timeout=180, verify=vllm_verify_ssl) 
        request_time = time.time() - start_time
        logger.info(f"VLLM request completed in {request_time:.2f}s. Status: {response.status_code}")

        response.raise_for_status() 
        json_response = response.json()
        
        if json_response.get("choices") and len(json_response["choices"]) > 0:
            message = json_response["choices"][0].get("message")
            if message and message.get("content"):
                report_content = message["content"]
                if not report_content.strip():
                    logger.warning(f"VLLM model '{vllm_model}' returned an empty content string.")
                    return "ERROR: VLLM returned an empty response content.\n\n" + generate_template_report({"serverResults": []})
                
                logger.info(f"VLLM report generated (length: {len(report_content)}). First 100 chars: {report_content[:100]}")
                return _validate_ai_response(report_content, vllm_model, "VLLM", actual_has_failed_updates) 
            else:
                logger.error(f"VLLM response missing expected content structure. Message: {message}")
                return "ERROR: VLLM response malformed (no content in message).\n\n" + generate_template_report({"serverResults": []})
        else:
            logger.error(f"VLLM response missing 'choices' or choices are empty. Response: {json_response}")
            return "ERROR: VLLM response malformed (no choices or empty choices array).\n\n" + generate_template_report({"serverResults": []})

    except requests.exceptions.SSLError as e_ssl:
        logger.error(f"VLLM SSL Error: {str(e_ssl)}. Check VLLM_VERIFY_SSL setting in config.py or environment.")
        return f"ERROR: SSL connection error with VLLM service: {str(e_ssl)}\n\n" + generate_template_report({"serverResults": []})
    except requests.exceptions.HTTPError as e:
        logger.error(f"VLLM HTTP Error: {e.response.status_code} - {e.response.text[:500]}")
        return f"ERROR: VLLM API returned an HTTP error: {e.response.status_code}. Details: {e.response.text[:200]}\n\n" + generate_template_report({"serverResults": []})
    except requests.exceptions.Timeout:
        logger.error(f"Timeout connecting to VLLM service at {vllm_url}.")
        return f"ERROR: Timeout connecting to VLLM service.\n\n" + generate_template_report({"serverResults": []})
    except requests.exceptions.RequestException as e:
        logger.error(f"Error connecting to VLLM service: {str(e)}")
        return f"ERROR: Cannot connect to VLLM service ({str(e)}).\n\n" + generate_template_report({"serverResults": []})
    except json.JSONDecodeError as e:
        logger.error(f"Error parsing VLLM JSON response: {str(e)}. Response text: {response.text[:500] if 'response' in locals() else 'N/A'}")
        return f"ERROR: Malformed JSON response from VLLM.\n\n" + generate_template_report({"serverResults": []})
    except Exception as e:
        logger.error(f"General error with VLLM: {str(e)}")
        logger.exception("Detailed traceback for VLLM error:")
        return f"ERROR: An unexpected error occurred with VLLM service.\n\n" + generate_template_report({"serverResults": []})


def generate_template_report(results_data): 
    if results_data is None: results_data = {} 

    server_results_list_from_data = results_data.get("serverResults", [])
    servers_for_template_report = []
    all_server_data_with_flags_template = []

    for sr_item in server_results_list_from_data:
        has_failed_updates = sr_item.get("failedUpdates") and len(sr_item.get("failedUpdates")) > 0
        
        has_diagnostic_issues = False
        diagnostics = sr_item.get("diagnosticChecks", {})
        if "low" in diagnostics.get("DiskC", {}).get("Status", "OK").lower() or \
           diagnostics.get("PendingReboot", {}).get("IsPending") == True:
            has_diagnostic_issues = True
        # Simplified check for other diagnostics for template
        elif diagnostics.get("Services") or diagnostics.get("ArcConnectivity") or diagnostics.get("CBSLog"): 
            has_diagnostic_issues = True # Assume any entry here is an issue for template simplicity

        all_server_data_with_flags_template.append({**sr_item, "_has_failed_updates": has_failed_updates, "_has_diagnostic_issues": has_diagnostic_issues})

        if has_failed_updates:
            servers_for_template_report.append(sr_item)

    total_servers_evaluated_template = len(server_results_list_from_data)
    servers_requiring_attention_template = len(servers_for_template_report)

    report = "Server Patch Status and Health Check Report (Template)\n"
    report += "=======================================================\n\n"
    report += f"Date of Report: {datetime.now().strftime('%Y-%m-%d %H:%M:%S %Z')}\n\n"
    
    report += "Executive Summary\n"
    report += "-----------------\n"
    report += f"Total Servers Evaluated: {total_servers_evaluated_template}\n"
    report += f"Servers Requiring Attention (due to failed updates): {servers_requiring_attention_template}\n\n"
    
    report += "Detailed Issue Overview\n"
    report += "-----------------------\n"
    if not servers_for_template_report:
        report += "All servers successfully applied updates. No servers require detailing for failed updates in this section.\n\n"
    else:
        for i, server_info in enumerate(servers_for_template_report):
            server_name = server_info.get("serverName", "Unknown Server")
            diagnostics = server_info.get("diagnosticChecks", {})
            failed_updates = server_info.get("failedUpdates", [])

            report += f"{i + 1}. Server: {server_name}\n"
            
            if failed_updates:
                report += "    Failed Updates:\n"
                for upd in failed_updates:
                    report += f"        - KB: {upd.get('UpdateKB', 'N/A')}, Title: {upd.get('Title', 'N/A')}, Status: {upd.get('Status', 'N/A')}\n"

            disk_info = diagnostics.get("DiskC", {})
            if "low" in disk_info.get("Status", "OK").lower():
                report += f"    Disk Space Alert:\n        Drive C: {disk_info.get('Details', 'N/A')}\n"

            pending_reboot_info = diagnostics.get("PendingReboot", {})
            if pending_reboot_info.get("IsPending") == True:
                report += f"    Pending Actions:\n        Pending reboot ({pending_reboot_info.get('Details', 'N/A')})\n"
            
            service_issues_texts = []
            service_checks = diagnostics.get("Services", {})
            for svc_name, svc_info in service_checks.items():
                if isinstance(svc_info, dict):
                    s_status = svc_info.get("Status", "N/A")
                    s_start_type = svc_info.get("StartType", "N/A")
                    if s_status != "Running" and s_status != "NotFound":
                        service_issues_texts.append(f"{svc_name} {s_status.lower()} (Startup Type: {s_start_type})")
            
            arc_details = diagnostics.get("ArcConnectivity", {}).get("Details", "")
            if "GuestConfigAgent service not found" in arc_details or "GuestConfigAgent not installed" in arc_details:
                 service_issues_texts.append("GuestConfigAgent not installed")
            elif "issues" in diagnostics.get("ArcConnectivity", {}).get("Status", "OK").lower():
                 service_issues_texts.append(f"Azure Arc Connectivity: {arc_details if arc_details else diagnostics.get('ArcConnectivity')['Status']}")

            if service_issues_texts:
                report += "    Service Issues:\n"
                for issue_text in service_issues_texts:
                    report += f"        - {issue_text}\n"
            report += "\n"
    report += "\n"

    report += "Actionable Recommendations\n"
    report += "--------------------------\n"
    if not servers_for_template_report:
        report += "No specific recommendations as no servers reported failed updates.\n"
    else:
        report += "- For servers listed in 'Detailed Issue Overview':\n"
        report += "  - Investigate and remediate all listed 'Failed Updates'.\n"
        report += "  - Address any 'Disk Space Alert' by freeing up or expanding storage.\n"
        report += "  - Perform required actions for 'Pending Actions' (e.g., schedule reboots).\n"
        report += "  - For 'Service Issues': Start stopped services. Consider changing startup types to 'Automatic' for critical services. Investigate 'GuestConfigAgent not installed' by checking Azure Arc agent status.\n"
    report += "\n"

    report += "Conclusion\n"
    report += "----------\n"
    if servers_requiring_attention_template == 0:
        report += "All evaluated servers have successfully applied recent patches. "
        any_diag_issues_overall_template = any(s['_has_diagnostic_issues'] for s in all_server_data_with_flags_template if not s['_has_failed_updates'])
        if any_diag_issues_overall_template:
            report += "Review system logs for any minor diagnostic alerts on successfully patched servers for routine maintenance.\n"
        else:
            report += "No critical issues or failed updates were detected.\n"
    else:
        report += f"{servers_requiring_attention_template} server(s) require attention for failed updates and potentially other listed issues. "
        other_successful_count = total_servers_evaluated_template - servers_requiring_attention_template
        if other_successful_count > 0:
            report += f"The other {other_successful_count} server(s) successfully applied all patches. "
        report += "Address the identified problems promptly.\n"
    
    report += "\nEnd of Template Report.\n"
    return report

# --- Email Notification ---

def send_email_notification(report_body, servers_requiring_attention_count): # Added servers_requiring_attention_count
    logger.info("Attempting to send email notification.")
    try:
        from_email = app.config['EMAIL_FROM']
        to_email = app.config['EMAIL_TO']
        smtp_server = app.config['SMTP_SERVER']
        smtp_port = int(app.config.get('SMTP_PORT', 25)) 
        smtp_user = app.config.get('SMTP_USERNAME') 
        smtp_pass = app.config.get('SMTP_PASSWORD') 
        
        if not all([from_email, to_email, smtp_server]):
            logger.error("Email configuration (FROM, TO, SERVER) is incomplete. Cannot send email.")
            return

        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        report_date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        # Updated subject status logic
        subject_status = "Action Required" if servers_requiring_attention_count > 0 else "All Clear"
        msg['Subject'] = f'PatchMate Server Health & Update Report - {report_date_str} - Status: {subject_status}'
        
        msg.attach(MIMEText(report_body, 'plain', 'utf-8')) 
        
        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server: 
            server.ehlo() 
            if smtp_user and smtp_pass: 
                try:
                    server.starttls() 
                    server.ehlo() 
                    server.login(smtp_user, smtp_pass)
                    logger.info("SMTP connection secured with STARTTLS and logged in.")
                except smtplib.SMTPException as e_tls:
                    logger.warning(f"Failed to use STARTTLS or login with SMTP server {smtp_server} (User: {smtp_user}). Error: {e_tls}. Will try unencrypted if allowed by server.")
            
            server.send_message(msg)
        logger.info(f"Email notification sent successfully to {to_email}.")

    except smtplib.SMTPAuthenticationError as e_auth:
        logger.error(f"SMTP Authentication Error: {str(e_auth)}. Check username/password for {smtp_server}.")
    except smtplib.SMTPConnectError as e_conn:
        logger.error(f"SMTP Connection Error: Failed to connect to server {smtp_server}:{smtp_port}. Error: {str(e_conn)}")
    except smtplib.SMTPServerDisconnected as e_dis:
        logger.error(f"SMTP Server Disconnected: Connection to {smtp_server} was lost. Error: {str(e_dis)}")
    except smtplib.SMTPException as e_smtp: 
        logger.error(f"SMTP Error: {str(e_smtp)} when sending email via {smtp_server}.")
    except ConnectionRefusedError: 
        logger.error(f"Connection Refused: Cannot connect to SMTP server {smtp_server}:{smtp_port}. Check server address and port, and if server is running.")
    except TimeoutError: 
        logger.error(f"Timeout sending email via {smtp_server}:{smtp_port}.")
    except Exception as e:
        logger.error(f"Unexpected error sending email: {str(e)}")
        logger.exception("Detailed traceback for email sending failure:")


# --- Web Interface Routes ---
@app.route('/')
def index():
    return render_template(
        'index.html',
        monitor_active=APP_STATE["monitor_active"],
        active_jobs=APP_STATE["active_jobs"],
        completed_jobs=APP_STATE["completed_jobs"],
        config=app.config, 
        current_ai_provider_from_config=app.config.get('AI_PROVIDER') 
    )

@app.route('/api/start-monitoring', methods=['POST'])
def api_start_monitoring():
    result = start_file_monitoring()
    status_code = 200 if result.get("status") == "started" or result.get("status") == "already_running" else 500
    return jsonify(result), status_code

@app.route('/api/stop-monitoring', methods=['POST'])
def api_stop_monitoring():
    result = stop_file_monitoring()
    status_code = 200 if result.get("status") == "stopped" or result.get("status") == "not_running" else 500
    return jsonify(result), status_code

@app.route('/api/process-file', methods=['POST']) 
def api_process_file():
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "No file part in the request."}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"status": "error", "message": "No file selected."}), 400
    
    original_filename = file.filename
    secured_filename = secure_filename(original_filename) 
    if not secured_filename: 
        secured_filename = f"uploaded_file_{int(time.time())}" 
    
    file_ext = os.path.splitext(secured_filename)[1].lower()
    if file_ext not in ['.csv', '.xlsx', '.xls']:
        return jsonify({"status": "error", "message": f"Unsupported file type: '{file_ext}'. Please upload CSV or Excel files."}), 415 
    
    try:
        upload_dir = app.config.get('UPLOAD_DIRECTORY', 'uploads')
        if not ensure_directory_exists(upload_dir):
            return jsonify({"status": "error", "message": f"Server error: Failed to create upload directory '{upload_dir}'."}), 500
            
        file_path = os.path.join(upload_dir, secured_filename)
        
        counter = 1
        base_name, ext = os.path.splitext(file_path)
        while os.path.exists(file_path):
            file_path = f"{base_name}_{counter}{ext}"
            counter += 1
            if counter > 100: 
                 logger.error(f"Too many duplicate filenames for {secured_filename} in {upload_dir}")
                 return jsonify({"status": "error", "message": "Server error: Could not save file due to naming conflict."}), 500

        file.save(file_path)
        logger.info(f"File '{original_filename}' (saved as '{os.path.basename(file_path)}') uploaded to {file_path}. Initiating processing.")
        
        threading.Thread(target=process_machine_file, args=(file_path,), daemon=True).start()
        
        return jsonify({"status": "success", "message": f"File '{original_filename}' received and processing started."}), 202 
    except Exception as e:
        logger.error(f"Error processing uploaded file '{original_filename}': {str(e)}")
        logger.exception("Detailed traceback for upload processing error:")
        return jsonify({"status": "error", "message": f"Server error processing file: {str(e)}"}), 500


@app.route('/api/status')
def api_status():
    active_jobs_list = [{"id": jid, **info} for jid, info in APP_STATE["active_jobs"].items()]
    
    completed_jobs_list = []
    try:
        sorted_completed_jobs = sorted(
            APP_STATE["completed_jobs"].items(), 
            key=lambda item: (item[1].get("completion_time", "0") if item[1].get("completion_time") else "0", item[0]),
            reverse=True
        )
    except Exception as e: 
        logger.warning(f"Could not sort completed jobs, using unsorted: {e}")
        sorted_completed_jobs = APP_STATE["completed_jobs"].items()

    for job_id, job_info in sorted_completed_jobs:
        job_data = {"id": job_id, **job_info} 
        
        if "results" in job_info and isinstance(job_info["results"], dict):
            results_summary = job_info["results"]
            job_data["total_updates_in_job"] = results_summary.get("totalUpdates", 0)
            job_data["failed_updates_in_job"] = len(results_summary.get("failedUpdates", []))
            job_data["execution_status_from_script"] = results_summary.get("executionStatus", "Unknown")
            job_data["script_error_message"] = results_summary.get("errorMessage")
            job_data["raw_output_excerpt"] = results_summary.get("rawOutputExcerpt", "N/A")

        job_data["status"] = job_info.get("status", "unknown") 
        if "results" in job_info and isinstance(job_info["results"], dict) and job_info["results"].get("executionStatus"):
             job_data["status"] = job_info["results"]["executionStatus"].lower() 

        completed_jobs_list.append(job_data)

    return jsonify({
        "monitor_active": APP_STATE["monitor_active"],
        "active_jobs_count": len(active_jobs_list),
        "active_jobs": active_jobs_list, 
        "completed_jobs_count": len(completed_jobs_list),
        "completed_jobs": completed_jobs_list, 
        "current_server_results_batch_size": len(APP_STATE["server_results"]),
        "current_ai_provider": app.config.get('AI_PROVIDER') 
    }), 200


@app.route('/api/reset-state', methods=['POST']) 
def api_reset_state():
    logger.info("API call to reset application state received.")
    if APP_STATE["monitor_active"] and APP_STATE["observer"]:
        logger.info("Stopping active file monitoring before full state reset.")
        stop_file_monitoring() 

    initialize_app_state() 
    logger.info("Application state has been fully reset via API.")
    return jsonify({"status": "success", "message": "Application state has been reset."}), 200

@app.route('/api/set-ai-provider', methods=['POST'])
def api_set_ai_provider():
    try:
        data = request.json
        new_provider = data.get('provider')
        
        if not new_provider or new_provider.lower() not in ['ollama', 'vllm', 'openai', 'template']:
            return jsonify({"status": "error", "message": "Invalid or missing AI provider specified."}), 400
        
        app.config['AI_PROVIDER'] = new_provider.lower()
        logger.info(f"AI Provider changed to: {app.config['AI_PROVIDER']} (in-memory config).")
        
        return jsonify({
            "status": "success",
            "message": f"AI Provider set to {app.config['AI_PROVIDER']}.",
            "current_ai_provider": app.config['AI_PROVIDER']
        }), 200
    except Exception as e:
        logger.error(f"Error setting AI provider: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/ai-provider-models') 
def api_ai_provider_models():
    ai_provider = app.config.get('AI_PROVIDER', 'template').lower() 
    current_model_key = None
    fetch_url = None
    provider_name_for_log = "Unknown"
    current_model_value = "N/A"
    vllm_verify_ssl = app.config.get('VLLM_VERIFY_SSL', True) 

    if ai_provider == 'ollama':
        ollama_url_config = app.config.get('OLLAMA_URL', 'http://host.docker.internal:11434/api/generate')
        base_url_parts = ollama_url_config.split('/api/')
        if len(base_url_parts) < 1: 
            logger.error(f"OLLAMA_URL format is unexpected: {ollama_url_config}")
            return jsonify({"status": "error", "provider": ai_provider, "message": "Ollama base URL misconfigured."}), 500
        base_url = base_url_parts[0]
        fetch_url = f"{base_url}/api/tags" 
        current_model_key = 'OLLAMA_MODEL'
        provider_name_for_log = "Ollama"
        current_model_value = app.config.get(current_model_key, 'mistral')
    elif ai_provider == 'vllm':
        fetch_url = app.config.get('VLLM_MODELS_URL')
        current_model_key = 'VLLM_MODEL'
        provider_name_for_log = "VLLM"
        current_model_value = app.config.get(current_model_key, 'your-default-vllm-model')
    elif ai_provider == 'openai':
        logger.info("OpenAI model listing requested. Returning configured model and common alternatives.")
        common_openai_models = [
            {"name": "gpt-4", "description": "Latest generation, most capable."},
            {"name": "gpt-4-turbo", "description": "Updated GPT-4 with wider context."},
            {"name": "gpt-3.5-turbo", "description": "Fast and cost-effective for many tasks."}
        ]
        current_openai_model = app.config.get('OPENAI_MODEL', 'gpt-3.5-turbo')
        if not any(m['name'] == current_openai_model for m in common_openai_models):
            common_openai_models.append({"name": current_openai_model, "description": "Currently configured model."})
        
        return jsonify({
            "status": "success", 
            "provider": "openai",
            "current_model": current_openai_model,
            "models": common_openai_models 
        }), 200
    else: 
        return jsonify({
            "status": "success", 
            "provider": ai_provider,
            "current_model": "N/A (template or unknown provider)",
            "models": []
        }), 200

    if not fetch_url:
        logger.warning(f"No model fetch URL defined for AI provider: {ai_provider}")
        return jsonify({"status": "error", "provider": ai_provider, "message": f"Model listing not supported or configured for {ai_provider}."}), 404

    logger.info(f"Querying {provider_name_for_log} models at: {fetch_url}. SSL Verify: {vllm_verify_ssl if ai_provider == 'vllm' else 'N/A'}")
    try:
        headers = {}
        request_kwargs = {'headers': headers, 'timeout': 15}
        if ai_provider == 'vllm':
            if app.config.get('VLLM_API_KEY'):
                headers["Authorization"] = f"Bearer {app.config['VLLM_API_KEY']}"
            request_kwargs['verify'] = vllm_verify_ssl 

        response = requests.get(fetch_url, **request_kwargs)
        response.raise_for_status()
        
        data = response.json()
        models_list = []

        if ai_provider == 'ollama' and 'models' in data:
            models_list = [{'name': m.get('name'), 'modified_at': m.get('modified_at'), 'size': m.get('size')}
                           for m in data.get('models', [])]
        elif ai_provider == 'vllm' and 'data' in data and isinstance(data['data'], list): 
            models_list = [{'name': m.get('id'), 'owned_by': m.get('owned_by', 'unknown'), 'object_type': m.get('object')}
                           for m in data['data']]
        else:
            logger.warning(f"Unexpected response format from {provider_name_for_log} model endpoint: {data}")
            models_list = [] 

        logger.info(f"Found {len(models_list)} models for {provider_name_for_log}.")
        return jsonify({
            "status": "success",
            "provider": ai_provider,
            "current_model": current_model_value,
            "models": models_list
        }), 200
    except requests.exceptions.SSLError as e_ssl:
        logger.error(f"{provider_name_for_log} SSL Error: {str(e_ssl)}. Check VLLM_VERIFY_SSL setting.")
        return jsonify({"status": "error", "provider": ai_provider, "message": f"SSL connection error with {provider_name_for_log}: {str(e_ssl)}"}), 503
    except requests.exceptions.HTTPError as e:
        logger.error(f"HTTP error querying {provider_name_for_log} models: {e.response.status_code} - {e.response.text[:200]}")
        return jsonify({"status": "error", "provider": ai_provider, "message": f"Failed to get models from {provider_name_for_log}: HTTP {e.response.status_code}"}), e.response.status_code
    except requests.exceptions.RequestException as e:
        logger.error(f"Error connecting to {provider_name_for_log} for models: {str(e)}")
        return jsonify({"status": "error", "provider": ai_provider, "message": f"Cannot connect to {provider_name_for_log}: {str(e)}"}), 503 
    except Exception as e:
        logger.error(f"Unexpected error getting {provider_name_for_log} models: {str(e)}")
        return jsonify({"status": "error", "provider": ai_provider, "message": str(e)}), 500


@app.route('/api/set-ai-model', methods=['POST']) 
def api_set_ai_model():
    try:
        data = request.json
        model_name = data.get('model')
        provider = app.config.get('AI_PROVIDER', 'template').lower()

        if not model_name:
            return jsonify({"status": "error", "message": "No model name provided."}), 400
        
        config_key_to_set = None
        if provider == 'ollama':
            config_key_to_set = 'OLLAMA_MODEL'
        elif provider == 'vllm':
            config_key_to_set = 'VLLM_MODEL'
        elif provider == 'openai':
            config_key_to_set = 'OPENAI_MODEL'
        else:
            logger.warning(f"Attempt to set model for an unsupported or template provider: {provider}")
            return jsonify({"status": "error", "message": f"Cannot set model for provider '{provider}'."}), 400

        app.config[config_key_to_set] = model_name
        logger.info(f"AI model for provider '{provider}' changed to: {model_name} (in-memory config).")
        
        return jsonify({"status": "success", "message": f"Model for {provider} changed to {model_name}.", "current_model": model_name, "provider": provider}), 200
    except Exception as e:
        logger.error(f"Error setting AI model: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500
    

@app.route('/api/clear-transient-data', methods=['POST']) 
def api_clear_transient_data():
    APP_STATE["completed_jobs"].clear() 
    APP_STATE["server_results"].clear() 
    APP_STATE["processed_files_on_startup"].clear() 
    logger.info("Transient data (completed jobs, server results, processed file list for session) cleared via API.")
    return jsonify({"status": "success", "message": "Completed jobs list, current server results batch, and session's processed file list have been cleared."}), 200


@app.route('/api/debug-info', methods=['GET']) 
def api_debug_info():
    try:
        debug_info = {
            "app_version": "1.2.3 (Report Format Update)", 
            "flask_debug_mode": app.debug,
            "current_time_utc": datetime.utcnow().isoformat(),
            "monitor_active": APP_STATE["monitor_active"],
            "active_jobs_count": len(APP_STATE["active_jobs"]),
            "completed_jobs_count": len(APP_STATE["completed_jobs"]),
            "server_results_batch_size": len(APP_STATE["server_results"]), 
            "processed_files_in_session_count": len(APP_STATE["processed_files_on_startup"]),
            "python_version": sys.version,
            "platform": sys.platform,
            "cwd": os.getcwd(),
            "current_runtime_ai_provider": app.config.get('AI_PROVIDER'), 
            "initial_config_ai_provider": app.config.get('AI_PROVIDER', os.getenv('AI_PROVIDER', 'ollama')), 
            "config_ollama_model": app.config.get('OLLAMA_MODEL'),
            "config_vllm_model": app.config.get('VLLM_MODEL'),
            "config_openai_model": app.config.get('OPENAI_MODEL'),
            "config_watch_dir": app.config.get('WATCH_DIRECTORY'),
            "config_upload_dir": app.config.get('UPLOAD_DIRECTORY'),
            "config_vllm_verify_ssl": app.config.get('VLLM_VERIFY_SSL') 
        }
        
        test_ai_param = request.args.get('test_ai_report', 'false').lower()
        if test_ai_param == 'true':
            sample_results_for_test_report = []
            if APP_STATE["server_results"]: # Prefer current batch if available
                 sample_results_for_test_report = APP_STATE["server_results"]
            elif APP_STATE["completed_jobs"]:
                 sample_results_for_test_report = [job.get("results") for job in list(APP_STATE["completed_jobs"].values()) if job.get("results")]

            if sample_results_for_test_report:
                test_aggregate_results = {
                    "serverResults": sample_results_for_test_report[:3], 
                }
                report_test_content, sra_count = generate_ai_report(test_aggregate_results)
                debug_info["test_ai_report_preview"] = report_test_content[:1000] + "..." if len(report_test_content) > 1000 else report_test_content
                debug_info["test_ai_report_provider_used"] = app.config.get('AI_PROVIDER')
                debug_info["test_ai_report_sra_count"] = sra_count
            else:
                debug_info["test_ai_report_message"] = "No server results or completed jobs with results available to generate a test report."
        
        return jsonify(debug_info), 200
    except Exception as e:
        logger.error(f"Error in /api/debug-info: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/reload-templates', methods=['POST']) 
def reload_template_cache():
    try:
        app.jinja_env.cache = {} 
        logger.info("Jinja2 template cache cleared via API. A page refresh may be needed.")
        return jsonify({"status": "success", "message": "Template cache cleared. Refresh your browser to see changes."}), 200
    except Exception as e:
        logger.error(f"Error clearing template cache: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500
    
@app.route('/api/system-info')
def api_system_info():
    try:
        watch_dir_config = app.config.get('WATCH_DIRECTORY', 'Not Set')
        upload_dir_config = app.config.get('UPLOAD_DIRECTORY', 'uploads') 

        excel_support_status = "Not installed"
        try:
            import openpyxl
            excel_support_status = f"Installed (openpyxl {openpyxl.__version__})"
            import xlrd 
            excel_support_status += f", (xlrd {xlrd.__version__ if hasattr(xlrd, '__version__') else 'version unknown'})"
        except ImportError:
            try:
                import openpyxl
                excel_support_status = f"Partial: openpyxl {openpyxl.__version__} (xlrd missing for .xls)"
            except ImportError:
                 try:
                    import xlrd
                    excel_support_status = f"Partial: xlrd {xlrd.__version__ if hasattr(xlrd, '__version__') else 'version unknown'} (openpyxl missing for .xlsx)"
                 except ImportError:
                    excel_support_status = "Neither openpyxl nor xlrd found."

        info = {
            "application_version": "1.2.3 (Report Format Update)", 
            "python_version": sys.version,
            "platform_details": f"{os.name} - {sys.platform}",
            "watch_directory_configured": watch_dir_config,
            "watch_directory_exists": os.path.exists(watch_dir_config) if watch_dir_config != 'Not Set' else False,
            "upload_directory_configured": upload_dir_config,
            "upload_directory_exists": os.path.exists(upload_dir_config), 
            "monitor_is_active": APP_STATE["monitor_active"],
            "current_runtime_ai_provider": app.config.get('AI_PROVIDER'), 
            "ollama_model_setting": app.config.get('OLLAMA_MODEL'),
            "vllm_model_setting": app.config.get('VLLM_MODEL'),
            "openai_model_setting": app.config.get('OPENAI_MODEL'),
            "vllm_ssl_verify_setting": app.config.get('VLLM_VERIFY_SSL'), 
            "installed_packages_check": {
                "pandas": pd.__version__ if 'pd' in globals() and hasattr(pd, '__version__') else "Not imported/found or version unknown",
                "flask": Flask.__version__ if 'Flask' in globals() and hasattr(Flask, '__version__') else "Not imported/found or version unknown",
                "watchdog": getattr(__import__('watchdog'), '__version__', 'Installed, version not detected'),
                "requests": requests.__version__ if 'requests' in globals() and hasattr(requests, '__version__') else "Not imported/found or version unknown",
                "werkzeug": getattr(__import__('werkzeug'), '__version__', 'Installed, version not detected')
            },
            "excel_support_libraries": excel_support_status
        }
        return jsonify(info), 200
    except Exception as e:
        logger.error(f"Error in /api/system-info: {str(e)}")
        logger.exception("Detailed traceback for system-info error:")
        return jsonify({"status": "error", "message": f"Internal server error: {str(e)}"}), 500

# --- Main Entry Point ---
if __name__ == '__main__':
    try:
        update_requirements() 
    except Exception as e:
        logger.warning(f"Could not run update_requirements() on startup: {str(e)}")
    
    logger.info(f"Application starting with AI_PROVIDER set to: {app.config.get('AI_PROVIDER')}")
    logger.info(f"VLLM SSL Verification is set to: {app.config.get('VLLM_VERIFY_SSL')}")


    if app.config.get('AUTO_START_MONITORING', False):
        logger.info("AUTO_START_MONITORING is enabled. Starting file monitoring...")
        startup_monitor_thread = threading.Thread(target=start_file_monitoring, daemon=True)
        startup_monitor_thread.start()
    else:
        logger.info("AUTO_START_MONITORING is disabled. File monitoring can be started via API.")
    
    try:
        port_num = int(app.config.get('PORT', 5000))
    except ValueError:
        logger.warning(f"Invalid PORT value '{app.config.get('PORT')}' in config. Defaulting to 5000.")
        port_num = 5000

    app.run(
        host=app.config.get('HOST', '0.0.0.0'),
        port=port_num,
        debug=app.config.get('DEBUG', False),
    )

