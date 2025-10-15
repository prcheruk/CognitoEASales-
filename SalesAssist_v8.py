import oracledb as cx_Oracle
import os
import logging
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime, timedelta
import requests
import json
import base64
import re
import io
import csv
import openai
import numpy as np
from typing import Tuple, List

# Load environment variables
load_dotenv()

# --- DIAGNOSTIC: Verify .env loading ---
print("\n--- .env Loading Diagnostics (Global Scope) ---")
dot_env_path = os.path.join(os.getcwd(), '.env')
if os.path.exists(dot_env_path):
    print(f".env file found at: {dot_env_path}")
    if load_dotenv(verbose=True, override=True):
        print("load_dotenv() reported successful loading/overriding.")
    else:
        print("load_dotenv() reported no new variables loaded or no override occurred.")
else:
    print(f".env file NOT found at: {dot_env_path}. This is likely the issue.")

LOCAL_ONEDRIVE_ROOT = os.getenv("LOCAL_ONEDRIVE_ROOT")

print(f"Value of LOCAL_ONEDRIVE_ROOT from os.getenv() (Global Scope): {repr(LOCAL_ONEDRIVE_ROOT)}")
print(f"Is LOCAL_ONEDRIVE_ROOT in os.environ (Global Scope)? {'LOCAL_ONEDRIVE_ROOT' in os.environ}")
print("--- End .env Loading Diagnostics (Global Scope) ---\n")


# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


# --- Webex API Configuration ---
WEBEX_API_URL = "https://webexapis.com/v1/messages"
AUTHORIZATION_TOKEN = os.getenv("WEBEX_AUTHORIZATION_TOKEN")
WEBEX_ROOM_ID = os.getenv("WEBEX_ROOM_ID")
FALLBACK_WEBEX_EMAIL = os.getenv("FALLBACK_EMAIL", "prcheruk@cisco.com")

if not AUTHORIZATION_TOKEN:
    logging.error("WEBEX_AUTHORIZATION_TOKEN environment variable not set. Webex communication may not work.")

# --- Oracle Database Connection Details ---
DB_USERNAME = os.getenv("DB_USERNAME")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_DSN = (
    "(DESCRIPTION=(CONNECT_TIMEOUT=5)(TRANSPORT_CONNECT_TIMEOUT=3)(RETRY_COUNT=1)"
    "(ADDRESS_LIST=(LOAD_BALANCE=ON)(FAILOVER=ON)"
    "(ADDRESS=(PROTOCOL=TCP)(HOST=scan-prd-2102)(PORT=1541))"
    "(ADDRESS=(PROTOCOL=TCP)(HOST=scan-prd-2101)(PORT=1541)))"
    "(CONNECT_DATA=(SERVICE_NAME=CSFPRD_SRVC_RO.cisco.com)(SERVER=DEDICATED)))"
)

if not DB_USERNAME or not DB_PASSWORD or not DB_DSN:
    logging.error("DB_USERNAME, DB_PASSWORD, or DB_DSN environment variables not set. Exiting.")
    exit(1)

# --- Path to Oracle Instant Client ---
ORACLE_CLIENT_LIB_DIR = os.getenv("ORACLE_CLIENT_LIB_DIR")
try:
    if ORACLE_CLIENT_LIB_DIR and os.path.exists(ORACLE_CLIENT_LIB_DIR):
        try:
            cx_Oracle.init_oracle_client(lib_dir=ORACLE_CLIENT_LIB_DIR)
            logging.info(f"Initialized Oracle Client from: {ORACLE_CLIENT_LIB_DIR}")
        except cx_Oracle.Error as e:
            logging.warning(f"Failed to initialize Oracle Client from specified lib_dir '{ORACLE_CLIENT_LIB_DIR}': {e}. Attempting default initialization.")
            cx_Oracle.init_oracle_client()
            logging.info("Initialized Oracle Client using default search paths.")
    else:
        logging.warning("ORACLE_CLIENT_LIB_DIR not set or path invalid. Attempting default Oracle Client initialization.")
        cx_Oracle.init_oracle_client()
        logging.info("Initialized Oracle Client using default search paths.")
except cx_Oracle.Error as e:
    logging.error(f"FATAL: Error initializing Oracle Client (even with default search): {e}")
    logging.error("Please ensure Oracle Instant Client is correctly installed and configured, and its dependencies (like Visual C++ Redistributable) are met.")
    exit(1)

# --- Cisco AI Configuration ---
CISCO_AI_TOKEN_URL = "https://id.cisco.com/oauth2/default/v1/token"
CISCO_AI_CLIENT_ID = os.getenv("CISCO_AI_CLIENT_ID")
CISCO_AI_CLIENT_SECRET = os.getenv("CISCO_AI_CLIENT_SECRET")
CISCO_AI_APP_KEY = os.getenv("CISCO_AI_APP_KEY")
CISCO_AI_ENDPOINT = 'https://chat-ai.cisco.com'
CISCO_AI_API_VERSION = "2023-08-01-preview"
CISCO_AI_MODEL = "gpt-4.1" # Or other available model

if not all([CISCO_AI_CLIENT_ID, CISCO_AI_CLIENT_SECRET, CISCO_AI_APP_KEY]):
    logging.warning("One or more Cisco AI environment variables (CLIENT_ID, CLIENT_SECRET, APP_KEY) not set. AI summarization may not work.")


# --- Utility Functions ---

def json_serial(obj):
    """JSON serializer for objects not serializable by default json code"""
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Type {type(obj)} not serializable}}")

def clean_excel_string(text):
    """
    Removes characters that are illegal in Excel worksheets (invalid XML characters).
    These typically include control characters.
    """
    if isinstance(text, str):
        return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
    return text

def dicts_to_csv_string(data):
    """
    Converts a list of dictionaries to a CSV-formatted string.
    """
    if not data:
        return ""
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=data[0].keys())
    writer.writeheader()
    writer.writerows(data)
    return output.getvalue()

def get_cisco_ai_client():
    """
    Obtains an OAuth token from Cisco and initializes the AzureOpenAI client.
    Returns the initialized client object or None on failure.
    """
    if not all([CISCO_AI_CLIENT_ID, CISCO_AI_CLIENT_SECRET, CISCO_AI_APP_KEY]):
        logging.error("Cisco AI credentials (CLIENT_ID, CLIENT_SECRET, APP_KEY) are incomplete. Cannot initialize AI client.")
        return None
    
    try:
        base64_auth = base64.b64encode(f'{CISCO_AI_CLIENT_ID}:{CISCO_AI_CLIENT_SECRET}'.encode('utf-8')).decode('utf-8')
        payload = "grant_type=client_credentials"
        headers = {
            "Accept": "*/*",
            "Content-Type": "application/x-www-form-urlencoded",
            "Authorization": f"Basic {base64_auth}"
        }

        logging.info("Attempting to get Cisco AI OAuth token...")
        token_response = requests.request("POST", CISCO_AI_TOKEN_URL, headers=headers, data=payload)
        token_response.raise_for_status()
        access_token = token_response.json()["access_token"]
        logging.info("Successfully obtained Cisco AI OAuth token.")

        client = openai.AzureOpenAI(
            azure_endpoint=CISCO_AI_ENDPOINT,
            api_key=access_token,
            api_version=CISCO_AI_API_VERSION
        )
        return client

    except requests.exceptions.RequestException as e:
        logging.error(f"Error obtaining Cisco AI token: {e}")
        if hasattr(e, 'response') and e.response is not None:
            logging.error(f"Response status code: {e.response.status_code}")
            logging.error(f"Response content: {e.response.text}")
        return None
    except Exception as e:
        logging.error(f"An unexpected error occurred during Cisco AI client initialization: {e}")
        return None

def summarize_with_cisco_ai(ai_client, text_to_summarize, prompt_prefix="Summarize the following data and provide actionable recommendations for a sales team:"):
    """
    Sends text to Cisco AI for summarization and recommendation.
    :param ai_client: An initialized openai.AzureOpenAI client.
    :param text_to_summarize: The analysis summary string to send to AI.
    :param prompt_prefix: An optional prefix for the AI prompt to guide summarization.
    :return: A string containing the AI-generated summary, or an error message.
    """
    if not ai_client:
        logging.warning("AI client is not initialized. Cannot summarize.")
        return "AI client is not initialized. Cannot summarize."

    if not text_to_summarize or not text_to_summarize.strip():
        logging.info("No data to summarize (text_to_summarize is empty).")
        return "No analysis data provided for AI summarization."

    messages = [
        {"role": "system", "content": "You are an expert Cisco Services Sales Consultant. Your task is to analyze customer data and provide concise, actionable recommendations for Services EA scope, highlighting opportunities and justifications."},
        {"role": "user", "content": f"{prompt_prefix}\n\n{text_to_summarize}"}
    ]

    try:
        logging.info("Sending data to Cisco AI for summarization...")
        response = ai_client.chat.completions.create(
            model=CISCO_AI_MODEL,
            messages=messages,
            user=f'{{"appkey": "{CISCO_AI_APP_KEY}"}}'
        )
        summary = response.choices[0].message.content
        
        if summary is None:
            logging.warning("Cisco AI returned None for summary content. This might mean the AI found nothing to summarize or an internal issue.")
            return "AI could not generate a summary for the provided data."
        elif not summary.strip():
            logging.warning("Cisco AI returned an empty or whitespace-only summary.")
            return "AI generated an empty summary. Data might be too sparse or prompt too restrictive."

        logging.info("Successfully received summary from Cisco AI.")
        return summary
    except openai.APIError as e:
        logging.error(f"Cisco AI API Error during summarization: {e}")
        if hasattr(e, 'response') and e.response is not None:
            logging.error(f"AI API Error Response status code: {e.response.status_code}")
            logging.error(f"AI API Error Response content: {e.response.text}")
        return f"Error summarizing data with AI: {e}"
    except Exception as e:
        logging.error(f"An unexpected error occurred during AI summarization: {e}")
        return f"Error summarizing data with AI: {e}"

def send_webex_message(to_email=None, room_id=None, text_message=None, card_payload=None):
    """
    Sends a Webex message, either plain text or an Adaptive Card, to a user and/or a room.
    """
    if not AUTHORIZATION_TOKEN:
        logging.error("WEBEX_AUTHORIZATION_TOKEN not set. Skipping sending Webex message.")
        return False

    if not AUTHORIZATION_TOKEN.startswith("Bearer "):
        bearer_token = f"Bearer {AUTHORIZATION_TOKEN}"
    else:
        bearer_token = AUTHORIZATION_TOKEN
    
    headers = {"Authorization": bearer_token}
    message_payload = {}

    target_identifier_for_log = None

    if room_id:
        message_payload["roomId"] = room_id
        target_identifier_for_log = f"Room ID: {repr(room_id)}"
    elif to_email:
        message_payload["toPersonEmail"] = to_email
        target_identifier_for_log = f"Email: {repr(to_email)}"
    else:
        logging.error("No recipient (to_email or room_id) specified for Webex message.")
        return False

    if card_payload:
        headers["Content-Type"] = "application/json"
        message_payload["text"] = text_message if text_message else "AI Analysis Summary"
        message_payload["attachments"] = [card_payload]
    elif text_message:
        headers["Content-Type"] = "application/json"
        message_payload["text"] = text_message
    else:
        logging.warning("No text message or card payload provided to send a Webex message.")
        return False

    print(f"\n--- Webex Send Diagnostic ---")
    print(f"Attempting to send to: {target_identifier_for_log}")
    print(f"Type of target ID: {type(room_id) if room_id else type(to_email)}")
    print(f"Length of target ID: {len(room_id) if room_id else len(to_email)}")
    print(f"Message Payload (first 500 chars): {json.dumps(message_payload, indent=2)[:500]}...")
    print(f"Headers: {headers}")
    print(f"--- End Webex Send Diagnostic ---\n")

    try:
        logging.info(f"Attempting to send Webex message to {room_id if room_id else to_email}...")
        response = requests.post(WEBEX_API_URL, headers=headers, data=json.dumps(message_payload))
        response.raise_for_status()
        logging.info(f"Successfully sent Webex message. Message ID: {response.json().get('id', 'N/A')}")
        return True
    except requests.exceptions.RequestException as e:
        logging.error(f"Error sending Webex message: {e}")
        if hasattr(e, 'response') and e.response is not None:
            logging.error(f"Webex API Response status code: {e.response.status_code}")
            logging.error(f"Webex API Response content: {e.response.text}")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred during Webex message sending: {e}")
        return False


def execute_sql_file(sql_file_path, bind_params=None, db_username=DB_USERNAME, db_password=DB_PASSWORD, db_dsn=DB_DSN):
    """
    Executes an SQL query from a file, passing bind parameters.
    Returns a list of dictionaries, where each dictionary represents a row.
    """
    connection = None
    cursor = None
    results = []
    if bind_params is None:
        bind_params = {}

    try:
        with open(sql_file_path, 'r') as f:
            sql_content = f.read()
        logging.info(f"SQL query loaded from '{sql_file_path}'.")

        logging.info(f"Attempting to connect to Oracle database using DSN: {db_dsn}...")
        connection = cx_Oracle.connect(user=db_username, password=db_password, dsn=db_dsn)
        cursor = connection.cursor()
        logging.info("Successfully connected to Oracle database.")

        logging.info(f"Executing query with params: {bind_params}...")
        cursor.execute(sql_content, bind_params)
        logging.info(f"Query executed on DB server. Now fetching results for params: {bind_params}...")
        rows = cursor.fetchall()

        columns = [col[0].upper() for col in cursor.description]

        for row in rows:
            results.append(dict(zip(columns, row)))

        logging.info(f"Found {len(results)} records.")
        return results

    except FileNotFoundError:
        logging.error(f"Error: SQL file not found at '{sql_file_path}'.")
        return []
    except cx_Oracle.Error as e:
        error_obj, = e.args
        logging.error(f"Oracle Database Error ({error_obj.code}): {error_obj.message}")
        logging.error(f"Failed to execute query from '{sql_file_path}' with params '{bind_params}'.")
        return []
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        return []
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()
            logging.info("Database connection closed.")

def save_dataframe_to_csv(df, file_path, index=False):
    """
    Saves a pandas DataFrame to a CSV file.
    """
    try:
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        df.to_csv(file_path, index=index)
        logging.info(f"DataFrame successfully saved to CSV: {file_path}")
        return True
    except Exception as e:
        logging.error(f"Error saving DataFrame to CSV '{file_path}': {e}")
        return False

def save_dataframe_to_excel(df, file_path, sheet_name="Data", index=False):
    """
    Saves a pandas DataFrame to an Excel file.
    """
    try:
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        df.to_excel(file_path, sheet_name=sheet_name, index=index)
        logging.info(f"DataFrame successfully saved to Excel: {file_path}")
        return True
    except Exception as e:
        logging.error(f"Error saving DataFrame to Excel '{file_path}': {e}")
        return False

def load_excel_from_local_onedrive(relative_onedrive_path, file_name, sheet_name=0):
    """
    Loads an Excel file from a locally synced OneDrive folder into a pandas DataFrame.
    """
    if not LOCAL_ONEDRIVE_ROOT:
        logging.error("LOCAL_ONEDRIVE_ROOT is not configured. Cannot load Excel from OneDrive.")
        return None
    if not os.path.isdir(LOCAL_ONEDRIVE_ROOT):
        logging.error(f"Configured LOCAL_ONEDRIVE_ROOT '{LOCAL_ONEDRIVE_ROOT}' is not a valid directory.")
        return None

    if relative_onedrive_path and (relative_onedrive_path.startswith('/') or relative_onedrive_path.startswith('\\')):
        logging.warning(f"Relative OneDrive path '{relative_onedrive_path}' starts with a separator. Removing it for correct joining.")
        relative_onedrive_path = relative_onedrive_path.lstrip('/\\')

    full_local_path = os.path.join(LOCAL_ONEDRIVE_ROOT, relative_onedrive_path, file_name)

    logging.info(f"Attempting to load Excel file from: {full_local_path}")

    if not os.path.exists(full_local_path):
        logging.error(f"Error: Excel file not found at '{full_local_path}'. "
                      f"Please ensure the file exists and your OneDrive is synced.")
        return None

    try:
        df = pd.read_excel(full_local_path, sheet_name=sheet_name)
        logging.info(f"Successfully loaded '{file_name}' into DataFrame (shape: {df.shape}).")
        return df
    except Exception as e:
        print(f"CRITICAL ERROR: Failed to read Excel file '{full_local_path}'. Exception: {e}")
        logging.error(f"Error loading Excel file '{file_name}' from '{full_local_path}': {e}")
        return None

import pandas as pd
from typing import List, Tuple

def perform_ea_analysis(
    target_cav_id_int: int,
    customer_detail_by_cav_df: pd.DataFrame,  # CHANGED: Now using CAV-specific customer data
    case_history_df: pd.DataFrame,
    install_base_df: pd.DataFrame,
    sales_history_df: pd.DataFrame,
    customer_serial_ib_df: pd.DataFrame,
    eamp_proposal_quote_df: pd.DataFrame,
) -> Tuple[pd.DataFrame, str]:
    """
    Perform EA analysis using CAV-specific customer data that matches IB/case/sales data.
    Returns (final_activity_summary_df, analysis_log_string).
    """
    messages: List[str] = []

    def _upper_cols(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = df.columns.str.upper()
        return df

    def _safe_numeric(df: pd.DataFrame, cols: List[str]):
        for c in cols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        return df

    def _aggregate_ib_from_install_base(ib_df: pd.DataFrame) -> pd.DataFrame:
        """Aggregate IB counts from install_base_df using QUANTITY and COVERAGE_STATUS."""
        ib_df = _upper_cols(ib_df)
        
        group_keys = ["CAV_BU_ID", "PARTY_ID"]
        ib_count_cols = ["IB_COUNT_COVERED", "IB_COUNT_UNCOVERED", "IB_COUNT_NEVER_COVERED"]

        if not all(k in ib_df.columns for k in group_keys):
            messages.append(f"Warning: IB aggregation skipped - missing keys {group_keys}")
            return pd.DataFrame(columns=group_keys + ib_count_cols)

        if "COVERAGE_STATUS" not in ib_df.columns:
            messages.append(f"Warning: IB aggregation skipped - missing COVERAGE_STATUS")
            return pd.DataFrame(columns=group_keys + ib_count_cols)

        # Convert merge keys to consistent types
        ib_df['CAV_BU_ID'] = pd.to_numeric(ib_df['CAV_BU_ID'], errors='coerce').astype(pd.Int64Dtype())
        ib_df['PARTY_ID'] = pd.to_numeric(ib_df['PARTY_ID'], errors='coerce').astype(pd.Int64Dtype())
        
        ib_df_valid = ib_df.dropna(subset=['CAV_BU_ID', 'PARTY_ID']).copy()
        
        if ib_df_valid.empty:
            messages.append("Warning: No valid IB rows after removing null keys")
            return pd.DataFrame(columns=group_keys + ib_count_cols)

        def classify_coverage(val):
            if pd.isna(val):
                return "UNKNOWN"
            s = str(val).strip().upper()
            if s == "A":
                return "COVERED"
            if s == "I":
                return "UNCOVERED"
            if s == "N":
                return "NEVER_COVERED"
            return "UNKNOWN"

        ib_df_valid["_IB_CLASS"] = ib_df_valid["COVERAGE_STATUS"].apply(classify_coverage)
        
        # Use QUANTITY if available
        if "QUANTITY" in ib_df_valid.columns:
            ib_df_valid["_CNT"] = pd.to_numeric(ib_df_valid["QUANTITY"], errors='coerce').fillna(1)
            messages.append("DEBUG: Using QUANTITY column for IB counts")
        else:
            ib_df_valid["_CNT"] = 1
            messages.append("DEBUG: Using row count (1 per row) for IB counts")

        # Pivot to get counts per class
        pivot = ib_df_valid.pivot_table(index=group_keys, columns="_IB_CLASS", values="_CNT", aggfunc="sum", fill_value=0)
        pivot = pivot.reset_index()
        
        # FIX: Properly handle pivot result columns
        result = pivot.copy()
        
        # Ensure all classification columns exist with 0 defaults
        for col in ["COVERED", "UNCOVERED", "UNKNOWN", "NEVER_COVERED"]:
            if col not in result.columns:
                result[col] = 0
        
        # Create the final count columns
        result["IB_COUNT_COVERED"] = result["COVERED"]
        result["IB_COUNT_UNCOVERED"] = result["UNCOVERED"] + result["UNKNOWN"]
        result["IB_COUNT_NEVER_COVERED"] = result["NEVER_COVERED"]
        
        # Keep only the needed columns
        result = result[group_keys + ib_count_cols].copy()
        
        print(f"IB Aggregation: {len(result)} parties, Covered={result['IB_COUNT_COVERED'].sum():,.0f}, Uncovered={result['IB_COUNT_UNCOVERED'].sum():,.0f}, Never={result['IB_COUNT_NEVER_COVERED'].sum():,.0f}")
        
        return result

    def _aggregate_case_history(case_df: pd.DataFrame) -> pd.DataFrame:
        """Aggregate case counts from case_history_df."""
        case_df = _upper_cols(case_df)
        
        group_keys = ["CAV_BU_ID", "PARTY_ID"]
        
        if not all(k in case_df.columns for k in group_keys):
            messages.append(f"Warning: Case aggregation skipped - missing keys {group_keys}")
            return pd.DataFrame(columns=group_keys + ["CASE_COUNT"])
        
        case_df['CAV_BU_ID'] = pd.to_numeric(case_df['CAV_BU_ID'], errors='coerce').astype(pd.Int64Dtype())
        case_df['PARTY_ID'] = pd.to_numeric(case_df['PARTY_ID'], errors='coerce').astype(pd.Int64Dtype())
        
        case_df_valid = case_df.dropna(subset=['CAV_BU_ID', 'PARTY_ID']).copy()
        
        if case_df_valid.empty:
            messages.append("Warning: No valid case rows")
            return pd.DataFrame(columns=group_keys + ["CASE_COUNT"])
        
        if "CASE_NUMBER" in case_df_valid.columns:
            result = case_df_valid.groupby(group_keys)["CASE_NUMBER"].nunique().reset_index()
            result.rename(columns={"CASE_NUMBER": "CASE_COUNT"}, inplace=True)
            messages.append("DEBUG: Using CASE_NUMBER for case counts (distinct count)")
        elif "CASE_ID" in case_df_valid.columns:
            result = case_df_valid.groupby(group_keys)["CASE_ID"].nunique().reset_index()
            result.rename(columns={"CASE_ID": "CASE_COUNT"}, inplace=True)
            messages.append("DEBUG: Using CASE_ID for case counts (distinct count)")
        else:
            result = case_df_valid.groupby(group_keys).size().reset_index(name="CASE_COUNT")
            messages.append("DEBUG: Using row count for case counts")
        
        print(f"Case Aggregation: {len(result)} parties, Total cases={result['CASE_COUNT'].sum():,.0f}")
        
        return result

    def _aggregate_sales_history(sales_df: pd.DataFrame) -> pd.DataFrame:
        """Aggregate sales amounts from sales_history_df."""
        sales_df = _upper_cols(sales_df)
        
        group_keys = ["CAV_BU_ID", "PARTY_ID"]
        
        if not all(k in sales_df.columns for k in group_keys):
            messages.append(f"Warning: Sales aggregation skipped - missing keys {group_keys}")
            return pd.DataFrame(columns=group_keys + ["TOTAL_SALES"])
        
        sales_df['CAV_BU_ID'] = pd.to_numeric(sales_df['CAV_BU_ID'], errors='coerce').astype(pd.Int64Dtype())
        sales_df['PARTY_ID'] = pd.to_numeric(sales_df['PARTY_ID'], errors='coerce').astype(pd.Int64Dtype())
        
        sales_df_valid = sales_df.dropna(subset=['CAV_BU_ID', 'PARTY_ID']).copy()
        
        if sales_df_valid.empty:
            messages.append("Warning: No valid sales rows")
            return pd.DataFrame(columns=group_keys + ["TOTAL_SALES"])
        
        # Check for sales amount column - Added TOTAL_SALES_AMOUNT as first option
        amount_col = None
        for col_name in ["TOTAL_SALES_AMOUNT", "AMOUNT", "SALES_AMOUNT", "TOTAL_AMOUNT", "REVENUE", "SALES"]:
            if col_name in sales_df_valid.columns:
                amount_col = col_name
                break
        
        if amount_col is None:
            messages.append("Warning: No sales amount column found")
            result = sales_df_valid[group_keys].drop_duplicates()
            result["TOTAL_SALES"] = 0
            return result
        
        sales_df_valid[amount_col] = pd.to_numeric(sales_df_valid[amount_col], errors='coerce').fillna(0)
        result = sales_df_valid.groupby(group_keys)[amount_col].sum().reset_index()
        result.rename(columns={amount_col: "TOTAL_SALES"}, inplace=True)
        
        messages.append(f"DEBUG: Using {amount_col} column for sales amounts")
        print(f"Sales Aggregation: {len(result)} parties, Total sales=${result['TOTAL_SALES'].sum():,.2f}")
        
        return result

    # ===== MAIN PROCESSING =====
    
    # Step 1: Process CAV-specific customer base
    print("\n=== STEP 1: Processing CAV-specific customer base ===")
    base_df = _upper_cols(customer_detail_by_cav_df)
    
    if "CAV_BU_ID" not in base_df.columns or "PARTY_ID" not in base_df.columns:
        error_msg = f"Error: customer_detail_by_cav_df missing required merge keys CAV_BU_ID and/or PARTY_ID"
        print(error_msg)
        messages.append(error_msg)
        return pd.DataFrame(), "\n".join(messages)
    
    base_df['CAV_BU_ID'] = pd.to_numeric(base_df['CAV_BU_ID'], errors='coerce').astype(pd.Int64Dtype())
    base_df['PARTY_ID'] = pd.to_numeric(base_df['PARTY_ID'], errors='coerce').astype(pd.Int64Dtype())
    base_df = base_df.dropna(subset=['CAV_BU_ID', 'PARTY_ID']).copy()
    
    if base_df.empty:
        error_msg = "Error: No valid customers after removing null CAV_BU_ID/PARTY_ID"
        print(error_msg)
        messages.append(error_msg)
        return pd.DataFrame(), "\n".join(messages)
    
    # Keep customer info columns
    customer_cols = ["CAV_BU_ID", "PARTY_ID"]
    for col in ["CUSTOMER_NAME", "PARTY_NAME", "ACCOUNT_NAME", "CUSTOMER_ID", "PARTY_NUMBER"]:
        if col in base_df.columns:
            customer_cols.append(col)
    
    base_df = base_df[customer_cols].drop_duplicates(subset=["CAV_BU_ID", "PARTY_ID"])
    print(f"Customer base: {len(base_df)} unique party IDs for CAV {target_cav_id_int}")
    
    # Step 2-4: Aggregate data
    print("\n=== STEP 2-4: Aggregating IB, Cases, Sales ===")
    ib_agg = _aggregate_ib_from_install_base(install_base_df)
    case_agg = _aggregate_case_history(case_history_df)
    sales_agg = _aggregate_sales_history(sales_history_df)
    
    # Step 5: Check for PARTY_ID overlap
    print("\n=== STEP 5: PARTY_ID Overlap Check ===")
    base_party_ids = set(base_df['PARTY_ID'].dropna().astype(int))
    ib_party_ids = set(ib_agg['PARTY_ID'].dropna().astype(int)) if not ib_agg.empty else set()
    case_party_ids = set(case_agg['PARTY_ID'].dropna().astype(int)) if not case_agg.empty else set()
    sales_party_ids = set(sales_agg['PARTY_ID'].dropna().astype(int)) if not sales_agg.empty else set()
    
    print(f"Customer base: {len(base_party_ids)} party IDs")
    print(f"IB data: {len(ib_party_ids)} party IDs")
    print(f"Case data: {len(case_party_ids)} party IDs")
    print(f"Sales data: {len(sales_party_ids)} party IDs")
    print(f"Overlap (base ∩ IB): {len(base_party_ids & ib_party_ids)}")
    print(f"Overlap (base ∩ cases): {len(base_party_ids & case_party_ids)}")
    print(f"Overlap (base ∩ sales): {len(base_party_ids & sales_party_ids)}")
    
    # Step 6: Merge all aggregations
    print("\n=== STEP 6: Merging aggregations ===")
    merge_keys = ["CAV_BU_ID", "PARTY_ID"]
    
    final_df = base_df.copy()
    print(f"Starting with customer base: {len(final_df)} rows")
    
    # Merge IB
    if not ib_agg.empty:
        final_df = final_df.merge(ib_agg, on=merge_keys, how="left")
        for col in ["IB_COUNT_COVERED", "IB_COUNT_UNCOVERED", "IB_COUNT_NEVER_COVERED"]:
            if col in final_df.columns:
                final_df[col] = final_df[col].fillna(0)
        print(f"After IB merge: {len(final_df)} rows, Total IB = {final_df[['IB_COUNT_COVERED', 'IB_COUNT_UNCOVERED', 'IB_COUNT_NEVER_COVERED']].sum().sum():,.0f}")
    else:
        final_df["IB_COUNT_COVERED"] = 0
        final_df["IB_COUNT_UNCOVERED"] = 0
        final_df["IB_COUNT_NEVER_COVERED"] = 0
        messages.append("Warning: IB aggregation empty")
    
    # Merge cases
    if not case_agg.empty:
        final_df = final_df.merge(case_agg, on=merge_keys, how="left")
        if "CASE_COUNT" in final_df.columns:
            final_df["CASE_COUNT"] = final_df["CASE_COUNT"].fillna(0)
        print(f"After case merge: {len(final_df)} rows, Total cases = {final_df['CASE_COUNT'].sum():,.0f}")
    else:
        final_df["CASE_COUNT"] = 0
        messages.append("Warning: Case aggregation empty")
    
    # Merge sales
    if not sales_agg.empty:
        final_df = final_df.merge(sales_agg, on=merge_keys, how="left")
        if "TOTAL_SALES" in final_df.columns:
            final_df["TOTAL_SALES"] = final_df["TOTAL_SALES"].fillna(0)
        print(f"After sales merge: {len(final_df)} rows, Total sales = ${final_df['TOTAL_SALES'].sum():,.2f}")
    else:
        final_df["TOTAL_SALES"] = 0
        messages.append("Warning: Sales aggregation empty")
    
    # Step 7: Calculate totals
    print("\n=== STEP 7: Calculating totals ===")
    final_df["TOTAL_IB_COUNT"] = final_df["IB_COUNT_COVERED"] + final_df["IB_COUNT_UNCOVERED"] + final_df["IB_COUNT_NEVER_COVERED"]
    
    # Ensure numeric columns
    numeric_cols = ["IB_COUNT_COVERED", "IB_COUNT_UNCOVERED", "IB_COUNT_NEVER_COVERED", "TOTAL_IB_COUNT", "CASE_COUNT", "TOTAL_SALES"]
    final_df = _safe_numeric(final_df, numeric_cols)
    
    # Summary
    print(f"\n=== FINAL SUMMARY ===")
    print(f"Total parties: {len(final_df)}")
    print(f"Total IB Count: {final_df['TOTAL_IB_COUNT'].sum():,.0f}")
    print(f"  - Covered: {final_df['IB_COUNT_COVERED'].sum():,.0f}")
    print(f"  - Uncovered: {final_df['IB_COUNT_UNCOVERED'].sum():,.0f}")
    print(f"  - Never Covered: {final_df['IB_COUNT_NEVER_COVERED'].sum():,.0f}")
    print(f"Total Case Count: {final_df['CASE_COUNT'].sum():,.0f}")
    print(f"Total Sales: ${final_df['TOTAL_SALES'].sum():,.2f}")
    
    messages.append(f"✓ Analysis complete: {len(final_df)} parties | Total IB: {final_df['TOTAL_IB_COUNT'].sum():,.0f} | Cases: {final_df['CASE_COUNT'].sum():,.0f} | Sales: ${final_df['TOTAL_SALES'].sum():,.2f}")
    
    log_output = "\n".join(messages)
    return final_df, log_output

# --- Main execution logic ---
if __name__ == "__main__":
    # Define SQL file paths
    SQL_FILE_FOR_GU = "get_customers_by_gu_name.sql"
    # Commenting out SQL files for case history and install base
    # SQL_FILE_CASE_HISTORY = "get_case_history_by_cav.sql"
    # SQL_FILE_INSTALL_BASE = "get_install_base_by_cav.sql"

    # Define OneDrive paths and file names
    ONEDRIVE_SALES_HISTORY_FOLDER = "IB Analyst\\Metro Fire Dept" 
    SALES_HISTORY_EXCEL_FILE = "sales_history_data.xlsx" # Excel file name

    ONEDRIVE_PROPOSAL_INPUTS_FOLDER = "IB Analyst\\Metro Fire Dept"
    CUSTOMER_SERIAL_IB_EXCEL_FILE = "customer_serial_ib_data.xlsx"
    EAMP_PROPOSAL_QUOTE_EXCEL_FILE = "eamp_proposal_quote.xlsx"

    ONEDRIVE_CASE_INSTALL_BASE_FOLDER = "IB Analyst\\Metro Fire Dept"
    CASE_HISTORY_EXCEL_FILE = "case_history_by_cav.xlsx"
    INSTALL_BASE_EXCEL_FILE = "install_base_by_cav.xlsx"

    ONEDRIVE_CUSTOMER_DETAIL_FOLDER = "IB Analyst\\Metro Fire Dept"
    CUSTOMER_DETAIL_BY_CAV_EXCEL_FILE = "customer_detail_by_cavid.xlsx"


    # --- PROMPT TO CAPTURE GU_NAME ---
    target_gu_name = input("Enter the Global Ultimate Customer Name (e.g., 'Amazon'): ").strip()

    if not target_gu_name:
        logging.error("Global Ultimate Customer Name cannot be empty. Exiting.")
        exit(1)
    # --- END PROMPT TO CAPTURE GU_NAME ---

    # --- NEW: Define output_dir, sanitized_gu_name, timestamp early ---
    sanitized_gu_name = "".join(c for c in target_gu_name if c.isalnum() or c in (' ', '_')).replace(' ', '_')
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = "data_verification_output" # Folder to save files
    # --- END NEW ---

    print(f"\n--- Fetching Customer Records for Global Ultimate Customer Name: '{target_gu_name}' ---")
    customer_records_list = execute_sql_file(
        sql_file_path=SQL_FILE_FOR_GU,
        bind_params={'gu_name': target_gu_name}
    )

    if customer_records_list:
        print(f"Successfully retrieved {len(customer_records_list)} customer branch/subsidiary records for '{target_gu_name}'.")

        customer_df = pd.DataFrame(customer_records_list)
        customer_df.columns = customer_df.columns.str.upper()
        for col in ['CAV_ID', 'CAV_BU_ID', 'PARTY_ID']:
            if col in customer_df.columns:
                customer_df[col] = pd.to_numeric(customer_df[col], errors='coerce').astype(pd.Int64Dtype())


        print("\n--- Customer Data Loaded into DataFrame ---")
        print("\nDataFrame Info:")
        customer_df.info()

        expected_cav_id_str = 'CAV_ID'
        expected_cav_name_str = 'CAV_NAME'
        is_cav_id_in_cols = expected_cav_id_str in customer_df.columns
        is_cav_name_in_cols = expected_cav_name_str in customer_df.columns

        if not is_cav_id_in_cols or not is_cav_name_in_cols:
            logging.error(f"CAV_ID or CAV_NAME column not found in DataFrame. Cannot aggregate by CAV. Please update SQL query '{SQL_FILE_FOR_GU}'.")
            exit(1)

        required_ib_cols = ['IB_COUNT_COVERED', 'IB_COUNT_UNCOVERED', 'IB_COUNT_NEVER_COVERED']
        for col in required_ib_cols:
            if col not in customer_df.columns:
                logging.warning(f"Column '{col}' not found in DataFrame. Please update SQL query '{SQL_FILE_FOR_GU}'. Setting to 0 for aggregation.")
                customer_df[col] = 0
            else:
                customer_df[col] = pd.to_numeric(customer_df[col], errors='coerce').fillna(0)

        cav_aggregated_df = customer_df.groupby(['CAV_ID', 'CAV_NAME'])[[
            'IB_COUNT_COVERED', 'IB_COUNT_UNCOVERED', 'IB_COUNT_NEVER_COVERED'
        ]].sum().reset_index()

        print(f"\n--- Aggregated Customer Data by CAV for '{target_gu_name}' ---")
        print(cav_aggregated_df.to_string())

        available_cav_ids = cav_aggregated_df['CAV_ID'].unique().tolist()
        if not available_cav_ids:
            logging.info("No CAV IDs found for further analysis.")
            exit(0)

        print(f"\nAvailable CAV IDs for further analysis: {', '.join(map(str, available_cav_ids))}")
        target_cav_id = input("Enter the CAV ID for detailed analysis: ").strip()

        if not target_cav_id:
            logging.error("CAV ID cannot be empty. Exiting.")
            exit(1)
        if str(target_cav_id) not in map(str, available_cav_ids):
            logging.error(f"Invalid CAV ID '{target_cav_id}'. Please select from available IDs.")
            exit(1)
        try:
            target_cav_id_int = int(target_cav_id)
        except ValueError:
            logging.error(f"CAV ID '{target_cav_id}' is not a valid number. Exiting.")
            exit(1)


        print(f"\n--- Performing Detailed Analysis for CAV ID: {target_cav_id} ---")

        # --- Load Case History from OneDrive Excel ---
        print("\nLoading Case History from OneDrive Excel...")
        case_history_df = load_excel_from_local_onedrive(
            relative_onedrive_path=ONEDRIVE_CASE_INSTALL_BASE_FOLDER,
            file_name=CASE_HISTORY_EXCEL_FILE
        )
        if case_history_df is None: case_history_df = pd.DataFrame()
        else:
            case_history_df.columns = case_history_df.columns.str.upper()
            for col in ['CAV_ID', 'CAV_BU_ID', 'PARTY_ID']:
                if col in case_history_df.columns:
                    case_history_df[col] = pd.to_numeric(case_history_df[col], errors='coerce').astype(pd.Int64Dtype())
        if not case_history_df.empty and 'CAV_ID' in case_history_df.columns:
            case_history_df = case_history_df[case_history_df['CAV_ID'] == target_cav_id_int]
        print(f"Loaded {len(case_history_df)} case history records (filtered for CAV ID {target_cav_id_int}).")


        # --- Load Install Base from OneDrive Excel ---
        print("\nLoading Install Base Data from OneDrive Excel...")
        install_base_df = load_excel_from_local_onedrive(
            relative_onedrive_path=ONEDRIVE_CASE_INSTALL_BASE_FOLDER,
            file_name=INSTALL_BASE_EXCEL_FILE
        )
        if install_base_df is None: install_base_df = pd.DataFrame()
        else:
            install_base_df.columns = install_base_df.columns.str.upper()
            for col in ['CAV_ID', 'CAV_BU_ID', 'PARTY_ID']:
                if col in install_base_df.columns:
                    install_base_df[col] = pd.to_numeric(install_base_df[col], errors='coerce').astype(pd.Int64Dtype())
        if not install_base_df.empty and 'CAV_ID' in install_base_df.columns:
            install_base_df = install_base_df[install_base_df['CAV_ID'] == target_cav_id_int]
        print(f"Loaded {len(install_base_df)} install base records (filtered for CAV ID {target_cav_id_int}).")


        # --- Read Sales History from Excel in OneDrive ---
        print("\nLoading Sales History from OneDrive Excel...")
        sales_history_df = load_excel_from_local_onedrive(
            relative_onedrive_path=ONEDRIVE_SALES_HISTORY_FOLDER,
            file_name=SALES_HISTORY_EXCEL_FILE
        )
        if sales_history_df is None: sales_history_df = pd.DataFrame()
        else:
            sales_history_df.columns = sales_history_df.columns.str.upper()
            for col in ['CAV_ID', 'CAV_BU_ID', 'PARTY_ID']: # Assuming these might be in sales_history_df
                if col in sales_history_df.columns:
                    sales_history_df[col] = pd.to_numeric(sales_history_df[col], errors='coerce').astype(pd.Int64Dtype())
            # <<< END FIX >>>
        if not sales_history_df.empty and 'CAV_ID' in sales_history_df.columns:
            sales_history_df = sales_history_df[sales_history_df['CAV_ID'] == target_cav_id_int]
        print(f"Loaded {len(sales_history_df)} sales history records (filtered for CAV ID {target_cav_id_int}).")


        # --- Load Customer-Provided Serial/Install Base Data from Excel ---
        print("\nLoading Customer-Provided Serial/Install Base Data from OneDrive Excel...")
        customer_serial_ib_df = load_excel_from_local_onedrive(
            relative_onedrive_path=ONEDRIVE_PROPOSAL_INPUTS_FOLDER,
            file_name=CUSTOMER_SERIAL_IB_EXCEL_FILE
        )
        if customer_serial_ib_df is None: customer_serial_ib_df = pd.DataFrame()
        else:
            customer_serial_ib_df.columns = customer_serial_ib_df.columns.str.upper()
            for col in ['CAV_ID', 'CAV_BU_ID', 'PARTY_ID']: # Assuming these might be in customer_serial_ib_df
                if col in customer_serial_ib_df.columns:
                    customer_serial_ib_df[col] = pd.to_numeric(customer_serial_ib_df[col], errors='coerce').astype(pd.Int64Dtype())
            # <<< END FIX >>>
        if not customer_serial_ib_df.empty and 'CAV_ID' in customer_serial_ib_df.columns:
            customer_serial_ib_df = customer_serial_ib_df[customer_serial_ib_df['CAV_ID'] == target_cav_id_int]
        print(f"Loaded {len(customer_serial_ib_df)} customer-provided serial/IB records (filtered for CAV ID {target_cav_id_int}).")


        # --- Load EAMP Proposal Quote Data from Excel ---
        print("\nLoading EAMP Proposal Quote Data from OneDrive Excel...")
        eamp_proposal_quote_df = load_excel_from_local_onedrive(
            relative_onedrive_path=ONEDRIVE_PROPOSAL_INPUTS_FOLDER,
            file_name=EAMP_PROPOSAL_QUOTE_EXCEL_FILE
        )
        if eamp_proposal_quote_df is None: eamp_proposal_quote_df = pd.DataFrame()
        else:
            eamp_proposal_quote_df.columns = eamp_proposal_quote_df.columns.str.upper()
            for col in ['CAV_ID', 'CAV_BU_ID', 'PARTY_ID']: # Assuming these might be in eamp_proposal_quote_df
                if col in eamp_proposal_quote_df.columns:
                    eamp_proposal_quote_df[col] = pd.to_numeric(eamp_proposal_quote_df[col], errors='coerce').astype(pd.Int64Dtype())
            # <<< END FIX >>>
        if not eamp_proposal_quote_df.empty and 'CAV_ID' in eamp_proposal_quote_df.columns:
            eamp_proposal_quote_df = eamp_proposal_quote_df[eamp_proposal_quote_df['CAV_ID'] == target_cav_id_int]
        print(f"Loaded {len(eamp_proposal_quote_df)} EAMP proposal quote records (filtered for CAV ID {target_cav_id_int}).")


        # --- Load Customer Hierarchy Data by CAV from Excel ---
        print("\nLoading Customer Hierarchy Data by CAV from OneDrive Excel...")
        customer_detail_by_cav_df = load_excel_from_local_onedrive(
            relative_onedrive_path=ONEDRIVE_CUSTOMER_DETAIL_FOLDER,
            file_name=CUSTOMER_DETAIL_BY_CAV_EXCEL_FILE
        )
        if customer_detail_by_cav_df is None: customer_detail_by_cav_df = pd.DataFrame()
        else:
            customer_detail_by_cav_df.columns = customer_detail_by_cav_df.columns.str.upper()
            for col in ['CAV_ID', 'CAV_BU_ID', 'PARTY_ID']:
                if col in customer_detail_by_cav_df.columns:
                    customer_detail_by_cav_df[col] = pd.to_numeric(customer_detail_by_cav_df[col], errors='coerce').astype(pd.Int64Dtype())
            # <<< END FIX >>>
        if not customer_detail_by_cav_df.empty and 'CAV_ID' in customer_detail_by_cav_df.columns:
            customer_detail_by_cav_df = customer_detail_by_cav_df[customer_detail_by_cav_df['CAV_ID'] == target_cav_id_int]
        print(f"Loaded {len(customer_detail_by_cav_df)} customer hierarchy records (filtered for CAV ID {target_cav_id_int}).")


        print("\n--- All requested data sets for detailed analysis are loaded. ---")
        
        # --- Perform Analysis ---
        final_summary_df_for_debug, ea_analysis_summary_str = perform_ea_analysis(
            target_cav_id_int,
            customer_detail_by_cav_df,
            case_history_df,
            install_base_df,
            sales_history_df,
            customer_serial_ib_df,
            eamp_proposal_quote_df
            )
        print("\n" + ea_analysis_summary_str) # Print the string part of the analysis

        # --- Save final_summary_df to Excel for verification ---
        if not final_summary_df_for_debug.empty:
            debug_summary_filename = f"{sanitized_gu_name}_CAV_{target_cav_id}_FinalSummary_DEBUG_{timestamp}.xlsx"
            debug_summary_filepath = os.path.join(output_dir, debug_summary_filename)
            save_dataframe_to_excel(final_summary_df_for_debug, debug_summary_filepath, sheet_name="Final Summary")
            print(f"\nDEBUG: Final summary DataFrame saved to: {debug_summary_filepath}")
        else:
            print("\nDEBUG: Final summary DataFrame is empty, not saving to Excel.")


        # --- AI Summarization and Webex Notification ---
        # COMMENTED OUT FOR FOCUSED DEBUGGING
        ai_input_summary_df = pd.DataFrame({'AI_Input_Summary': [ea_analysis_summary_str]})
        ai_input_summary_filename = f"{sanitized_gu_name}_CAV_{target_cav_id}_AI_Input_Summary_{timestamp}.xlsx"
        ai_input_summary_filepath = os.path.join(output_dir, ai_input_summary_filename)
        save_dataframe_to_excel(ai_input_summary_df, ai_input_summary_filepath, sheet_name="AI Input Summary")
        print(f"\nAI input summary saved to: {ai_input_summary_filepath}")

        cisco_ai_client = get_cisco_ai_client()
        if cisco_ai_client:
            logging.info("Sending analysis summary to Cisco AI for further insights...")
            ai_generated_summary = summarize_with_cisco_ai(
                ai_client=cisco_ai_client,
                text_to_summarize=ea_analysis_summary_str,
                prompt_prefix=f"Here is a detailed analysis for CAV ID {target_cav_id_int} under Global Ultimate Customer '{target_gu_name}'. Please provide a concise, actionable summary for a sales team, highlighting key opportunities for Services EA scope, potential risks, and recommended next steps. Focus on BUs/Parties with high uncovered IB, critical cases, or significant sales, and address serial number reconciliation findings. Also, comment on BUs/Parties with no activity."
            )
            print("\n--- AI-Generated Summary for Sales Team ---")
            print(ai_generated_summary)
            if AUTHORIZATION_TOKEN:
                logging.info("Sending AI-generated summary to Webex...")
                send_webex_message(
                    room_id=WEBEX_ROOM_ID,
                    to_email=FALLBACK_WEBEX_EMAIL,
                    text_message=f"AI-Generated Services EA Analysis for {target_gu_name} (CAV ID: {target_cav_id_int})\n\n{ai_generated_summary}"
                )
            else:
                logging.warning("Webex Authorization Token not set. Skipping Webex notification.")
        else:
            logging.warning("Cisco AI client not initialized. Skipping AI summarization and Webex notification.")


    else:
        print(f"No customer records found for '{target_gu_name}' or an error occurred during retrieval.")

    print("\n--- Script finished ---")