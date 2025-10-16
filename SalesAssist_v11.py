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
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor

# Load environment variables
load_dotenv()

# --- DIAGNOSTIC: Verify .env loading ---
print("\n--- .env Loading Diagnostics (Global Scope) ---")
dot_env_path = os.path.join(os.getcwd(), ".env")
if os.path.exists(dot_env_path):
    print(f".env file found at: {dot_env_path}")
    if load_dotenv(verbose=True, override=True):
        print("load_dotenv() reported successful loading/overriding.")
    else:
        print("load_dotenv() reported no new variables loaded or no override occurred.")
else:
    print(f".env file NOT found at: {dot_env_path}. This is likely the issue.")

LOCAL_ONEDRIVE_ROOT = os.getenv("LOCAL_ONEDRIVE_ROOT")

print(
    f"Value of LOCAL_ONEDRIVE_ROOT from os.getenv() (Global Scope): {repr(LOCAL_ONEDRIVE_ROOT)}"
)
print(
    f"Is LOCAL_ONEDRIVE_ROOT in os.environ (Global Scope)? {'LOCAL_ONEDRIVE_ROOT' in os.environ}"
)
print("--- End .env Loading Diagnostics (Global Scope) ---\n")

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# --- Webex API Configuration ---
WEBEX_API_URL = "https://webexapis.com/v1/messages"
AUTHORIZATION_TOKEN = os.getenv("WEBEX_AUTHORIZATION_TOKEN")
WEBEX_ROOM_ID = os.getenv("WEBEX_ROOM_ID")
FALLBACK_WEBEX_EMAIL = os.getenv("FALLBACK_EMAIL", "prcheruk@cisco.com")

if not AUTHORIZATION_TOKEN:
    logging.error(
        "WEBEX_AUTHORIZATION_TOKEN environment variable not set. Webex communication may not work."
    )

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
    logging.error(
        "DB_USERNAME, DB_PASSWORD, or DB_DSN environment variables not set. Exiting."
    )
    exit(1)

# --- Path to Oracle Instant Client ---
ORACLE_CLIENT_LIB_DIR = os.getenv("ORACLE_CLIENT_LIB_DIR")
try:
    if ORACLE_CLIENT_LIB_DIR and os.path.exists(ORACLE_CLIENT_LIB_DIR):
        try:
            cx_Oracle.init_oracle_client(lib_dir=ORACLE_CLIENT_LIB_DIR)
            logging.info(f"Initialized Oracle Client from: {ORACLE_CLIENT_LIB_DIR}")
        except cx_Oracle.Error as e:
            logging.warning(
                f"Failed to initialize Oracle Client from specified lib_dir '{ORACLE_CLIENT_LIB_DIR}': {e}. Attempting default initialization."
            )
            cx_Oracle.init_oracle_client()
            logging.info("Initialized Oracle Client using default search paths.")
    else:
        logging.warning(
            "ORACLE_CLIENT_LIB_DIR not set or path invalid. Attempting default Oracle Client initialization."
        )
        cx_Oracle.init_oracle_client()
        logging.info("Initialized Oracle Client using default search paths.")
except cx_Oracle.Error as e:
    logging.error(
        f"FATAL: Error initializing Oracle Client (even with default search): {e}"
    )
    logging.error(
        "Please ensure Oracle Instant Client is correctly installed and configured, and its dependencies (like Visual C++ Redistributable) are met."
    )
    exit(1)

# --- Cisco AI Configuration ---
CISCO_AI_TOKEN_URL = "https://id.cisco.com/oauth2/default/v1/token"
CISCO_AI_CLIENT_ID = os.getenv("CISCO_AI_CLIENT_ID")
CISCO_AI_CLIENT_SECRET = os.getenv("CISCO_AI_CLIENT_SECRET")
CISCO_AI_APP_KEY = os.getenv("CISCO_AI_APP_KEY")
CISCO_AI_ENDPOINT = "https://chat-ai.cisco.com"
CISCO_AI_API_VERSION = "2023-08-01-preview"
CISCO_AI_MODEL = "gpt-4.1"

if not all([CISCO_AI_CLIENT_ID, CISCO_AI_CLIENT_SECRET, CISCO_AI_APP_KEY]):
    logging.warning(
        "One or more Cisco AI environment variables (CLIENT_ID, CLIENT_SECRET, APP_KEY) not set. AI summarization may not work."
    )

# --- Utility Functions ---


def json_serial(obj):
    """JSON serializer for objects not serializable by default json code"""
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Type {type(obj)} not serializable}}")


def clean_excel_string(text):
    """Removes characters that are illegal in Excel worksheets (invalid XML characters)."""
    if isinstance(text, str):
        return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    return text


def dicts_to_csv_string(data):
    """Converts a list of dictionaries to a CSV-formatted string."""
    if not data:
        return ""
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=data[0].keys())
    writer.writeheader()
    writer.writerows(data)
    return output.getvalue()


def get_cisco_ai_client():
    """Obtains an OAuth token from Cisco and initializes the AzureOpenAI client."""
    if not all([CISCO_AI_CLIENT_ID, CISCO_AI_CLIENT_SECRET, CISCO_AI_APP_KEY]):
        logging.error(
            "Cisco AI credentials (CLIENT_ID, CLIENT_SECRET, APP_KEY) are incomplete. Cannot initialize AI client."
        )
        return None

    try:
        base64_auth = base64.b64encode(
            f"{CISCO_AI_CLIENT_ID}:{CISCO_AI_CLIENT_SECRET}".encode("utf-8")
        ).decode("utf-8")
        payload = "grant_type=client_credentials"
        headers = {
            "Accept": "*/*",
            "Content-Type": "application/x-www-form-urlencoded",
            "Authorization": f"Basic {base64_auth}",
        }

        logging.info("Attempting to get Cisco AI OAuth token...")
        token_response = requests.request(
            "POST", CISCO_AI_TOKEN_URL, headers=headers, data=payload
        )
        token_response.raise_for_status()
        access_token = token_response.json()["access_token"]
        logging.info("Successfully obtained Cisco AI OAuth token.")

        client = openai.AzureOpenAI(
            azure_endpoint=CISCO_AI_ENDPOINT,
            api_key=access_token,
            api_version=CISCO_AI_API_VERSION,
        )
        return client

    except requests.exceptions.RequestException as e:
        logging.error(f"Error obtaining Cisco AI token: {e}")
        if hasattr(e, "response") and e.response is not None:
            logging.error(f"Response status code: {e.response.status_code}")
            logging.error(f"Response content: {e.response.text}")
        return None
    except Exception as e:
        logging.error(
            f"An unexpected error occurred during Cisco AI client initialization: {e}"
        )
        return None


def summarize_with_cisco_ai(
    ai_client,
    text_to_summarize,
    prompt_prefix="Summarize the following data and provide actionable recommendations for a sales team:",
):
    """Sends text to Cisco AI for summarization and recommendation."""
    if not ai_client:
        logging.warning("AI client is not initialized. Cannot summarize.")
        return "AI client is not initialized. Cannot summarize."

    if not text_to_summarize or not text_to_summarize.strip():
        logging.info("No data to summarize (text_to_summarize is empty).")
        return "No analysis data provided for AI summarization."

    messages = [
        {
            "role": "system",
            "content": "You are an expert Cisco Services Sales Consultant. Your task is to analyze customer data and provide concise, actionable recommendations for Services EA scope, highlighting opportunities and justifications.",
        },
        {"role": "user", "content": f"{prompt_prefix}\n\n{text_to_summarize}"},
    ]

    try:
        logging.info("Sending data to Cisco AI for summarization...")
        response = ai_client.chat.completions.create(
            model=CISCO_AI_MODEL,
            messages=messages,
            user=f'{{"appkey": "{CISCO_AI_APP_KEY}"}}',
        )
        summary = response.choices[0].message.content

        if summary is None:
            logging.warning("Cisco AI returned None for summary content.")
            return "AI could not generate a summary for the provided data."
        elif not summary.strip():
            logging.warning("Cisco AI returned an empty or whitespace-only summary.")
            return "AI generated an empty summary. Data might be too sparse or prompt too restrictive."

        logging.info("Successfully received summary from Cisco AI.")
        return summary
    except openai.APIError as e:
        logging.error(f"Cisco AI API Error during summarization: {e}")
        if hasattr(e, "response") and e.response is not None:
            logging.error(
                f"AI API Error Response status code: {e.response.status_code}"
            )
            logging.error(f"AI API Error Response content: {e.response.text}")
        return f"Error summarizing data with AI: {e}"
    except Exception as e:
        logging.error(f"An unexpected error occurred during AI summarization: {e}")
        return f"Error summarizing data with AI: {e}"


def send_webex_message(
    to_email=None, room_id=None, text_message=None, card_payload=None
):
    """Sends a Webex message, either plain text or an Adaptive Card, to a user and/or a room."""
    if not AUTHORIZATION_TOKEN:
        logging.error(
            "WEBEX_AUTHORIZATION_TOKEN not set. Skipping sending Webex message."
        )
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
        message_payload["text"] = (
            text_message if text_message else "AI Analysis Summary"
        )
        message_payload["attachments"] = [card_payload]
    elif text_message:
        headers["Content-Type"] = "application/json"
        message_payload["text"] = text_message
    else:
        logging.warning(
            "No text message or card payload provided to send a Webex message."
        )
        return False

    print(f"\n--- Webex Send Diagnostic ---")
    print(f"Attempting to send to: {target_identifier_for_log}")
    print(f"Type of target ID: {type(room_id) if room_id else type(to_email)}")
    print(f"Length of target ID: {len(room_id) if room_id else len(to_email)}")
    print(
        f"Message Payload (first 500 chars): {json.dumps(message_payload, indent=2)[:500]}..."
    )
    print(f"Headers: {headers}")
    print(f"--- End Webex Send Diagnostic ---\n")

    try:
        logging.info(
            f"Attempting to send Webex message to {room_id if room_id else to_email}..."
        )
        response = requests.post(
            WEBEX_API_URL, headers=headers, data=json.dumps(message_payload)
        )
        response.raise_for_status()
        logging.info(
            f"Successfully sent Webex message. Message ID: {response.json().get('id', 'N/A')}"
        )
        return True
    except requests.exceptions.RequestException as e:
        logging.error(f"Error sending Webex message: {e}")
        if hasattr(e, "response") and e.response is not None:
            logging.error(f"Webex API Response status code: {e.response.status_code}")
            logging.error(f"Webex API Response content: {e.response.text}")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred during Webex message sending: {e}")
        return False


def execute_sql_file(
    sql_file_path,
    bind_params=None,
    db_username=DB_USERNAME,
    db_password=DB_PASSWORD,
    db_dsn=DB_DSN,
):
    """Executes an SQL query from a file, passing bind parameters."""
    connection = None
    cursor = None
    results = []
    if bind_params is None:
        bind_params = {}

    try:
        with open(sql_file_path, "r") as f:
            sql_content = f.read()
        logging.info(f"SQL query loaded from '{sql_file_path}'.")

        logging.info(f"Attempting to connect to Oracle database using DSN: {db_dsn}...")
        connection = cx_Oracle.connect(
            user=db_username, password=db_password, dsn=db_dsn
        )
        cursor = connection.cursor()
        logging.info("Successfully connected to Oracle database.")

        logging.info(f"Executing query with params: {bind_params}...")
        cursor.execute(sql_content, bind_params)
        logging.info(
            f"Query executed on DB server. Now fetching results for params: {bind_params}..."
        )
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
        (error_obj,) = e.args
        logging.error(f"Oracle Database Error ({error_obj.code}): {error_obj.message}")
        logging.error(
            f"Failed to execute query from '{sql_file_path}' with params '{bind_params}'."
        )
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
    """Saves a pandas DataFrame to a CSV file."""
    try:
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        df.to_csv(file_path, index=index)
        logging.info(f"DataFrame successfully saved to CSV: {file_path}")
        return True
    except Exception as e:
        logging.error(f"Error saving DataFrame to CSV '{file_path}': {e}")
        return False


def save_dataframe_to_excel(df, file_path, sheet_name="Data", index=False):
    """Saves a pandas DataFrame to an Excel file."""
    try:
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        df.to_excel(file_path, sheet_name=sheet_name, index=index)
        logging.info(f"DataFrame successfully saved to Excel: {file_path}")
        return True
    except Exception as e:
        logging.error(f"Error saving DataFrame to Excel '{file_path}': {e}")
        return False


def load_excel_from_local_onedrive(relative_onedrive_path, file_name, sheet_name=0):
    """Loads an Excel file from a locally synced OneDrive folder into a pandas DataFrame."""
    if not LOCAL_ONEDRIVE_ROOT:
        logging.error(
            "LOCAL_ONEDRIVE_ROOT is not configured. Cannot load Excel from OneDrive."
        )
        return None
    if not os.path.isdir(LOCAL_ONEDRIVE_ROOT):
        logging.error(
            f"Configured LOCAL_ONEDRIVE_ROOT '{LOCAL_ONEDRIVE_ROOT}' is not a valid directory."
        )
        return None

    if relative_onedrive_path and (
        relative_onedrive_path.startswith("/")
        or relative_onedrive_path.startswith("\\")
    ):
        logging.warning(
            f"Relative OneDrive path '{relative_onedrive_path}' starts with a separator. Removing it for correct joining."
        )
        relative_onedrive_path = relative_onedrive_path.lstrip("/\\")

    full_local_path = os.path.join(
        LOCAL_ONEDRIVE_ROOT, relative_onedrive_path, file_name
    )

    logging.info(f"Attempting to load Excel file from: {full_local_path}")

    if not os.path.exists(full_local_path):
        logging.error(
            f"Error: Excel file not found at '{full_local_path}'. "
            f"Please ensure the file exists and your OneDrive is synced."
        )
        return None

    try:
        df = pd.read_excel(full_local_path, sheet_name=sheet_name)
        logging.info(
            f"Successfully loaded '{file_name}' into DataFrame (shape: {df.shape})."
        )
        return df
    except Exception as e:
        print(
            f"CRITICAL ERROR: Failed to read Excel file '{full_local_path}'. Exception: {e}"
        )
        logging.error(
            f"Error loading Excel file '{file_name}' from '{full_local_path}': {e}"
        )
        return None


def send_webex_message_with_files(
    to_email=None,
    room_id=None,
    text_message=None,
    file_paths=None,
    authorization_token=None,
):
    """Sends a Webex message with file attachments."""
    if not authorization_token:
        logging.error("WEBEX_AUTHORIZATION_TOKEN not set. Cannot send message.")
        return False

    if not authorization_token.startswith("Bearer "):
        bearer_token = f"Bearer {authorization_token}"
    else:
        bearer_token = authorization_token

    WEBEX_API_URL = "https://webexapis.com/v1/messages"

    form_data = {}

    if room_id:
        form_data["roomId"] = room_id
    elif to_email:
        form_data["toPersonEmail"] = to_email
    else:
        logging.error("No recipient specified for Webex message.")
        return False

    if text_message:
        form_data["markdown"] = text_message

    files_to_upload = []
    if file_paths:
        for file_path in file_paths:
            try:
                file_obj = open(file_path, "rb")
                file_name = file_path.split("/")[-1].split("\\")[-1]
                files_to_upload.append(
                    ("files", (file_name, file_obj, "application/octet-stream"))
                )
            except Exception as e:
                logging.error(f"Error opening file {file_path}: {e}")

    try:
        headers = {"Authorization": bearer_token}

        logging.info(f"Sending Webex message with {len(files_to_upload)} file(s)...")
        response = requests.post(
            WEBEX_API_URL,
            headers=headers,
            data=form_data,
            files=files_to_upload if files_to_upload else None,
        )
        response.raise_for_status()

        logging.info(
            f"Successfully sent Webex message. Message ID: {response.json().get('id', 'N/A')}"
        )

        for _, (_, file_obj, _) in files_to_upload:
            file_obj.close()

        return True

    except requests.exceptions.RequestException as e:
        logging.error(f"Error sending Webex message: {e}")
        if hasattr(e, "response") and e.response is not None:
            logging.error(f"Response status: {e.response.status_code}")
            logging.error(f"Response content: {e.response.text}")

        for _, (_, file_obj, _) in files_to_upload:
            file_obj.close()

        return False


def prepare_data_for_openai(
    final_df: pd.DataFrame,
    serial_recon_df: pd.DataFrame = None,
    serial_summary: str = None,
) -> str:
    """Convert the final analysis DataFrame and serial reconciliation into a structured text format for OpenAI."""
    summary_parts = []

    # FIXED: De-duplicate by PARTY_ID to avoid counting the same party multiple times
    #   'IB_COUNT_COVERED': 'first',
    #    'IB_COUNT_UNCOVERED': 'first',
    #    'IB_COUNT_NEVER_COVERED': 'first',
    #    'TOTAL_IB_COUNT': 'first',
    #    'CASE_COUNT': 'first',
    #    'TOTAL_SALES': 'first',
    #    'CAV_BU_NAME': 'first' if 'CAV_BU_NAME' in final_df.columns else 'max',
    #    'PARTY_NAME': 'first' if 'PARTY_NAME' in final_df.columns else 'max',
    #    'BE_BREAKDOWN': 'first' if 'BE_BREAKDOWN' in final_df.columns else 'max',
    #    'SUBBE_BREAKDOWN': 'first' if 'SUBBE_BREAKDOWN' in final_df.columns else 'max'
    # }).reset_index()

    agg_dict = {
        "IB_COUNT_COVERED": "first",
        "IB_COUNT_UNCOVERED": "first",
        "IB_COUNT_NEVER_COVERED": "first",
        "TOTAL_IB_COUNT": "first",
        "CASE_COUNT": "first",
        "TOTAL_SALES": "first",
    }

    # Add optional columns if they exist
    if "CAV_BU_NAME" in final_df.columns:
        agg_dict["CAV_BU_NAME"] = "first"
    if "PARTY_NAME" in final_df.columns:
        agg_dict["PARTY_NAME"] = "first"
    if "BE_BREAKDOWN" in final_df.columns:
        agg_dict["BE_BREAKDOWN"] = "first"
    if "SUBBE_BREAKDOWN" in final_df.columns:
        agg_dict["SUBBE_BREAKDOWN"] = "first"

    # Add EOL/EOS columns if they exist
    for col in [
        "EOL_PASSED",
        "EOL_WITHIN_1YR",
        "EOL_FUTURE",
        "EOL_UNKNOWN",
        "EOS_PASSED",
        "EOS_WITHIN_1YR",
        "EOS_FUTURE",
        "EOS_UNKNOWN",
    ]:
        if col in final_df.columns:
            agg_dict[col] = "first"

    party_level_df = (
        final_df.groupby(["CAV_BU_ID", "PARTY_ID"]).agg(agg_dict).reset_index()
    )
    summary_parts.append("=== OVERALL STATISTICS ===")

    # Count unique customer business units
    unique_customer_bu_count = 0
    if "CAV_BU_NAME" in party_level_df.columns:
        unique_customer_bu_count = party_level_df["CAV_BU_NAME"].nunique()
    elif "CAV_BU_ID" in party_level_df.columns:
        unique_customer_bu_count = party_level_df["CAV_BU_ID"].nunique()
    else:
        unique_customer_bu_count = len(party_level_df)

    summary_parts.append(f"Total Customer Business Units: {unique_customer_bu_count}")
    summary_parts.append(f"Total Parties (Locations) Analyzed: {len(party_level_df)}")

    # Use party_level_df for all calculations to avoid double-counting
    total_ib = party_level_df.get("TOTAL_IB_COUNT", pd.Series([0])).sum()
    ib_covered = party_level_df.get("IB_COUNT_COVERED", pd.Series([0])).sum()
    ib_uncovered = party_level_df.get("IB_COUNT_UNCOVERED", pd.Series([0])).sum()
    ib_never_covered = party_level_df.get(
        "IB_COUNT_NEVER_COVERED", pd.Series([0])
    ).sum()
    total_cases = party_level_df.get("CASE_COUNT", pd.Series([0])).sum()
    total_sales = party_level_df.get("TOTAL_SALES", pd.Series([0])).sum()

    # Count unique Cisco Business Entities (U_BE)
    unique_be_count = 0
    unique_subbe_count = 0
    all_be_names = set()
    all_subbe_names = set()

    if "BE_BREAKDOWN" in party_level_df.columns:
        for idx, row in party_level_df.iterrows():
            if (
                pd.notna(row.get("BE_BREAKDOWN"))
                and str(row.get("BE_BREAKDOWN")).strip()
            ):
                be_str = str(row["BE_BREAKDOWN"])
                for item in be_str.split(";"):
                    if ":" in item:
                        be_name = item.split(":", 1)[0].strip()
                        if be_name:
                            all_be_names.add(be_name)
        unique_be_count = len(all_be_names)

    if "SUBBE_BREAKDOWN" in party_level_df.columns:
        for idx, row in party_level_df.iterrows():
            if (
                pd.notna(row.get("SUBBE_BREAKDOWN"))
                and str(row.get("SUBBE_BREAKDOWN")).strip()
            ):
                subbe_str = str(row["SUBBE_BREAKDOWN"])
                for item in subbe_str.split(";"):
                    if ":" in item:
                        subbe_name = item.split(":", 1)[0].strip()
                        if subbe_name:
                            all_subbe_names.add(subbe_name)
        unique_subbe_count = len(all_subbe_names)

    summary_parts.append(f"Total Cisco Business Entities (U_BE): {unique_be_count}")
    summary_parts.append(
        f"Total Cisco Sub-Business Entities (U_SUBBE): {unique_subbe_count}"
    )
    summary_parts.append(f"Total Install Base: {total_ib:,.0f}")
    summary_parts.append(f"  - Covered: {ib_covered:,.0f}")
    summary_parts.append(f"  - Uncovered: {ib_uncovered:,.0f}")
    summary_parts.append(f"  - Never Covered: {ib_never_covered:,.0f}")
    summary_parts.append(f"Total Support Cases: {total_cases:,.0f}")
    summary_parts.append(f"Total Sales Revenue: ${total_sales:,.2f}")

    # List the customer business units if count is reasonable
    if unique_customer_bu_count > 0 and unique_customer_bu_count <= 50:
        summary_parts.append(f"\nCustomer Business Units ({unique_customer_bu_count}):")
        if "CAV_BU_NAME" in party_level_df.columns:
            unique_bus = party_level_df["CAV_BU_NAME"].unique()
            for bu_name in sorted(unique_bus):
                if pd.notna(bu_name):
                    summary_parts.append(f"  - {bu_name}")

    # Continue with U_BE breakdown section...
    if (
        "BE_BREAKDOWN" in party_level_df.columns
        or "SUBBE_BREAKDOWN" in party_level_df.columns
    ):
        summary_parts.append("\n=== CISCO BUSINESS ENTITY (U_BE) DISTRIBUTION ===")
        summary_parts.append(
            "NOTE: U_BE = Cisco Business Entity (product line like Routing, Switching, Security, etc.)"
        )
        summary_parts.append(
            "      U_SUBBE = Cisco Sub-Business Entity (more specific product category)"
        )

        # Aggregate all U_BE breakdowns across all parties
        all_be_data = {}
        all_subbe_data = {}

        for idx, row in party_level_df.iterrows():
            # Parse BE_BREAKDOWN
            if (
                "BE_BREAKDOWN" in party_level_df.columns
                and pd.notna(row.get("BE_BREAKDOWN"))
                and str(row.get("BE_BREAKDOWN")).strip()
            ):
                be_str = str(row["BE_BREAKDOWN"])
                for item in be_str.split(";"):
                    if ":" in item:
                        be_name, count_str = item.split(":", 1)
                        be_name = be_name.strip()
                        try:
                            count = int(count_str.strip())
                            all_be_data[be_name] = all_be_data.get(be_name, 0) + count
                        except:
                            pass

            # Parse SUBBE_BREAKDOWN
            if (
                "SUBBE_BREAKDOWN" in party_level_df.columns
                and pd.notna(row.get("SUBBE_BREAKDOWN"))
                and str(row.get("SUBBE_BREAKDOWN")).strip()
            ):
                subbe_str = str(row["SUBBE_BREAKDOWN"])
                for item in subbe_str.split(";"):
                    if ":" in item:
                        subbe_name, count_str = item.split(":", 1)
                        subbe_name = subbe_name.strip()
                        try:
                            count = int(count_str.strip())
                            all_subbe_data[subbe_name] = (
                                all_subbe_data.get(subbe_name, 0) + count
                            )
                        except:
                            pass

        # Display top U_BE categories WITH COVERAGE BREAKDOWN
        if all_be_data:
            summary_parts.append("\nTop Cisco Business Entities (U_BE) by Asset Count:")
            sorted_be = sorted(all_be_data.items(), key=lambda x: x[1], reverse=True)

            # Also track coverage by U_BE
            be_coverage = {}

            # Parse coverage status by U_BE from the dataframe
            for idx, row in party_level_df.iterrows():
                if (
                    "BE_BREAKDOWN" in party_level_df.columns
                    and pd.notna(row.get("BE_BREAKDOWN"))
                    and str(row.get("BE_BREAKDOWN")).strip()
                ):
                    be_str = str(row["BE_BREAKDOWN"])
                    for item in be_str.split(";"):
                        if ":" in item:
                            be_name, count_str = item.split(":", 1)
                            be_name = be_name.strip()
                            try:
                                count = int(count_str.strip())
                                if be_name not in be_coverage:
                                    be_coverage[be_name] = {
                                        "covered": 0,
                                        "uncovered": 0,
                                        "never_covered": 0,
                                        "total": 0,
                                    }

                                # Estimate coverage distribution proportionally
                                party_total_ib = row.get("TOTAL_IB_COUNT", 0)
                                if party_total_ib > 0:
                                    covered_pct = (
                                        row.get("IB_COUNT_COVERED", 0) / party_total_ib
                                    )
                                    uncovered_pct = (
                                        row.get("IB_COUNT_UNCOVERED", 0)
                                        / party_total_ib
                                    )
                                    never_pct = (
                                        row.get("IB_COUNT_NEVER_COVERED", 0)
                                        / party_total_ib
                                    )

                                    be_coverage[be_name]["covered"] += int(
                                        count * covered_pct
                                    )
                                    be_coverage[be_name]["uncovered"] += int(
                                        count * uncovered_pct
                                    )
                                    be_coverage[be_name]["never_covered"] += int(
                                        count * never_pct
                                    )

                                be_coverage[be_name]["total"] += count
                            except:
                                pass

            for be_name, count in sorted_be[:10]:  # Top 10
                percentage = (count / total_ib * 100) if total_ib > 0 else 0

                # Add coverage breakdown if available
                if be_name in be_coverage:
                    cov = be_coverage[be_name]
                    covered = cov["covered"]
                    uncovered = cov["uncovered"]
                    never = cov["never_covered"]

                    summary_parts.append(
                        f"  - {be_name}: {count:,} assets ({percentage:.1f}%) "
                        f"[Covered: {covered:,}, Uncovered: {uncovered:,}, Never: {never:,}]"
                    )
                else:
                    summary_parts.append(
                        f"  - {be_name}: {count:,} assets ({percentage:.1f}%)"
                    )

        # Display top U_SUBBE categories WITH COVERAGE BREAKDOWN
        if all_subbe_data:
            summary_parts.append(
                "\nTop Cisco Sub-Business Entities (U_SUBBE) by Asset Count:"
            )
            sorted_subbe = sorted(
                all_subbe_data.items(), key=lambda x: x[1], reverse=True
            )

            # Also track coverage by U_SUBBE
            subbe_coverage = {}

            for idx, row in party_level_df.iterrows():
                if (
                    "SUBBE_BREAKDOWN" in party_level_df.columns
                    and pd.notna(row.get("SUBBE_BREAKDOWN"))
                    and str(row.get("SUBBE_BREAKDOWN")).strip()
                ):
                    subbe_str = str(row["SUBBE_BREAKDOWN"])
                    for item in subbe_str.split(";"):
                        if ":" in item:
                            subbe_name, count_str = item.split(":", 1)
                            subbe_name = subbe_name.strip()
                            try:
                                count = int(count_str.strip())
                                if subbe_name not in subbe_coverage:
                                    subbe_coverage[subbe_name] = {
                                        "covered": 0,
                                        "uncovered": 0,
                                        "never_covered": 0,
                                        "total": 0,
                                    }

                                # Estimate coverage distribution proportionally
                                party_total_ib = row.get("TOTAL_IB_COUNT", 0)
                                if party_total_ib > 0:
                                    covered_pct = (
                                        row.get("IB_COUNT_COVERED", 0) / party_total_ib
                                    )
                                    uncovered_pct = (
                                        row.get("IB_COUNT_UNCOVERED", 0)
                                        / party_total_ib
                                    )
                                    never_pct = (
                                        row.get("IB_COUNT_NEVER_COVERED", 0)
                                        / party_total_ib
                                    )

                                    subbe_coverage[subbe_name]["covered"] += int(
                                        count * covered_pct
                                    )
                                    subbe_coverage[subbe_name]["uncovered"] += int(
                                        count * uncovered_pct
                                    )
                                    subbe_coverage[subbe_name]["never_covered"] += int(
                                        count * never_pct
                                    )

                                subbe_coverage[subbe_name]["total"] += count
                            except:
                                pass

            for subbe_name, count in sorted_subbe[:10]:  # Top 10
                percentage = (count / total_ib * 100) if total_ib > 0 else 0

                # Add coverage breakdown if available
                if subbe_name in subbe_coverage:
                    cov = subbe_coverage[subbe_name]
                    covered = cov["covered"]
                    uncovered = cov["uncovered"]
                    never = cov["never_covered"]

                    summary_parts.append(
                        f"  - {subbe_name}: {count:,} assets ({percentage:.1f}%) "
                        f"[Covered: {covered:,}, Uncovered: {uncovered:,}, Never: {never:,}]"
                    )
                else:
                    summary_parts.append(
                        f"  - {subbe_name}: {count:,} assets ({percentage:.1f}%)"
                    )
    # ============================================================================
    # ADD THE NEW EOL/EOS BY ARCHITECTURE CODE HERE (PASTE BELOW)
    # ============================================================================

    # Add EOL/EOS metrics by Cisco Business Entity (U_BE/Architecture)
    if "BE_BREAKDOWN" in party_level_df.columns:
        summary_parts.append(
            "\n=== EOL/EOS RISK BY CISCO BUSINESS ENTITY (ARCHITECTURE) ==="
        )

        # Check if EOL/EOS columns exist
        has_eol_eos = any(
            col in party_level_df.columns
            for col in ["EOL_PASSED", "EOL_WITHIN_1YR", "EOS_PASSED", "EOS_WITHIN_1YR"]
        )

        if not has_eol_eos:
            # Need to add EOL/EOS columns to the aggregation
            summary_parts.append("Note: EOL/EOS data not available in current dataset")
        else:
            # Track EOL/EOS metrics by U_BE
            be_eol_eos_metrics = {}

            for idx, row in party_level_df.iterrows():
                if (
                    pd.notna(row.get("BE_BREAKDOWN"))
                    and str(row.get("BE_BREAKDOWN")).strip()
                ):
                    be_str = str(row["BE_BREAKDOWN"])
                    for item in be_str.split(";"):
                        if ":" in item:
                            be_name, count_str = item.split(":", 1)
                            be_name = be_name.strip()
                            try:
                                count = int(count_str.strip())

                                if be_name not in be_eol_eos_metrics:
                                    be_eol_eos_metrics[be_name] = {
                                        "total_assets": 0,
                                        "eol_passed": 0,
                                        "eol_within_1yr": 0,
                                        "eol_future": 0,
                                        "eol_unknown": 0,
                                        "eos_passed": 0,
                                        "eos_within_1yr": 0,
                                        "eos_future": 0,
                                        "eos_unknown": 0,
                                    }

                                # Estimate EOL/EOS distribution proportionally
                                party_total_ib = row.get("TOTAL_IB_COUNT", 0)
                                if party_total_ib > 0:
                                    proportion = count / party_total_ib

                                    be_eol_eos_metrics[be_name]["total_assets"] += count
                                    be_eol_eos_metrics[be_name]["eol_passed"] += int(
                                        row.get("EOL_PASSED", 0) * proportion
                                    )
                                    be_eol_eos_metrics[be_name]["eol_within_1yr"] += (
                                        int(row.get("EOL_WITHIN_1YR", 0) * proportion)
                                    )
                                    be_eol_eos_metrics[be_name]["eol_future"] += int(
                                        row.get("EOL_FUTURE", 0) * proportion
                                    )
                                    be_eol_eos_metrics[be_name]["eol_unknown"] += int(
                                        row.get("EOL_UNKNOWN", 0) * proportion
                                    )
                                    be_eol_eos_metrics[be_name]["eos_passed"] += int(
                                        row.get("EOS_PASSED", 0) * proportion
                                    )
                                    be_eol_eos_metrics[be_name]["eos_within_1yr"] += (
                                        int(row.get("EOS_WITHIN_1YR", 0) * proportion)
                                    )
                                    be_eol_eos_metrics[be_name]["eos_future"] += int(
                                        row.get("EOS_FUTURE", 0) * proportion
                                    )
                                    be_eol_eos_metrics[be_name]["eos_unknown"] += int(
                                        row.get("EOS_UNKNOWN", 0) * proportion
                                    )
                            except:
                                pass

            if be_eol_eos_metrics:
                summary_parts.append("\nEOL/EOS Metrics by Cisco Architecture (U_BE):")

                # Sort by total assets (descending)
                sorted_be_eol = sorted(
                    be_eol_eos_metrics.items(),
                    key=lambda x: x[1]["total_assets"],
                    reverse=True,
                )

                for be_name, metrics in sorted_be_eol[:10]:  # Top 10
                    total = metrics["total_assets"]
                    eol_risk = metrics["eol_passed"] + metrics["eol_within_1yr"]
                    eos_risk = metrics["eos_passed"] + metrics["eos_within_1yr"]

                    summary_parts.append(f"\n  {be_name} ({total:,} assets):")
                    summary_parts.append(f"    EOL Status:")
                    summary_parts.append(f"      - Past EOL: {metrics['eol_passed']:,}")
                    summary_parts.append(
                        f"      - EOL Within 1 Year: {metrics['eol_within_1yr']:,}"
                    )
                    summary_parts.append(
                        f"      - Future EOL: {metrics['eol_future']:,}"
                    )
                    summary_parts.append(
                        f"      - Unknown EOL: {metrics['eol_unknown']:,}"
                    )
                    summary_parts.append(
                        f"      - TOTAL AT RISK (Past + <1Yr): {eol_risk:,}"
                    )

                    summary_parts.append(f"    EOS Status:")
                    summary_parts.append(f"      - Past EOS: {metrics['eos_passed']:,}")
                    summary_parts.append(
                        f"      - EOS Within 1 Year: {metrics['eos_within_1yr']:,}"
                    )
                    summary_parts.append(
                        f"      - Future EOS: {metrics['eos_future']:,}"
                    )
                    summary_parts.append(
                        f"      - Unknown EOS: {metrics['eos_unknown']:,}"
                    )
                    summary_parts.append(
                        f"      - TOTAL AT RISK (Past + <1Yr): {eos_risk:,}"
                    )

                    # Calculate risk percentages
                    eol_risk_pct = (eol_risk / total * 100) if total > 0 else 0
                    eos_risk_pct = (eos_risk / total * 100) if total > 0 else 0
                    summary_parts.append(
                        f"    Risk Assessment: {eol_risk_pct:.1f}% EOL risk, {eos_risk_pct:.1f}% EOS risk"
                    )

    # ============================================================================
    # END OF NEW CODE - Serial reconciliation section continues below
    # ============================================================================

    # Add serial reconciliation summary if provided
    if serial_recon_df is not None and not serial_recon_df.empty and serial_summary:
        summary_parts.append(f"\n=== SERIAL NUMBER RECONCILIATION ===")
        summary_parts.append(serial_summary)

    return "\n".join(summary_parts)


def create_ea_powerpoint_with_ai(
    final_df: pd.DataFrame,
    ai_summary: str,
    ai_client,
    cav_name: str,
    cav_id: int,
    gu_name: str,
    output_folder: str = "analysis_output",
) -> str:
    """Creates a PowerPoint presentation with AI-generated content and structure."""
    os.makedirs(output_folder, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_cav_name = cav_name.replace(" ", "_").replace("/", "_")
    filename = f"{safe_cav_name}_CAV_{cav_id}_EA_Presentation_{timestamp}.pptx"
    filepath = os.path.join(output_folder, filename)

    # Safely get metrics with defaults
    total_ib = final_df.get("TOTAL_IB_COUNT", pd.Series([0])).sum()
    ib_covered = final_df.get("IB_COUNT_COVERED", pd.Series([0])).sum()
    ib_uncovered = final_df.get("IB_COUNT_UNCOVERED", pd.Series([0])).sum()
    ib_never_covered = final_df.get("IB_COUNT_NEVER_COVERED", pd.Series([0])).sum()
    total_cases = final_df.get("CASE_COUNT", pd.Series([0])).sum()
    total_sales = final_df.get("TOTAL_SALES", pd.Series([0])).sum()

    # Prepare data summary for AI
    data_summary = f"""
Customer: {gu_name}
CAV: {cav_name} (ID: {cav_id})

Key Metrics:
- Total Parties: {len(final_df)}
- Total Install Base: {total_ib:,.0f}
  - Covered: {ib_covered:,.0f}
  - Uncovered: {ib_uncovered:,.0f}
  - Never Covered: {ib_never_covered:,.0f}
- Total Cases: {total_cases:,.0f}
- Total Sales: ${total_sales:,.2f}

EOL/EOS Risk:
- Assets Past EOL: {final_df.get("EOL_PASSED", pd.Series([0])).sum():,.0f}
- Assets EOL Within 1 Year: {final_df.get("EOL_WITHIN_1YR", pd.Series([0])).sum():,.0f}
- Assets Past EOS: {final_df.get("EOS_PASSED", pd.Series([0])).sum():,.0f}
- Assets EOS Within 1 Year: {final_df.get("EOS_WITHIN_1YR", pd.Series([0])).sum():,.0f}

Top 5 Customers by Install Base:
"""

    if "TOTAL_IB_COUNT" in final_df.columns and total_ib > 0:
        top_5 = final_df.nlargest(5, "TOTAL_IB_COUNT")
        for idx, row in top_5.iterrows():
            party_name = row.get(
                "PARTY_NAME", f"Party {row.get('PARTY_ID', 'Unknown')}"
            )
            party_ib = row.get("TOTAL_IB_COUNT", 0)
            party_cases = row.get("CASE_COUNT", 0)
            data_summary += (
                f"\n- {party_name}: {party_ib:,.0f} assets, {party_cases:,.0f} cases"
            )

    print("🤖 Asking AI to generate PowerPoint content...")

    ppt_prompt = f"""Based on this EA analysis data, create a PowerPoint presentation structure with 6-8 slides.

{data_summary}

Previous AI Analysis:
{ai_summary}

For each slide, provide:
1. Slide title
2. Key points (3-5 bullet points per slide)
3. Any specific recommendations or insights

Format your response as:
SLIDE 1: [Title]
- [Bullet point 1]
- [Bullet point 2]
...

SLIDE 2: [Title]
...

Focus on:
- Executive summary
- Key findings and metrics
- Risk analysis (EOL/EOS)
- Coverage opportunities
- Top customer insights
- Actionable recommendations
"""

    try:
        ppt_content_response = ai_client.chat.completions.create(
            model=CISCO_AI_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert at creating executive presentations for Cisco Services sales. Create clear, actionable slide content.",
                },
                {"role": "user", "content": ppt_prompt},
            ],
            user=f'{{"appkey": "{CISCO_AI_APP_KEY}"}}',
        )
        ppt_content = ppt_content_response.choices[0].message.content
        print("✅ AI generated PowerPoint content structure")
    except Exception as e:
        logging.error(f"Error getting AI PowerPoint content: {e}")
        ppt_content = f"SLIDE 1: Executive Summary\n{ai_summary}"

    # Parse AI response into slides
    slides_data = []
    current_slide = None

    for line in ppt_content.split("\n"):
        line = line.strip()
        if line.startswith("SLIDE"):
            if current_slide:
                slides_data.append(current_slide)
            title = line.split(":", 1)[1].strip() if ":" in line else "Slide"
            current_slide = {"title": title, "bullets": []}
        elif line.startswith("-") and current_slide:
            bullet = line.lstrip("- ").strip()
            if bullet:
                current_slide["bullets"].append(bullet)

    if current_slide:
        slides_data.append(current_slide)

    # Create PowerPoint
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Title Slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = f"Enterprise Agreement Analysis"
    subtitle.text = f"{gu_name}\n{cav_name} (CAV ID: {cav_id})\n{datetime.now().strftime('%B %d, %Y')}"

    # Add AI-generated content slides
    for slide_data in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        title_shape = slide.shapes.title
        title_shape.text = slide_data["title"]

        body_shape = slide.placeholders[1]
        tf = body_shape.text_frame
        tf.clear()

        for bullet in slide_data["bullets"]:
            p = tf.add_paragraph()
            p.text = bullet
            p.level = 0
            p.font.size = Pt(14)
            p.space_after = Pt(12)

    # Add a data visualization slide
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title_shape = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    title_frame = title_shape.text_frame
    title_frame.text = "Install Base Coverage Distribution"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(28)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(0, 51, 141)

    chart_data = CategoryChartData()
    chart_data.categories = ["Covered", "Uncovered", "Never Covered"]

    # Safely get values
    covered_val = final_df.get("IB_COUNT_COVERED", pd.Series([0])).sum()
    uncovered_val = final_df.get("IB_COUNT_UNCOVERED", pd.Series([0])).sum()
    never_covered_val = final_df.get("IB_COUNT_NEVER_COVERED", pd.Series([0])).sum()

    chart_data.add_series("IB Count", (covered_val, uncovered_val, never_covered_val))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, Inches(1.5), Inches(1.5), Inches(7), Inches(5), chart_data
    ).chart
    chart.has_legend = True
    chart.legend.position = 2

    # Add metrics summary slide
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title_shape = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    title_frame = title_shape.text_frame
    title_frame.text = "Key Metrics at a Glance"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(28)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(0, 51, 141)

    metrics = [
        ("Total Parties", f"{len(final_df):,}"),
        (
            "Total Install Base",
            f"{final_df.get('TOTAL_IB_COUNT', pd.Series([0])).sum():,.0f}",
        ),
        (
            "IB Covered",
            f"{final_df.get('IB_COUNT_COVERED', pd.Series([0])).sum():,.0f}",
        ),
        (
            "IB Uncovered",
            f"{final_df.get('IB_COUNT_UNCOVERED', pd.Series([0])).sum():,.0f}",
        ),
        ("Total Cases", f"{final_df.get('CASE_COUNT', pd.Series([0])).sum():,.0f}"),
        ("Total Sales", f"${final_df.get('TOTAL_SALES', pd.Series([0])).sum():,.2f}"),
    ]

    box_width, box_height = 2.8, 1.2
    start_x, start_y = 0.5, 1.2

    for i, (label, value) in enumerate(metrics):
        row, col = i // 3, i % 3
        x = start_x + col * (box_width + 0.3)
        y = start_y + row * (box_height + 0.3)

        box = slide.shapes.add_shape(
            1, Inches(x), Inches(y), Inches(box_width), Inches(box_height)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(240, 240, 240)
        box.line.color.rgb = RGBColor(0, 51, 141)
        box.line.width = Pt(2)

        text_frame = box.text_frame
        text_frame.clear()

        p1 = text_frame.paragraphs[0]
        p1.text = label
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.alignment = PP_ALIGN.CENTER

        p2 = text_frame.add_paragraph()
        p2.text = value
        p2.font.size = Pt(20)
        p2.font.bold = True
        p2.font.color.rgb = RGBColor(0, 51, 141)
        p2.alignment = PP_ALIGN.CENTER

    prs.save(filepath)
    print(f"✅ AI-powered PowerPoint presentation saved: {filepath}")
    return filepath


def create_ea_powerpoint(
    final_df: pd.DataFrame,
    ai_summary: str,
    cav_name: str,
    cav_id: int,
    gu_name: str,
    output_folder: str = "analysis_output",
) -> str:
    """Creates a PowerPoint presentation with EA analysis results and AI insights."""
    os.makedirs(output_folder, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_cav_name = cav_name.replace(" ", "_").replace("/", "_")
    filename = f"{safe_cav_name}_CAV_{cav_id}_EA_Presentation_{timestamp}.pptx"
    filepath = os.path.join(output_folder, filename)

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Slide 1: Title
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = f"Enterprise Agreement Analysis"
    subtitle.text = f"{gu_name}\n{cav_name} (CAV ID: {cav_id})\n{datetime.now().strftime('%B %d, %Y')}"

    # Slide 2: AI Summary
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title_shape = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    title_frame = title_shape.text_frame
    title_frame.text = "Executive Summary"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(28)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(0, 51, 141)

    summary_shape = slide.shapes.add_textbox(
        Inches(0.5), Inches(1.0), Inches(9), Inches(5.5)
    )
    summary_frame = summary_shape.text_frame
    summary_frame.word_wrap = True
    summary_frame.text = ai_summary
    for paragraph in summary_frame.paragraphs:
        paragraph.font.size = Pt(12)

    # Slide 3: Key Metrics
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title_shape = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    title_frame = title_shape.text_frame
    title_frame.text = "Key Metrics Overview"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(28)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(0, 51, 141)

    metrics = [
        ("Total Parties", f"{len(final_df):,}"),
        (
            "Total Install Base",
            f"{final_df.get('TOTAL_IB_COUNT', pd.Series([0])).sum():,.0f}",
        ),
        (
            "IB Covered",
            f"{final_df.get('IB_COUNT_COVERED', pd.Series([0])).sum():,.0f}",
        ),
        (
            "IB Uncovered",
            f"{final_df.get('IB_COUNT_UNCOVERED', pd.Series([0])).sum():,.0f}",
        ),
        ("Total Cases", f"{final_df.get('CASE_COUNT', pd.Series([0])).sum():,.0f}"),
        ("Total Sales", f"${final_df.get('TOTAL_SALES', pd.Series([0])).sum():,.2f}"),
    ]

    box_width, box_height = 2.8, 1.2
    start_x, start_y = 0.5, 1.2

    for i, (label, value) in enumerate(metrics):
        row, col = i // 3, i % 3
        x = start_x + col * (box_width + 0.3)
        y = start_y + row * (box_height + 0.3)

        box = slide.shapes.add_shape(
            1, Inches(x), Inches(y), Inches(box_width), Inches(box_height)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(240, 240, 240)
        box.line.color.rgb = RGBColor(0, 51, 141)
        box.line.width = Pt(2)

        text_frame = box.text_frame
        text_frame.clear()

        p1 = text_frame.paragraphs[0]
        p1.text = label
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.alignment = PP_ALIGN.CENTER

        p2 = text_frame.add_paragraph()
        p2.text = value
        p2.font.size = Pt(20)
        p2.font.bold = True
        p2.font.color.rgb = RGBColor(0, 51, 141)
        p2.alignment = PP_ALIGN.CENTER

    # Slide 4: Coverage Chart
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title_shape = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    title_frame = title_shape.text_frame
    title_frame.text = "Install Base Coverage Distribution"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(28)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(0, 51, 141)

    chart_data = CategoryChartData()
    chart_data.categories = ["Covered", "Uncovered", "Never Covered"]

    covered_val = final_df.get("IB_COUNT_COVERED", pd.Series([0])).sum()
    uncovered_val = final_df.get("IB_COUNT_UNCOVERED", pd.Series([0])).sum()
    never_covered_val = final_df.get("IB_COUNT_NEVER_COVERED", pd.Series([0])).sum()

    chart_data.add_series("IB Count", (covered_val, uncovered_val, never_covered_val))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, Inches(1.5), Inches(1.5), Inches(7), Inches(5), chart_data
    ).chart
    chart.has_legend = True
    chart.legend.position = 2

    prs.save(filepath)
    print(f"✅ PowerPoint presentation saved: {filepath}")
    return filepath


# ADD THESE TWO FUNCTIONS RIGHT BEFORE perform_ea_analysis_with_serial_reconciliation


def aggregate_ib_by_business_entity_external(ib_df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate IB counts by U_BE and U_SUBBE."""
    if ib_df.empty:
        return pd.DataFrame()

    # Uppercase columns
    ib_df = ib_df.copy()
    ib_df.columns = [str(col).upper() for col in ib_df.columns]

    group_keys = ["CAV_BU_ID", "PARTY_ID"]

    if not all(k in ib_df.columns for k in group_keys):
        print("Warning: BE aggregation skipped - missing keys")
        return pd.DataFrame()

    if "U_BE" not in ib_df.columns and "U_SUBBE" not in ib_df.columns:
        print("Warning: U_BE and U_SUBBE columns not found")
        return pd.DataFrame()

    ib_df["CAV_BU_ID"] = pd.to_numeric(ib_df["CAV_BU_ID"], errors="coerce").astype(
        pd.Int64Dtype()
    )
    ib_df["PARTY_ID"] = pd.to_numeric(ib_df["PARTY_ID"], errors="coerce").astype(
        pd.Int64Dtype()
    )
    ib_df_valid = ib_df.dropna(subset=["CAV_BU_ID", "PARTY_ID"]).copy()

    if ib_df_valid.empty:
        return pd.DataFrame()

    # FIXED: Use QUANTITY properly and filter out zero/negative
    if "QUANTITY" in ib_df_valid.columns:
        ib_df_valid["_CNT"] = pd.to_numeric(
            ib_df_valid["QUANTITY"], errors="coerce"
        ).fillna(0)
    else:
        ib_df_valid["_CNT"] = 1

    ib_df_valid = ib_df_valid[ib_df_valid["_CNT"] > 0].copy()

    if ib_df_valid.empty:
        print("Warning: No valid rows with positive quantity")
        return pd.DataFrame()

    result_data = []

    for (cav_bu_id, party_id_val), party_data in ib_df_valid.groupby(group_keys):
        row = {"CAV_BU_ID": cav_bu_id, "PARTY_ID": party_id_val}

        be_breakdown = []
        if "U_BE" in party_data.columns:
            be_counts = (
                party_data.groupby("U_BE")["_CNT"].sum().sort_values(ascending=False)
            )
            be_breakdown = [
                f"{be}: {int(cnt)}"
                for be, cnt in be_counts.items()
                if pd.notna(be) and str(be).strip() != ""
            ]

        subbe_breakdown = []
        if "U_SUBBE" in party_data.columns:
            subbe_counts = (
                party_data.groupby("U_SUBBE")["_CNT"].sum().sort_values(ascending=False)
            )
            subbe_breakdown = [
                f"{subbe}: {int(cnt)}"
                for subbe, cnt in subbe_counts.items()
                if pd.notna(subbe) and str(subbe).strip() != ""
            ]

        row["BE_BREAKDOWN"] = "; ".join(be_breakdown) if be_breakdown else ""
        row["SUBBE_BREAKDOWN"] = "; ".join(subbe_breakdown) if subbe_breakdown else ""
        result_data.append(row)

    result = pd.DataFrame(result_data)

    if not result.empty:
        result["CAV_BU_ID"] = result["CAV_BU_ID"].astype(pd.Int64Dtype())
        result["PARTY_ID"] = result["PARTY_ID"].astype(pd.Int64Dtype())

        # Print total counts for verification
        total_assets = 0
        for _, row in result.iterrows():
            if row["BE_BREAKDOWN"]:
                for item in row["BE_BREAKDOWN"].split(";"):
                    if ":" in item:
                        total_assets += int(item.split(":")[1].strip())

        print(f"BE Summary: {len(result)} parties with BE/Sub-BE breakdowns")
        print(f"Total assets in BE breakdown: {total_assets:,}")

    return result


def aggregate_eol_eos_metrics_external(ib_df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate EOL and EOS metrics."""
    if ib_df.empty:
        return pd.DataFrame()

    # Uppercase columns
    ib_df = ib_df.copy()
    ib_df.columns = [str(col).upper() for col in ib_df.columns]

    group_keys = ["CAV_BU_ID", "PARTY_ID"]

    if not all(k in ib_df.columns for k in group_keys):
        return pd.DataFrame()

    if "EOL_DATE" not in ib_df.columns and "EOS_DATE" not in ib_df.columns:
        return pd.DataFrame()

    ib_df["CAV_BU_ID"] = pd.to_numeric(ib_df["CAV_BU_ID"], errors="coerce").astype(
        pd.Int64Dtype()
    )
    ib_df["PARTY_ID"] = pd.to_numeric(ib_df["PARTY_ID"], errors="coerce").astype(
        pd.Int64Dtype()
    )
    ib_df_valid = ib_df.dropna(subset=["CAV_BU_ID", "PARTY_ID"]).copy()

    if ib_df_valid.empty:
        return pd.DataFrame()

    if "EOL_DATE" in ib_df_valid.columns:
        ib_df_valid["EOL_DATE"] = pd.to_datetime(
            ib_df_valid["EOL_DATE"], errors="coerce"
        )
    if "EOS_DATE" in ib_df_valid.columns:
        ib_df_valid["EOS_DATE"] = pd.to_datetime(
            ib_df_valid["EOS_DATE"], errors="coerce"
        )

    if "QUANTITY" in ib_df_valid.columns:
        ib_df_valid["_CNT"] = pd.to_numeric(
            ib_df_valid["QUANTITY"], errors="coerce"
        ).fillna(1)
    else:
        ib_df_valid["_CNT"] = 1

    current_date = pd.Timestamp.now()
    result_data = []

    for (cav_bu_id, party_id_val), party_data in ib_df_valid.groupby(group_keys):
        row = {"CAV_BU_ID": cav_bu_id, "PARTY_ID": party_id_val}

        # EOL Metrics
        if "EOL_DATE" in party_data.columns:
            eol_data = party_data[party_data["EOL_DATE"].notna()]
            eol_passed = eol_data[eol_data["EOL_DATE"] < current_date]["_CNT"].sum()
            eol_upcoming_1yr = eol_data[
                (eol_data["EOL_DATE"] >= current_date)
                & (eol_data["EOL_DATE"] <= current_date + pd.DateOffset(years=1))
            ]["_CNT"].sum()
            eol_future = eol_data[
                eol_data["EOL_DATE"] > current_date + pd.DateOffset(years=1)
            ]["_CNT"].sum()
            eol_unknown = party_data[party_data["EOL_DATE"].isna()]["_CNT"].sum()

            row["EOL_PASSED"] = int(eol_passed)
            row["EOL_WITHIN_1YR"] = int(eol_upcoming_1yr)
            row["EOL_FUTURE"] = int(eol_future)
            row["EOL_UNKNOWN"] = int(eol_unknown)
        else:
            row["EOL_PASSED"] = row["EOL_WITHIN_1YR"] = row["EOL_FUTURE"] = 0
            row["EOL_UNKNOWN"] = int(party_data["_CNT"].sum())

        # EOS Metrics
        if "EOS_DATE" in party_data.columns:
            eos_data = party_data[party_data["EOS_DATE"].notna()]
            eos_passed = eos_data[eos_data["EOS_DATE"] < current_date]["_CNT"].sum()
            eos_upcoming_1yr = eos_data[
                (eos_data["EOS_DATE"] >= current_date)
                & (eos_data["EOS_DATE"] <= current_date + pd.DateOffset(years=1))
            ]["_CNT"].sum()
            eos_future = eos_data[
                eos_data["EOS_DATE"] > current_date + pd.DateOffset(years=1)
            ]["_CNT"].sum()
            eos_unknown = party_data[party_data["EOS_DATE"].isna()]["_CNT"].sum()

            row["EOS_PASSED"] = int(eos_passed)
            row["EOS_WITHIN_1YR"] = int(eos_upcoming_1yr)
            row["EOS_FUTURE"] = int(eos_future)
            row["EOS_UNKNOWN"] = int(eos_unknown)
        else:
            row["EOS_PASSED"] = row["EOS_WITHIN_1YR"] = row["EOS_FUTURE"] = 0
            row["EOS_UNKNOWN"] = int(party_data["_CNT"].sum())

        result_data.append(row)

    result = pd.DataFrame(result_data)

    if not result.empty:
        result["CAV_BU_ID"] = result["CAV_BU_ID"].astype(pd.Int64Dtype())
        result["PARTY_ID"] = result["PARTY_ID"].astype(pd.Int64Dtype())
        print(f"EOL/EOS Summary: {len(result)} parties")

    return result


def perform_ea_analysis_with_serial_reconciliation(
    target_cav_id_int: int,
    customer_detail_by_cav_df: pd.DataFrame,
    case_history_df: pd.DataFrame,
    install_base_df: pd.DataFrame,
    sales_history_df: pd.DataFrame,
    customer_serial_ib_df: pd.DataFrame,
    eamp_proposal_quote_df: pd.DataFrame,
    output_folder: str = "analysis_output",
) -> Tuple[pd.DataFrame, str, pd.DataFrame, str, str]:
    """
    Unified EA analysis with integrated serial number reconciliation.

    Returns:
        Tuple of (final_activity_summary_df, analysis_log_string, serial_reconciliation_df, serial_summary_string, serial_excel_path)
    """
    messages: List[str] = []

    def _upper_cols(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = [str(col).upper() for col in df.columns]
        return df

    def _safe_numeric(df: pd.DataFrame, cols: List[str]):
        for c in cols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        return df

    def _aggregate_ib_by_business_entity(ib_df: pd.DataFrame) -> pd.DataFrame:
        """Aggregate IB counts by U_BE and U_SUBBE."""
        if ib_df.empty:
            return pd.DataFrame()

        ib_df = _upper_cols(ib_df)

        group_keys = ["CAV_BU_ID", "PARTY_ID"]

        if not all(k in ib_df.columns for k in group_keys):
            messages.append(f"Warning: BE aggregation skipped - missing keys")
            return pd.DataFrame()

        if "U_BE" not in ib_df.columns and "U_SUBBE" not in ib_df.columns:
            messages.append(
                "Warning: U_BE and U_SUBBE columns not found in install_base_df"
            )
            return pd.DataFrame()

        ib_df["CAV_BU_ID"] = pd.to_numeric(ib_df["CAV_BU_ID"], errors="coerce").astype(
            pd.Int64Dtype()
        )
        ib_df["PARTY_ID"] = pd.to_numeric(ib_df["PARTY_ID"], errors="coerce").astype(
            pd.Int64Dtype()
        )

        ib_df_valid = ib_df.dropna(subset=["CAV_BU_ID", "PARTY_ID"]).copy()

        if ib_df_valid.empty:
            messages.append("Warning: No valid IB rows for BE aggregation")
            return pd.DataFrame()

        # FIXED: Use QUANTITY column properly
        if "QUANTITY" in ib_df_valid.columns:
            ib_df_valid["_CNT"] = pd.to_numeric(
                ib_df_valid["QUANTITY"], errors="coerce"
            ).fillna(0)
        else:
            ib_df_valid["_CNT"] = 1

        # Remove rows with 0 or negative quantity
        ib_df_valid = ib_df_valid[ib_df_valid["_CNT"] > 0].copy()

        if ib_df_valid.empty:
            messages.append(
                "Warning: No valid IB rows with positive quantity for BE aggregation"
            )
            return pd.DataFrame()

        result_data = []

        for (cav_bu_id, party_id_val), party_data in ib_df_valid.groupby(group_keys):
            row = {"CAV_BU_ID": cav_bu_id, "PARTY_ID": party_id_val}

            be_breakdown = []
            if "U_BE" in party_data.columns:
                be_counts = (
                    party_data.groupby("U_BE")["_CNT"]
                    .sum()
                    .sort_values(ascending=False)
                )
                be_breakdown = [
                    f"{be}: {int(cnt)}"
                    for be, cnt in be_counts.items()
                    if pd.notna(be) and str(be).strip() != ""
                ]

            subbe_breakdown = []
            if "U_SUBBE" in party_data.columns:
                subbe_counts = (
                    party_data.groupby("U_SUBBE")["_CNT"]
                    .sum()
                    .sort_values(ascending=False)
                )
                subbe_breakdown = [
                    f"{subbe}: {int(cnt)}"
                    for subbe, cnt in subbe_counts.items()
                    if pd.notna(subbe) and str(subbe).strip() != ""
                ]

            row["BE_BREAKDOWN"] = "; ".join(be_breakdown) if be_breakdown else ""
            row["SUBBE_BREAKDOWN"] = (
                "; ".join(subbe_breakdown) if subbe_breakdown else ""
            )
            result_data.append(row)

        result = pd.DataFrame(result_data)

        if not result.empty:
            result["CAV_BU_ID"] = result["CAV_BU_ID"].astype(pd.Int64Dtype())
            result["PARTY_ID"] = result["PARTY_ID"].astype(pd.Int64Dtype())
            print(f"BE Summary: {len(result)} parties with BE/Sub-BE breakdowns")

        return result

    def _aggregate_eol_eos_metrics(ib_df: pd.DataFrame) -> pd.DataFrame:
        """Aggregate EOL and EOS date metrics by party."""

        if ib_df.empty:
            return pd.DataFrame()

        # FIXED: Correct indentation (4 spaces, not 8)
        ib_df = _upper_cols(ib_df)

        group_keys = ["CAV_BU_ID", "PARTY_ID"]

        if not all(k in ib_df.columns for k in group_keys):
            return pd.DataFrame()

        if "EOL_DATE" not in ib_df.columns and "EOS_DATE" not in ib_df.columns:
            return pd.DataFrame()

        ib_df["CAV_BU_ID"] = pd.to_numeric(ib_df["CAV_BU_ID"], errors="coerce").astype(
            pd.Int64Dtype()
        )
        ib_df["PARTY_ID"] = pd.to_numeric(ib_df["PARTY_ID"], errors="coerce").astype(
            pd.Int64Dtype()
        )

        ib_df_valid = ib_df.dropna(subset=["CAV_BU_ID", "PARTY_ID"]).copy()

        if ib_df_valid.empty:
            return pd.DataFrame()

        if "EOL_DATE" in ib_df_valid.columns:
            ib_df_valid["EOL_DATE"] = pd.to_datetime(
                ib_df_valid["EOL_DATE"], errors="coerce"
            )
        if "EOS_DATE" in ib_df_valid.columns:
            ib_df_valid["EOS_DATE"] = pd.to_datetime(
                ib_df_valid["EOS_DATE"], errors="coerce"
            )

        if "QUANTITY" in ib_df_valid.columns:
            ib_df_valid["_CNT"] = pd.to_numeric(
                ib_df_valid["QUANTITY"], errors="coerce"
            ).fillna(1)
        else:
            ib_df_valid["_CNT"] = 1

        current_date = pd.Timestamp.now()
        result_data = []

        for (cav_bu_id, party_id_val), party_data in ib_df_valid.groupby(group_keys):
            row = {"CAV_BU_ID": cav_bu_id, "PARTY_ID": party_id_val}

            # EOL Metrics
            if "EOL_DATE" in party_data.columns:
                eol_data = party_data[party_data["EOL_DATE"].notna()]

                eol_passed = eol_data[eol_data["EOL_DATE"] < current_date]["_CNT"].sum()
                eol_upcoming_1yr = eol_data[
                    (eol_data["EOL_DATE"] >= current_date)
                    & (eol_data["EOL_DATE"] <= current_date + pd.DateOffset(years=1))
                ]["_CNT"].sum()
                eol_future = eol_data[
                    eol_data["EOL_DATE"] > current_date + pd.DateOffset(years=1)
                ]["_CNT"].sum()
                eol_unknown = party_data[party_data["EOL_DATE"].isna()]["_CNT"].sum()

                row["EOL_PASSED"] = int(eol_passed)
                row["EOL_WITHIN_1YR"] = int(eol_upcoming_1yr)
                row["EOL_FUTURE"] = int(eol_future)
                row["EOL_UNKNOWN"] = int(eol_unknown)
            else:
                row["EOL_PASSED"] = row["EOL_WITHIN_1YR"] = row["EOL_FUTURE"] = 0
                row["EOL_UNKNOWN"] = int(party_data["_CNT"].sum())

            # EOS Metrics
            if "EOS_DATE" in party_data.columns:
                eos_data = party_data[party_data["EOS_DATE"].notna()]

                eos_passed = eos_data[eos_data["EOS_DATE"] < current_date]["_CNT"].sum()
                eos_upcoming_1yr = eos_data[
                    (eos_data["EOS_DATE"] >= current_date)
                    & (eos_data["EOS_DATE"] <= current_date + pd.DateOffset(years=1))
                ]["_CNT"].sum()
                eos_future = eos_data[
                    eos_data["EOS_DATE"] > current_date + pd.DateOffset(years=1)
                ]["_CNT"].sum()
                eos_unknown = party_data[party_data["EOS_DATE"].isna()]["_CNT"].sum()

                row["EOS_PASSED"] = int(eos_passed)
                row["EOS_WITHIN_1YR"] = int(eos_upcoming_1yr)
                row["EOS_FUTURE"] = int(eos_future)
                row["EOS_UNKNOWN"] = int(eos_unknown)
            else:
                row["EOS_PASSED"] = row["EOS_WITHIN_1YR"] = row["EOS_FUTURE"] = 0
                row["EOS_UNKNOWN"] = int(party_data["_CNT"].sum())

            result_data.append(row)

        result = pd.DataFrame(result_data)

        if not result.empty:
            result["CAV_BU_ID"] = result["CAV_BU_ID"].astype(pd.Int64Dtype())
            result["PARTY_ID"] = result["PARTY_ID"].astype(pd.Int64Dtype())
            print(f"EOL/EOS Summary: {len(result)} parties")

        return result

    # Serial Number Reconciliation (Integrated)
    def perform_serial_reconciliation() -> Tuple[pd.DataFrame, str, str]:
        """Internal function to validate and reconcile serial numbers. Returns (df, summary, filepath)"""
        serial_messages = []

        print("\n" + "=" * 70)
        print("SERIAL NUMBER VALIDATION & RECONCILIATION")
        print("=" * 70)

        customer_df = customer_serial_ib_df.copy()
        # FIXED: Convert columns to string first, then uppercase
        customer_df.columns = [str(col).upper() for col in customer_df.columns]

        ib_df = install_base_df.copy()
        # FIXED: Convert columns to string first, then uppercase
        ib_df.columns = [str(col).upper() for col in ib_df.columns]

        print(f"Customer file: {len(customer_df)} records")
        print(f"Cisco IB file: {len(ib_df)} records")

        customer_serial_col = None
        for col in ["SERIAL_NUMBER", "SERIAL", "SERIAL_NO", "SN"]:
            if col in customer_df.columns:
                customer_serial_col = col
                break

        if not customer_serial_col:
            error_msg = f"Error: Serial number column not found in customer file."
            print(error_msg)
            serial_messages.append(error_msg)
            return pd.DataFrame(), "\n".join(serial_messages), ""

        ib_serial_col = None
        for col in ["SERIAL_NUMBER", "SERIAL", "SERIAL_NO", "SN"]:
            if col in ib_df.columns:
                ib_serial_col = col
                break

        if not ib_serial_col:
            error_msg = f"Error: Serial number column not found in Cisco IB file."
            print(error_msg)
            serial_messages.append(error_msg)
            return pd.DataFrame(), "\n".join(serial_messages), ""

        print(
            f"Using serial columns: Customer='{customer_serial_col}', Cisco IB='{ib_serial_col}'"
        )

        customer_df["SERIAL_NUMBER_CLEAN"] = (
            customer_df[customer_serial_col].astype(str).str.strip().str.upper()
        )
        ib_df["SERIAL_NUMBER_CLEAN"] = (
            ib_df[ib_serial_col].astype(str).str.strip().str.upper()
        )

        customer_df = customer_df[customer_df["SERIAL_NUMBER_CLEAN"].notna()]
        customer_df = customer_df[customer_df["SERIAL_NUMBER_CLEAN"] != ""]
        customer_df = customer_df[customer_df["SERIAL_NUMBER_CLEAN"] != "NAN"]

        ib_df = ib_df[ib_df["SERIAL_NUMBER_CLEAN"].notna()]
        ib_df = ib_df[ib_df["SERIAL_NUMBER_CLEAN"] != ""]
        ib_df = ib_df[ib_df["SERIAL_NUMBER_CLEAN"] != "NAN"]

        print(f"After cleaning - Customer: {len(customer_df)}, Cisco IB: {len(ib_df)}")

        customer_serials = set(customer_df["SERIAL_NUMBER_CLEAN"].unique())
        ib_serials = set(ib_df["SERIAL_NUMBER_CLEAN"].unique())

        in_both = customer_serials & ib_serials
        customer_only = customer_serials - ib_serials
        cisco_only = ib_serials - customer_serials

        print(f"\n{'=' * 70}")
        print("RECONCILIATION RESULTS")
        print(f"{'=' * 70}")
        print(f"✓ Matched (In Both):     {len(in_both):>6,}")
        print(f"⚠ Customer Only:         {len(customer_only):>6,}")
        print(f"ℹ Cisco Only:            {len(cisco_only):>6,}")
        print(f"{'=' * 70}")

        match_rate = (
            (len(in_both) / len(customer_serials) * 100)
            if len(customer_serials) > 0
            else 0
        )

        serial_messages.append(
            f"Serial Reconciliation: {len(in_both)} matched ({match_rate:.1f}%)"
        )

        # Create reconciliation dataframe
        reconciliation_df = pd.DataFrame(
            {
                "SERIAL_NUMBER": list(customer_serials | ib_serials),
                "RECONCILIATION_STATUS": [
                    "Matched"
                    if s in in_both
                    else "Customer Only"
                    if s in customer_only
                    else "Cisco Only"
                    for s in (customer_serials | ib_serials)
                ],
            }
        )

        # Save to Excel
        os.makedirs(output_folder, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"Serial_Number_Reconciliation_{timestamp}.xlsx"
        excel_filepath = os.path.join(output_folder, excel_filename)

        # Create summary dataframe
        summary_data = pd.DataFrame(
            {
                "Category": [
                    "Matched (In Both)",
                    "Customer Only",
                    "Cisco Only",
                    "Total Unique Serials",
                ],
                "Count": [
                    len(in_both),
                    len(customer_only),
                    len(cisco_only),
                    len(customer_serials | ib_serials),
                ],
                "Percentage": [
                    f"{(len(in_both) / len(customer_serials) * 100) if len(customer_serials) > 0 else 0:.1f}%",
                    f"{(len(customer_only) / len(customer_serials) * 100) if len(customer_serials) > 0 else 0:.1f}%",
                    f"{(len(cisco_only) / len(ib_serials) * 100) if len(ib_serials) > 0 else 0:.1f}%",
                    "100.0%",
                ],
            }
        )

        with pd.ExcelWriter(excel_filepath, engine="xlsxwriter") as writer:
            # Sheet 1: Summary
            summary_data.to_excel(writer, sheet_name="Summary", index=False)

        # Sheet 2: Detailed Reconciliation
        reconciliation_df.to_excel(
            writer, sheet_name="Detailed Reconciliation", index=False
        )

        workbook = writer.book

        # Format Summary sheet
        worksheet_summary = writer.sheets["Summary"]
        header_format = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#4472C4",
                "font_color": "white",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        )

        # Highlight format for metrics
        metric_format = workbook.add_format(
            {"bold": True, "font_size": 12, "align": "left", "valign": "vcenter"}
        )
        count_format = workbook.add_format(
            {
                "font_size": 12,
                "align": "center",
                "valign": "vcenter",
                "num_format": "#,##0",
            }
        )

        for col_num, col_name in enumerate(summary_data.columns):
            worksheet_summary.write(0, col_num, col_name, header_format)

            worksheet_summary.set_column(0, 0, 25)  # Category column
            worksheet_summary.set_column(1, 1, 15)  # Count column
            worksheet_summary.set_column(2, 2, 15)  # Percentage column

            # Format Detailed Reconciliation sheet
            worksheet_detail = writer.sheets["Detailed Reconciliation"]
        for col_num, col_name in enumerate(reconciliation_df.columns):
            worksheet_detail.write(0, col_num, col_name, header_format)
            worksheet_detail.set_column(col_num, col_num, 20)
            worksheet_detail.freeze_panes(1, 0)
        # print(f"✅ Reconciliation report saved: {excel_filepath}")
        print(f"✅ Reconciliation report saved: {excel_filepath}")
        print(f"   📊 Sheet 1: Summary ({len(summary_data)} metrics)")
        print(
            f"   📋 Sheet 2: Detailed Reconciliation ({len(reconciliation_df)} serials)"
        )

        # Create detailed summary for AI
        match_rate = (
            (len(in_both) / len(customer_serials) * 100)
            if len(customer_serials) > 0
            else 0
        )
        customer_only_pct = (
            (len(customer_only) / len(customer_serials) * 100)
            if len(customer_serials) > 0
            else 0
        )
        cisco_only_pct = (
            (len(cisco_only) / len(ib_serials) * 100) if len(ib_serials) > 0 else 0
        )

        detailed_summary = f"""
        SERIAL NUMBER RECONCILIATION SUMMARY:
        =====================================
        Total Customer Serials: {len(customer_serials):,}
        Total Cisco IB Serials: {len(ib_serials):,}
        Total Unique Serials: {len(customer_serials | ib_serials):,}

        MATCH RESULTS:
        - Matched (In Both):     {len(in_both):,} ({match_rate:.1f}% of customer serials)
        - Customer Only:         {len(customer_only):,} ({customer_only_pct:.1f}% of customer serials)
        - Cisco Only:            {len(cisco_only):,} ({cisco_only_pct:.1f}% of Cisco IB serials)

        DATA QUALITY ASSESSMENT:
        - Match Rate: {match_rate:.1f}%
        - Discrepancy Rate: {100 - match_rate:.1f}%

        RECOMMENDED ACTIONS:
        {"- ✅ Excellent data quality. Serial matching is highly accurate." if match_rate >= 95 else ""}
        {"- ⚠️ Good data quality. Review customer-only serials for potential data entry errors." if 90 <= match_rate < 95 else ""}
        {"- ❌ Poor data quality. Immediate reconciliation needed. Many serials don't match." if match_rate < 90 else ""}
        {"- Review " + str(len(customer_only)) + " customer-only serials - may be non-Cisco or decommissioned assets." if len(customer_only) > 0 else ""}
        {"- " + str(len(cisco_only)) + " Cisco serials not in customer inventory - may indicate missing coverage opportunities." if len(cisco_only) > 0 else ""}
        """

        serial_messages.append(detailed_summary)

        return reconciliation_df, "\n".join(serial_messages), excel_filepath

        # return reconciliation_df, "\n".join(serial_messages), excel_filepath

    # EA Analysis Helper Functions
    def _aggregate_ib_from_install_base(ib_df: pd.DataFrame) -> pd.DataFrame:
        if ib_df.empty:
            return pd.DataFrame()

        ib_df = _upper_cols(ib_df)
        group_keys = ["CAV_BU_ID", "PARTY_ID"]

        if not all(k in ib_df.columns for k in group_keys):
            messages.append(
                "Warning: IB aggregation skipped - missing CAV_BU_ID or PARTY_ID"
            )
            return pd.DataFrame()

        ib_df["CAV_BU_ID"] = pd.to_numeric(ib_df["CAV_BU_ID"], errors="coerce").astype(
            pd.Int64Dtype()
        )
        ib_df["PARTY_ID"] = pd.to_numeric(ib_df["PARTY_ID"], errors="coerce").astype(
            pd.Int64Dtype()
        )
        ib_df_valid = ib_df.dropna(subset=group_keys).copy()

        if ib_df_valid.empty:
            messages.append("Warning: No valid IB rows after cleaning")
            return pd.DataFrame()

        # Calculate count - CRITICAL: Use QUANTITY column directly
        if "QUANTITY" in ib_df_valid.columns:
            ib_df_valid["_CNT"] = pd.to_numeric(
                ib_df_valid["QUANTITY"], errors="coerce"
            ).fillna(0)
        else:
            ib_df_valid["_CNT"] = 1

        # Remove rows with 0 or negative quantity
        ib_df_valid = ib_df_valid[ib_df_valid["_CNT"] > 0].copy()

        if ib_df_valid.empty:
            messages.append("Warning: No valid IB rows with positive quantity")
            return pd.DataFrame()

        # Check for COVERAGE_STATUS column
        if "COVERAGE_STATUS" not in ib_df_valid.columns:
            messages.append("Warning: COVERAGE_STATUS column not found in install base")
            return pd.DataFrame()

        # Clean coverage status
        ib_df_valid["_COV_STATUS"] = (
            ib_df_valid["COVERAGE_STATUS"].astype(str).str.strip().str.upper()
        )

        result_data = []

        for (cav_bu_id, party_id_val), party_data in ib_df_valid.groupby(group_keys):
            # A = Active (covered)
            # I = Inactive (uncovered)
            # N = Never covered
            covered = party_data[party_data["_COV_STATUS"] == "A"]["_CNT"].sum()
            uncovered = party_data[party_data["_COV_STATUS"] == "I"]["_CNT"].sum()
            never_covered = party_data[party_data["_COV_STATUS"] == "N"]["_CNT"].sum()

            result_data.append(
                {
                    "CAV_BU_ID": cav_bu_id,
                    "PARTY_ID": party_id_val,
                    "IB_COUNT_COVERED": int(covered),
                    "IB_COUNT_UNCOVERED": int(uncovered),
                    "IB_COUNT_NEVER_COVERED": int(never_covered),
                }
            )

        result = pd.DataFrame(result_data)

        if not result.empty:
            result["CAV_BU_ID"] = result["CAV_BU_ID"].astype(pd.Int64Dtype())
            result["PARTY_ID"] = result["PARTY_ID"].astype(pd.Int64Dtype())
            total_assets = (
                result["IB_COUNT_COVERED"].sum()
                + result["IB_COUNT_UNCOVERED"].sum()
                + result["IB_COUNT_NEVER_COVERED"].sum()
            )
            print(
                f"IB Summary: {len(result)} parties with {total_assets:,} total assets"
            )
            print(f"  - Covered (A): {result['IB_COUNT_COVERED'].sum():,}")
            print(f"  - Uncovered (I): {result['IB_COUNT_UNCOVERED'].sum():,}")
            print(f"  - Never Covered (N): {result['IB_COUNT_NEVER_COVERED'].sum():,}")

        return result

    def _aggregate_case_history(case_df: pd.DataFrame) -> pd.DataFrame:
        if case_df.empty:
            return pd.DataFrame()

        case_df = _upper_cols(case_df)
        group_keys = ["CAV_BU_ID", "PARTY_ID"]

        if not all(k in case_df.columns for k in group_keys):
            return pd.DataFrame()

        case_df["CAV_BU_ID"] = pd.to_numeric(
            case_df["CAV_BU_ID"], errors="coerce"
        ).astype(pd.Int64Dtype())
        case_df["PARTY_ID"] = pd.to_numeric(
            case_df["PARTY_ID"], errors="coerce"
        ).astype(pd.Int64Dtype())
        case_df_valid = case_df.dropna(subset=group_keys).copy()

        if case_df_valid.empty:
            return pd.DataFrame()

        result = case_df_valid.groupby(group_keys).size().reset_index(name="CASE_COUNT")
        return result

    def _aggregate_sales_history(sales_df: pd.DataFrame) -> pd.DataFrame:
        if sales_df.empty:
            return pd.DataFrame()

        sales_df = _upper_cols(sales_df)
        group_keys = ["CAV_BU_ID", "PARTY_ID"]

        if not all(k in sales_df.columns for k in group_keys):
            return pd.DataFrame()

        sales_df["CAV_BU_ID"] = pd.to_numeric(
            sales_df["CAV_BU_ID"], errors="coerce"
        ).astype(pd.Int64Dtype())
        sales_df["PARTY_ID"] = pd.to_numeric(
            sales_df["PARTY_ID"], errors="coerce"
        ).astype(pd.Int64Dtype())
        sales_df_valid = sales_df.dropna(subset=group_keys).copy()

        if sales_df_valid.empty:
            return pd.DataFrame()

        amount_col = (
            "TOTAL_SALES_AMOUNT"
            if "TOTAL_SALES_AMOUNT" in sales_df_valid.columns
            else "AMOUNT"
        )
        if amount_col in sales_df_valid.columns:
            sales_df_valid[amount_col] = pd.to_numeric(
                sales_df_valid[amount_col], errors="coerce"
            ).fillna(0)
            result = sales_df_valid.groupby(group_keys)[amount_col].sum().reset_index()
            result.rename(columns={amount_col: "TOTAL_SALES"}, inplace=True)
            return result

        return pd.DataFrame()

    # Main Processing
    print("\n" + "=" * 70)
    print("ENTERPRISE ARCHITECTURE ANALYSIS WITH SERIAL RECONCILIATION")
    print("=" * 70)

    # Step 0: Serial Number Reconciliation
    print("\n[STEP 0/7] Serial Number Reconciliation")
    serial_reconciliation_df, serial_summary, serial_excel_path = (
        perform_serial_reconciliation()
    )
    messages.append(serial_summary)

    # Step 1: Process customer base
    print("\n[STEP 1/7] Processing customer base")
    base_df = _upper_cols(customer_detail_by_cav_df)

    if "CAV_BU_ID" not in base_df.columns or "PARTY_ID" not in base_df.columns:
        error_msg = "Error: Missing required columns"
        print(error_msg)
        messages.append(error_msg)
        return (
            pd.DataFrame(),
            "\n".join(messages),
            serial_reconciliation_df,
            serial_summary,
            serial_excel_path,
        )

    base_df["CAV_BU_ID"] = pd.to_numeric(base_df["CAV_BU_ID"], errors="coerce").astype(
        pd.Int64Dtype()
    )
    base_df["PARTY_ID"] = pd.to_numeric(base_df["PARTY_ID"], errors="coerce").astype(
        pd.Int64Dtype()
    )
    base_df = base_df.dropna(subset=["CAV_BU_ID", "PARTY_ID"]).copy()

    ## Step 2-4: Aggregate
    print("\n[STEP 2-4/7] Aggregating data")
    ib_agg = _aggregate_ib_from_install_base(install_base_df)
    be_agg = aggregate_ib_by_business_entity_external(install_base_df)
    eol_eos_agg = aggregate_eol_eos_metrics_external(install_base_df)
    case_agg = _aggregate_case_history(case_history_df)
    sales_agg = _aggregate_sales_history(sales_history_df)

    # Step 6: Merge
    print("\n[STEP 6/7] Merging")
    merge_keys = ["CAV_BU_ID", "PARTY_ID"]
    final_df = base_df.copy()

    # CRITICAL FIX: Drop any existing IB count columns from base_df to avoid _x/_y suffixes
    ib_cols_to_drop = [
        "IB_COUNT_COVERED",
        "IB_COUNT_UNCOVERED",
        "IB_COUNT_NEVER_COVERED",
        "IB_COUNT_INACTIVE",
    ]
    for col in ib_cols_to_drop:
        if col in final_df.columns:
            print(f"Dropping duplicate column from base_df: {col}")
            final_df = final_df.drop(columns=[col])

    if not ib_agg.empty:
        print(f"Merging IB data: {len(ib_agg)} parties")
        print(f"IB columns before merge: {list(ib_agg.columns)}")
        print(f"IB sample data:\n{ib_agg.head()}")
        final_df = final_df.merge(ib_agg, on=merge_keys, how="left")
        print(
            f"Columns after IB merge: {[col for col in final_df.columns if 'IB_COUNT' in col]}"
        )
    else:
        print("Warning: IB aggregation is empty - adding placeholder columns")
        final_df["IB_COUNT_COVERED"] = 0
        final_df["IB_COUNT_UNCOVERED"] = 0
        final_df["IB_COUNT_NEVER_COVERED"] = 0

    if not case_agg.empty:
        print(f"Merging Case data: {len(case_agg)} parties")
        final_df = final_df.merge(case_agg, on=merge_keys, how="left")

    if not sales_agg.empty:
        print(f"Merging Sales data: {len(sales_agg)} parties")
        final_df = final_df.merge(sales_agg, on=merge_keys, how="left")

    if not be_agg.empty:
        print(f"Merging BE data: {len(be_agg)} parties")
        final_df = final_df.merge(be_agg, on=merge_keys, how="left")
        if "BE_BREAKDOWN" in final_df.columns:
            final_df["BE_BREAKDOWN"] = final_df["BE_BREAKDOWN"].fillna("")
        if "SUBBE_BREAKDOWN" in final_df.columns:
            final_df["SUBBE_BREAKDOWN"] = final_df["SUBBE_BREAKDOWN"].fillna("")

    if not eol_eos_agg.empty:
        print(f"Merging EOL/EOS data: {len(eol_eos_agg)} parties")
        final_df = final_df.merge(eol_eos_agg, on=merge_keys, how="left")
        eol_eos_cols = [
            "EOL_PASSED",
            "EOL_WITHIN_1YR",
            "EOL_FUTURE",
            "EOL_UNKNOWN",
            "EOS_PASSED",
            "EOS_WITHIN_1YR",
            "EOS_FUTURE",
            "EOS_UNKNOWN",
        ]
        for col in eol_eos_cols:
            if col in final_df.columns:
                final_df[col] = final_df[col].fillna(0)

    # Check IB columns before fillna
    print(
        f"\nBefore fillna - IB_COUNT_COVERED sum: {final_df.get('IB_COUNT_COVERED', pd.Series([0])).sum()}"
    )

    # Fill NaN for key columns
    for col in [
        "IB_COUNT_COVERED",
        "IB_COUNT_UNCOVERED",
        "IB_COUNT_NEVER_COVERED",
        "CASE_COUNT",
        "TOTAL_SALES",
    ]:
        if col in final_df.columns:
            final_df[col] = final_df[col].fillna(0)

    # Check IB columns after fillna
    print(
        f"After fillna - IB_COUNT_COVERED sum: {final_df.get('IB_COUNT_COVERED', pd.Series([0])).sum()}"
    )
    print(f"After fillna - Sample of IB columns:")
    if "IB_COUNT_COVERED" in final_df.columns:
        print(
            final_df[
                [
                    "CAV_BU_ID",
                    "PARTY_ID",
                    "IB_COUNT_COVERED",
                    "IB_COUNT_UNCOVERED",
                    "IB_COUNT_NEVER_COVERED",
                ]
            ].head(10)
        )

    # Ensure IB columns exist before calculating total
    if "IB_COUNT_COVERED" not in final_df.columns:
        final_df["IB_COUNT_COVERED"] = 0
    if "IB_COUNT_UNCOVERED" not in final_df.columns:
        final_df["IB_COUNT_UNCOVERED"] = 0
    if "IB_COUNT_NEVER_COVERED" not in final_df.columns:
        final_df["IB_COUNT_NEVER_COVERED"] = 0

    # FIXED: Calculate TOTAL_IB_COUNT properly using column values
    final_df["TOTAL_IB_COUNT"] = (
        final_df["IB_COUNT_COVERED"].fillna(0)
        + final_df["IB_COUNT_UNCOVERED"].fillna(0)
        + final_df["IB_COUNT_NEVER_COVERED"].fillna(0)
    ).astype(int)

    # Print verification statistics
    # FIXED: Calculate totals using unique PARTY_ID to avoid double-counting
    # Group by PARTY_ID to get unique party-level totals
    party_level_totals = (
        final_df.groupby(["CAV_BU_ID", "PARTY_ID"])
        .agg(
            {
                "IB_COUNT_COVERED": "first",
                "IB_COUNT_UNCOVERED": "first",
                "IB_COUNT_NEVER_COVERED": "first",
                "TOTAL_IB_COUNT": "first",
            }
        )
        .reset_index()
    )

    total_covered = party_level_totals["IB_COUNT_COVERED"].sum()
    total_uncovered = party_level_totals["IB_COUNT_UNCOVERED"].sum()
    total_never = party_level_totals["IB_COUNT_NEVER_COVERED"].sum()
    total_all = party_level_totals["TOTAL_IB_COUNT"].sum()

    unique_parties = party_level_totals["PARTY_ID"].nunique()

    print(f"\n--- Final Merge Verification ---")
    print(f"Total parties in final dataframe: {len(final_df)} (rows)")
    print(f"Unique parties: {unique_parties}")
    print(
        f"Parties with install base: {len(party_level_totals[party_level_totals['TOTAL_IB_COUNT'] > 0])}"
    )
    print(f"Total IB Covered: {total_covered:,.0f}")
    print(f"Total IB Uncovered: {total_uncovered:,.0f}")
    print(f"Total IB Never Covered: {total_never:,.0f}")
    print(f"Grand Total IB: {total_all:,.0f}")

    # Add a note about duplicate party IDs if detected
    if len(final_df) > unique_parties:
        print(
            f"\n⚠️  Note: Some parties appear multiple times (e.g., multiple ship-to addresses)"
        )
        print(
            f"   IB counts above reflect unique party totals to avoid double-counting"
        )

    print(f"\n✓ Analysis complete: {len(final_df)} parties")

    log_output = "\n".join(messages)
    return (
        final_df,
        log_output,
        serial_reconciliation_df,
        serial_summary,
        serial_excel_path,
    )


def save_ea_analysis_to_excel(
    final_df: pd.DataFrame,
    cav_name: str,
    cav_id: int,
    output_folder: str = "analysis_output",
) -> str:
    """Save the EA analysis results to a formatted Excel file."""
    os.makedirs(output_folder, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_cav_name = cav_name.replace(" ", "_").replace("/", "_")
    filename = f"{safe_cav_name}_CAV_{cav_id}_EA_Analysis_{timestamp}.xlsx"
    filepath = os.path.join(output_folder, filename)

    with pd.ExcelWriter(filepath, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, sheet_name="EA Analysis", index=False)

        workbook = writer.book
        worksheet = writer.sheets["EA Analysis"]

        header_format = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#4472C4",
                "font_color": "white",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        )

        for col_num, col_name in enumerate(final_df.columns):
            worksheet.write(0, col_num, col_name, header_format)
            worksheet.set_column(col_num, col_num, 15)

        worksheet.freeze_panes(1, 0)

    print(f"✅ Excel file saved: {filepath}")
    return filepath


# Main execution logic
if __name__ == "__main__":
    SQL_FILE_FOR_GU = "get_customers_by_gu_name.sql"

    ONEDRIVE_SALES_HISTORY_FOLDER = "IB Analyst\\Metro Fire Dept"
    SALES_HISTORY_EXCEL_FILE = "sales_history_data.xlsx"

    ONEDRIVE_PROPOSAL_INPUTS_FOLDER = "IB Analyst\\Metro Fire Dept"
    CUSTOMER_SERIAL_IB_EXCEL_FILE = "customer_serial_ib_data.xlsx"
    EAMP_PROPOSAL_QUOTE_EXCEL_FILE = "eamp_proposal_quote.xlsx"

    ONEDRIVE_CASE_INSTALL_BASE_FOLDER = "IB Analyst\\Metro Fire Dept"
    CASE_HISTORY_EXCEL_FILE = "case_history_by_cav.xlsx"
    INSTALL_BASE_EXCEL_FILE = "install_base_by_cav.xlsx"

    ONEDRIVE_CUSTOMER_DETAIL_FOLDER = "IB Analyst\\Metro Fire Dept"
    CUSTOMER_DETAIL_BY_CAV_EXCEL_FILE = "customer_detail_by_cavid.xlsx"

    target_gu_name = input(
        "Enter the Global Ultimate Customer Name (e.g., 'Amazon'): "
    ).strip()

    if not target_gu_name:
        logging.error("Global Ultimate Customer Name cannot be empty. Exiting.")
        exit(1)

    sanitized_gu_name = "".join(
        c for c in target_gu_name if c.isalnum() or c in (" ", "_")
    ).replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = "data_verification_output"

    print(f"\n--- Fetching Customer Records for '{target_gu_name}' ---")
    customer_records_list = execute_sql_file(
        sql_file_path=SQL_FILE_FOR_GU, bind_params={"gu_name": target_gu_name}
    )

    if customer_records_list:
        print(
            f"Successfully retrieved {len(customer_records_list)} records for '{target_gu_name}'."
        )

        customer_df = pd.DataFrame(customer_records_list)
        customer_df.columns = customer_df.columns.str.upper()
        for col in ["CAV_ID", "CAV_BU_ID", "PARTY_ID"]:
            if col in customer_df.columns:
                customer_df[col] = pd.to_numeric(
                    customer_df[col], errors="coerce"
                ).astype(pd.Int64Dtype())

        available_cav_ids = (
            customer_df["CAV_ID"].unique().tolist()
            if "CAV_ID" in customer_df.columns
            else []
        )

        if not available_cav_ids:
            logging.info("No CAV IDs found.")
            exit(0)

        print(f"\nAvailable CAV IDs: {', '.join(map(str, available_cav_ids))}")
        target_cav_id = input("Enter the CAV ID for detailed analysis: ").strip()

        if not target_cav_id or str(target_cav_id) not in map(str, available_cav_ids):
            logging.error(f"Invalid CAV ID '{target_cav_id}'.")
            exit(1)

        target_cav_id_int = int(target_cav_id)

        print(f"\n--- Loading Data for CAV ID: {target_cav_id} ---")

        case_history_df = load_excel_from_local_onedrive(
            ONEDRIVE_CASE_INSTALL_BASE_FOLDER, CASE_HISTORY_EXCEL_FILE
        )
        if case_history_df is None:
            case_history_df = pd.DataFrame()

        install_base_df = load_excel_from_local_onedrive(
            ONEDRIVE_CASE_INSTALL_BASE_FOLDER, INSTALL_BASE_EXCEL_FILE
        )
        print(
            f"Install Base Data Loaded: {len(install_base_df) if install_base_df is not None else 0} records"
        )
        if install_base_df is None:
            install_base_df = pd.DataFrame()

        sales_history_df = load_excel_from_local_onedrive(
            ONEDRIVE_SALES_HISTORY_FOLDER, SALES_HISTORY_EXCEL_FILE
        )
        if sales_history_df is None:
            sales_history_df = pd.DataFrame()

        customer_serial_ib_df = load_excel_from_local_onedrive(
            ONEDRIVE_PROPOSAL_INPUTS_FOLDER, CUSTOMER_SERIAL_IB_EXCEL_FILE
        )
        if customer_serial_ib_df is None:
            customer_serial_ib_df = pd.DataFrame()

        eamp_proposal_quote_df = load_excel_from_local_onedrive(
            ONEDRIVE_PROPOSAL_INPUTS_FOLDER, EAMP_PROPOSAL_QUOTE_EXCEL_FILE
        )
        if eamp_proposal_quote_df is None:
            eamp_proposal_quote_df = pd.DataFrame()

        customer_detail_by_cav_df = load_excel_from_local_onedrive(
            ONEDRIVE_CUSTOMER_DETAIL_FOLDER, CUSTOMER_DETAIL_BY_CAV_EXCEL_FILE
        )
        if customer_detail_by_cav_df is None:
            customer_detail_by_cav_df = pd.DataFrame()

        print("\n--- All data sets loaded ---")

        # Perform Analysis
        (
            final_summary_df_for_debug,
            ea_analysis_summary_str,
            serial_recon_df,
            serial_log,
            serial_excel_path,
        ) = perform_ea_analysis_with_serial_reconciliation(
            target_cav_id_int=target_cav_id_int,
            customer_detail_by_cav_df=customer_detail_by_cav_df,
            case_history_df=case_history_df,
            install_base_df=install_base_df,
            sales_history_df=sales_history_df,
            customer_serial_ib_df=customer_serial_ib_df,
            eamp_proposal_quote_df=eamp_proposal_quote_df,
            output_folder=output_dir,
        )

        print("\n--- EA Analysis Summary ---")
        print(final_summary_df_for_debug.head())
        print("\n--- Serial Reconciliation Summary ---")
        print(serial_recon_df.head())

        # Save debug summary
        # Save debug summary with multiple sheets
        if not final_summary_df_for_debug.empty:
            # De-duplicate by PARTY_ID for unique party view
            deduplicated_df = (
                final_summary_df_for_debug.groupby(["CAV_BU_ID", "PARTY_ID"])
                .first()
                .reset_index()
            )

            # Create multi-sheet Excel file
            debug_summary_filename = (
                f"{sanitized_gu_name}_CAV_{target_cav_id}_EA_Analysis_{timestamp}.xlsx"
            )
            debug_summary_filepath = os.path.join(output_dir, debug_summary_filename)
        try:
            with pd.ExcelWriter(debug_summary_filepath, engine="xlsxwriter") as writer:
                # Sheet 1: Unique parties (for analysis and AI)
                deduplicated_df.to_excel(
                    writer, sheet_name="Unique Parties", index=False
                )

                # Sheet 2: All locations (for reference with all ship-to addresses)
                final_summary_df_for_debug.to_excel(
                    writer, sheet_name="All Locations", index=False
                )

                # Sheet 3: Summary statistics
                summary_stats = pd.DataFrame(
                    {
                        "Metric": [
                            "Total Unique Parties",
                            "Total Locations/Ship-Tos",
                            "Total Install Base",
                            "Covered Assets",
                            "Uncovered Assets",
                            "Never Covered Assets",
                            "Total Cases",
                            "Total Sales",
                        ],
                        "Value": [
                            len(deduplicated_df),
                            len(final_summary_df_for_debug),
                            int(deduplicated_df["TOTAL_IB_COUNT"].sum()),
                            int(deduplicated_df["IB_COUNT_COVERED"].sum()),
                            int(deduplicated_df["IB_COUNT_UNCOVERED"].sum()),
                            int(deduplicated_df["IB_COUNT_NEVER_COVERED"].sum()),
                            int(deduplicated_df["CASE_COUNT"].sum())
                            if "CASE_COUNT" in deduplicated_df.columns
                            else 0,
                            f"${deduplicated_df['TOTAL_SALES'].sum():,.2f}"
                            if "TOTAL_SALES" in deduplicated_df.columns
                            else "$0.00",
                        ],
                    }
                )
                summary_stats.to_excel(writer, sheet_name="Summary", index=False)

                # Format all sheets with headers
                workbook = writer.book
                header_format = workbook.add_format(
                    {
                        "bold": True,
                        "bg_color": "#4472C4",
                        "font_color": "white",
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                    }
                )

                # Format Unique Parties sheet
                worksheet1 = writer.sheets["Unique Parties"]
                for col_num, col_name in enumerate(deduplicated_df.columns):
                    worksheet1.write(0, col_num, col_name, header_format)
                    worksheet1.set_column(col_num, col_num, 15)
                worksheet1.freeze_panes(1, 0)

                # Format All Locations sheet
                worksheet2 = writer.sheets["All Locations"]
                for col_num, col_name in enumerate(final_summary_df_for_debug.columns):
                    worksheet2.write(0, col_num, col_name, header_format)
                    worksheet2.set_column(col_num, col_num, 15)
                worksheet2.freeze_panes(1, 0)

                # Format Summary sheet
                worksheet3 = writer.sheets["Summary"]
                for col_num, col_name in enumerate(summary_stats.columns):
                    worksheet3.write(0, col_num, col_name, header_format)
                worksheet3.set_column(0, 0, 30)  # Metric column wider
                worksheet3.set_column(1, 1, 20)  # Value column

            print(f"\n✅ EA Analysis saved to: {debug_summary_filepath}")
            print(
                f"   📊 Sheet 1 'Unique Parties': {len(deduplicated_df)} unique parties"
            )
            print(
                f"   📍 Sheet 2 'All Locations': {len(final_summary_df_for_debug)} locations (includes all ship-to addresses)"
            )
            print(f"   📈 Sheet 3 'Summary': Key metrics at a glance")

        except Exception as e:
            logging.error(f"Error creating multi-sheet Excel: {e}")
            # Fallback to simple save
            save_dataframe_to_excel(
                deduplicated_df, debug_summary_filepath, sheet_name="Final Summary"
            )
            print(f"\n⚠️  Saved simplified version to: {debug_summary_filepath}")
        # AI Summarization, Excel, PowerPoint, and Webex Delivery
        ai_input_text = prepare_data_for_openai(
            final_summary_df_for_debug, serial_recon_df, serial_log
        )
        # Use deduplicated dataframe for AI to avoid inflated numbers
        deduplicated_df_for_ai = (
            final_summary_df_for_debug.groupby(["CAV_BU_ID", "PARTY_ID"])
            .first()
            .reset_index()
        )
        ai_input_text = prepare_data_for_openai(
            deduplicated_df_for_ai, serial_recon_df, serial_log
        )
        cisco_ai_client = get_cisco_ai_client()
        if cisco_ai_client:
            logging.info("Sending analysis to AI...")

            # =================================================================
            # STEP 2: REPLACE WITH THIS NEW CODE
            # =================================================================

            # NEW CODE (with custom terminology):
            # ai_generated_summary = get_ai_summary(
            #    ai_client=cisco_ai_client,
            #    input_text=ai_input_text,
            #    max_retries=3,
            #    retry_delay=2
            # )
            ai_generated_summary = summarize_with_cisco_ai(
                ai_client=cisco_ai_client,
                text_to_summarize=ai_input_text,
                prompt_prefix=f"""Analyze this EA data for {target_gu_name} (CAV: {target_cav_id_int}).

IMPORTANT TERMINOLOGY - READ CAREFULLY:
- "Customer Business Units" = the customer's organizational divisions (identified by CAV_BU_NAME)
- "Parties" or "Locations" = individual sites/locations within customer business units
- U_BE = Cisco Business Entity (Cisco's product line/Architecture: Routing, Switching, Security, etc.)
- U_SUBBE = Cisco Sub-Business Entity (more specific Cisco product category)
- "Architecture" = Another term for Cisco Business Entity (U_BE)
- Use "assets" or "devices" instead of generic equipment terms

SERIAL NUMBER RECONCILIATION CONTEXT:
{serial_log}

CRITICAL: The input data includes detailed EOL/EOS metrics BY ARCHITECTURE (U_BE). You MUST analyze and present this data.

ANALYSIS REQUIRED - COMPLETE ALL 8 SECTIONS:

1) Executive Summary
   - MUST START WITH: Total Customer Business Entities, Total Parties/Locations, Total Cisco Business Entities/Architectures (U_BE)
   - Overall portfolio health across customer business entities
   - Total Install Base, Total Cases, Total Sales
   - Serial Reconciliation Match Rate
   - Highlight key statistics and overall assessment

2) Cisco Business Entity (U_BE/Product Architecture) Breakdown ⭐ KEY FOCUS
   - Create a TABLE showing top Cisco Business Entities/Architectures with:
     * Asset counts
     * Percentages
     * Coverage breakdown (Covered/Uncovered/Never Covered)
   - Identify coverage gaps by Cisco Business Entity/Architecture

3) EOL/EOS Risk Analysis by Cisco Product Architecture (U_BE) ⭐⭐ CRITICAL - MANDATORY SECTION
   - **YOU MUST CREATE A DETAILED TABLE** showing EOL/EOS metrics for EACH Architecture
   - **TABLE FORMAT REQUIRED:**
   
   | Architecture | Total Assets | Past EOL | EOL <1Yr | Past EOS | EOS <1Yr | EOL Risk % | Priority |
   |-------------|--------------|----------|----------|----------|----------|------------|----------|
   | Security    | XXX,XXX      | XX,XXX   | X,XXX    | XX,XXX   | X,XXX    | XX.X%      | HIGH     |
   | Collaboration | XX,XXX     | X,XXX    | XXX      | X,XXX    | XXX      | XX.X%      | MEDIUM   |
   
   - **Analysis MUST include:**
     * Which Architectures have highest EOL/EOS risk
     * Risk percentages for each Architecture
     * Priority ranking (HIGH/MEDIUM/LOW) for each Architecture
     * Specific refresh recommendations PER Architecture
   - **This section is MANDATORY - do not skip or summarize lightly**

4) Asset & Coverage Metrics by Customer Business Entity
   - If multiple customer business entities exist, show breakdown by customer BE
   - Asset count, coverage %, cases, sales per customer BE

5) Serial Number Reconciliation ⭐ IMPORTANT
   - Summarize match rate and data quality
   - Interpret customer-only and Cisco-only serials
   - Provide specific recommendations
   - Impact on coverage accuracy

6) Coverage Opportunities by Architecture
   - Uncovered assets by Cisco Business Entity/Product Architecture (U_BE)
   - Revenue potential by product Architecture
   - Prioritized recommendations

7) Top Customer Business Units & Locations
   - Largest customer business units by install base
   - Priority business units for engagement
   - Key locations with highest asset concentrations

8) Next Steps - Actionable Recommendations
   - Specific, prioritized recommendations
   - **MUST include Product Architecture-specific EOL/EOS refresh actions from Section 3**
   - Include serial reconciliation follow-up actions
   - Coverage expansion strategy by Product Architecture

FORMAT REQUIREMENTS:
- Use clear headings with numbering (## 1), ## 2), etc.)
- Section 3 MUST include a detailed table - DO NOT skip this section
- Use markdown tables for better readability
- Executive Summary MUST lead with all key metrics including Serial Match Rate
- Remember U_BE/Architecture = Cisco's product Architecture, not customer organizations""",
            )
            print("\n--- AI Summary ---")
            print(ai_generated_summary)

            # Generate Excel
            # excel_path = save_ea_analysis_to_excel(
            #    final_df=final_summary_df_for_debug,
            #    cav_name=target_gu_name,
            #    cav_id=target_cav_id_int,
            #    output_folder=output_dir
            # )
            # Use the multi-sheet Excel file created earlier
            excel_path = debug_summary_filepath
            # Generate PowerPoint with AI
            ppt_path = create_ea_powerpoint_with_ai(
                final_df=deduplicated_df_for_ai,
                ai_summary=ai_generated_summary,
                ai_client=cisco_ai_client,
                cav_name=target_gu_name,
                cav_id=target_cav_id_int,
                gu_name=target_gu_name,
                output_folder=output_dir,
            )

            # Send to Webex
            webex_message = f"""
# ✅ EA Analysis Complete: {target_gu_name}

## 📊 Key Metrics
- **Parties:** {len(final_summary_df_for_debug):,}
- **Install Base:** {final_summary_df_for_debug.get("TOTAL_IB_COUNT", pd.Series([0])).sum():,.0f}
- **Cases:** {final_summary_df_for_debug.get("CASE_COUNT", pd.Series([0])).sum():,.0f}
- **Sales:** ${final_summary_df_for_debug.get("TOTAL_SALES", pd.Series([0])).sum():,.2f}

## 🔍 Serial Number Reconciliation
- **Total Serials:** {len(serial_recon_df):,}
- **Matched:** {len(serial_recon_df[serial_recon_df["RECONCILIATION_STATUS"] == "Matched"]) if "RECONCILIATION_STATUS" in serial_recon_df.columns else 0:,}
- **Customer Only:** {len(serial_recon_df[serial_recon_df["RECONCILIATION_STATUS"] == "Customer Only"]) if "RECONCILIATION_STATUS" in serial_recon_df.columns else 0:,}
- **Cisco Only:** {len(serial_recon_df[serial_recon_df["RECONCILIATION_STATUS"] == "Cisco Only"]) if "RECONCILIATION_STATUS" in serial_recon_df.columns else 0:,}

## 🤖 AI Insights
{ai_generated_summary}

## 📁 Files Attached
1. Excel Report (EA Analysis)
2. PowerPoint Presentation
3. Serial Number Reconciliation Report
            """
            # if 1 == 2:
            if AUTHORIZATION_TOKEN:
                # Send message with AI summary and Excel file
                send_webex_message_with_files(
                    room_id=WEBEX_ROOM_ID,
                    to_email=FALLBACK_WEBEX_EMAIL,
                    text_message=webex_message.strip(),
                    file_paths=[excel_path],
                    authorization_token=AUTHORIZATION_TOKEN,
                )

                # Send PowerPoint in a separate message
                send_webex_message_with_files(
                    room_id=WEBEX_ROOM_ID,
                    to_email=FALLBACK_WEBEX_EMAIL,
                    text_message="📊 PowerPoint Presentation",
                    file_paths=[ppt_path],
                    authorization_token=AUTHORIZATION_TOKEN,
                )

                # Send Serial Reconciliation Excel in a separate message
                if serial_excel_path and os.path.exists(serial_excel_path):
                    send_webex_message_with_files(
                        room_id=WEBEX_ROOM_ID,
                        to_email=FALLBACK_WEBEX_EMAIL,
                        text_message="🔍 Serial Number Reconciliation Report",
                        file_paths=[serial_excel_path],
                        authorization_token=AUTHORIZATION_TOKEN,
                    )

                print(
                    f"\n✅ Sent to Webex: AI Summary + EA Excel + PowerPoint + Serial Reconciliation"
                )
        else:
            print("⚠️ AI client failed. Generating files only.")
            # print("⚠️ AI client failed. Generating files only.")
            # Use the multi-sheet Excel file created earlier
            excel_path = debug_summary_filepath
            # Use deduplicated data for PowerPoint
            deduplicated_df_for_ai = (
                final_summary_df_for_debug.groupby(["CAV_BU_ID", "PARTY_ID"])
                .first()
                .reset_index()
            )
            ppt_path = create_ea_powerpoint(
                deduplicated_df_for_ai,
                "AI unavailable",
                target_gu_name,
                target_cav_id_int,
                target_gu_name,
                output_dir,
            )
            # excel_path = save_ea_analysis_to_excel(final_summary_df_for_debug, target_gu_name, target_cav_id_int, output_dir)
            # ppt_path = create_ea_powerpoint(final_summary_df_for_debug, "AI unavailable", target_gu_name, target_cav_id_int, target_gu_name, output_dir)
    else:
        print(f"No customer records found for '{target_gu_name}'.")

    print("\n--- Script finished ---")
