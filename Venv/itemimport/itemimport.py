import pandas as pd
import requests
import os
import sys
import logging
import json

# Determine the log file path
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

log_file_path = os.path.join(base_dir, "item_import.log")

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_file_path, mode='w'),  # Overwrite log file on each run
        logging.StreamHandler(sys.stdout)
    ]
)

logging.info("Script started.")

# Define the file path for the Excel file
file_name = "Items.xlsx"
file_path = os.path.join(base_dir, file_name)

# Load the Excel file
def load_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        logging.info("Loaded Item Data:")
        logging.info(df.head())
        return df
    except Exception as e:
        logging.error(f"Error loading Excel file: {e}")
        sys.exit(1)

# Authentication function
def authenticate():
    try:
        auth_url = "http://rpsupport1003/v1/rest/auth"
        auth_response = requests.get(auth_url)

        if auth_response.status_code != 200:
            logging.error(f"Failed to get Auth-Nonce. HTTP Status Code: {auth_response.status_code}")
            sys.exit(1)

        auth_nonce = int(auth_response.headers.get("Auth-Nonce", 0))
        if auth_nonce == 0:
            logging.error("Auth-Nonce not found in response headers.")
            sys.exit(1)

        auth_nonce_response = (auth_nonce // 13) % 99999 * 17
        logging.info(f"Generated Auth-Nonce-Response: {auth_nonce_response}")

        auth_with_cred_url = "http://rpsupport1003/v1/rest/auth?usr=sysadmin&pwd=sysadmin"
        headers = {
            "Auth-Nonce": str(auth_nonce),
            "Auth-Nonce-Response": str(auth_nonce_response)
        }
        response = requests.get(auth_with_cred_url, headers=headers)

        if response.status_code != 200:
            logging.error(f"Authentication failed. HTTP Status Code: {response.status_code}")
            sys.exit(1)

        auth_session = response.headers.get("Auth-Session", None)
        if not auth_session:
            logging.error("Auth-Session not found in response headers.")
            sys.exit(1)

        logging.info(f"Authenticated successfully. Auth-Session: {auth_session}")
        return auth_session

    except Exception as e:
        logging.error(f"Error in authenticate: {e}")
        sys.exit(1)

# Function to send data to API
def send_data_to_api(df, auth_session, lookup_field):
    logging.info("Starting data upload process.")
    base_api_url = "http://rpsupport1003/api/backoffice/inventory?action=InventorySaveItems"

    # Log DataFrame columns for debugging
    logging.debug(f"DataFrame columns: {df.columns}")

    for index, row in df.iterrows():
        try:
            lookup_value = row.get(lookup_field, None)
            if pd.isna(lookup_value):
                logging.warning(f"Lookup value is missing for row {index + 1}. Skipping.")
                continue

            # Construct payload based on expected structure
            payload = {
                "data": [
                    {
                        "OriginApplication": "RProPrismWeb",
                        "PrimaryItemDefinition": {
                            "sid": "731360302000125768",
                            "dcssid": "731141629000187642"
                        },
                        "InventoryItems": [
                            {
                                "sbssid": "726950516000113256",
                                "udf1string": row.get("UDF1_STRING", ""),
                                "sid": "731360302000125768"
                            }
                        ],
                        "UpdateStyleDefinition": False,
                        "UpdateStyleCost": False,
                        "UpdateStylePrice": False
                    }
                ]
            }

            # Log payload for debugging
            logging.debug(f"Payload for row {index + 1}: {json.dumps(payload, indent=2)}")

            # Send API request
            headers = {
                "Auth-Session": auth_session,
                "Accept": "application/json,text/plain, version=2",
                "Content-Type": "application/json"
            }

            response = requests.post(base_api_url, json=payload, headers=headers)

            # Handle API response
            if response.ok:
                logging.info(f"Successfully uploaded data for Lookup Value: {lookup_value}")
            else:
                logging.warning(f"Failed to upload data for Lookup Value: {lookup_value}. "
                                f"Status Code: {response.status_code}. Response: {response.text}")

        except Exception as e:
            logging.error(f"Error processing row {index + 1}: {e}")

# Main execution
if __name__ == "__main__":
    try:
        # Load Excel file
        df = load_excel(file_path)

        # Get lookup field from user
        print("Choose a lookup field: ALU, UPC, or Description1")
        lookup_field = input("Enter your choice: ").strip()

        if lookup_field not in df.columns:
            logging.error(f"Invalid lookup field: {lookup_field}")
            sys.exit(1)

        # Authenticate
        session = authenticate()

        # Send data if authentication is successful
        if session:
            send_data_to_api(df, session, lookup_field)

        logging.info("Script finished successfully.")

    except Exception as e:
        logging.error(f"Unexpected error: {e}")