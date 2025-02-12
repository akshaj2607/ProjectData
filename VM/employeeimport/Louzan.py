import os
import sys
import logging
import pandas as pd
import requests
from logging.handlers import RotatingFileHandler

# Configure logging
log_file_path = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), "script.log")
handler = RotatingFileHandler(log_file_path, maxBytes=5*1024*1024, backupCount=3)
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[handler, logging.StreamHandler(sys.stdout)]
)
logging.info(f"Log file path: {log_file_path}")
logging.info("Script started.")

# Define file paths
base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
file_name = "Employee.xlsx"
file_path = os.path.join(base_dir, file_name)

# Function to display employee details
def display_employee_details(file_path):
    df = pd.read_excel(file_path)
    for index, row in df.iterrows():
        logging.info(f"Employee {index + 1}: {row.to_dict()}")

# Authentication function
def authenticate():
    try:
        auth_url = "http://prism-pos-lab/v1/rest/auth"
        auth_response = requests.get(auth_url)
        logging.info(f"Auth response: {auth_response.status_code} - {auth_response.text}")
        if auth_response.status_code != 200:
            logging.error(f"Failed to get Auth-Nonce. HTTP Status Code: {auth_response.status_code}")
            return None

        auth_nonce = int(auth_response.headers.get('Auth-Nonce', 0))
        auth_nonce_response = (auth_nonce // 13) % 99999 * 17
        logging.info(f"Generated Auth-Nonce-Response: {auth_nonce_response}")

        auth_with_cred_url = "http://prism-pos-lab/v1/rest/auth?usr=sysadmin&pwd=sysadmin1"
        headers = {
            "Auth-Nonce": str(auth_nonce),
            "Auth-Nonce-Response": str(auth_nonce_response)
        }
        response = requests.get(auth_with_cred_url, headers=headers)
        logging.info(f"Auth with credentials response: {response.status_code} - {response.text}")
        if response.status_code != 200:
            logging.error(f"Authentication failed. HTTP Status Code: {response.status_code}")
            return None

        auth_session = response.headers.get('Auth-Session', None)
        if not auth_session:
            logging.error("Auth-Session not found in the response headers.")
            return None

        logging.info(f"Authenticated successfully. Auth-Session: {auth_session}")
        return auth_session
    except Exception as e:
        logging.error(f"Error in authenticate: {e}")
        return None

# Function to send data to API
def send_data_to_api(file_path, auth_session):
    df = pd.read_excel(file_path)
    base_api_url = "http://prism-pos-lab/api/common/employee"

    for index, row in df.iterrows():
        username = row['USERNAME']
        newusername = row['NEWUSERNAME']
        fullname = row['FULLNAME']
        emplno1 = row['EMPLNO1']
        emplno2 = row['EMPLNO2']
        maxdiscperc = row['MAXDISCPERC']

        name_parts = fullname.split()
        firstname = name_parts[0]
        lastname = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""

        check_url = f"http://prism-pos-lab/api/common/employeelist?filter=(username,lk,{username})AND(active,eq,true)"
        headers = {
            "Auth-Session": auth_session,
            "Accept": "application/json,text/plain, version=2"
        }

        try:
            logging.info(f"Checking user {username} at URL: {check_url}")
            response = requests.get(check_url, headers=headers, timeout=10)
            logging.info(f"Check response: {response.status_code} - {response.text}")

            if response.ok:
                response_data = response.json()
                if 'data' in response_data and len(response_data['data']) > 0:
                    sid = response_data['data'][0]['sid']
                    rowversion = response_data['data'][0]['rowversion']
                    update_payload = {
                        "data": [{
                            "emplno1": emplno1,
                            "emplno2": emplno2,
                            "fullname": fullname,
                            "firstname": firstname,
                            "lastname": lastname,
                            "username": newusername,
                            "maxdiscperc": maxdiscperc,
                            "rowversion": rowversion
                        }]
                    }
                    update_response = requests.put(f"{base_api_url}/{sid}", json=update_payload, headers=headers, timeout=10)
                    logging.info(f"Update response for {username}: {update_response.status_code} - {update_response.text}")
                else:
                    create_payload = {
                        "data": [{
                            "originapplication": "RProPrismWeb",
                            "active": True,
                            "useractive": True,
                            "maxdiscperc": maxdiscperc,
                            "emplno1": emplno1,
                            "emplno2": emplno2,
                            "fullname": fullname,
                            "firstname": firstname,
                            "lastname": lastname,
                            "username": newusername,
                            "status": 1
                        }]
                    }
                    create_response = requests.post(base_api_url, json=create_payload, headers=headers, timeout=10)
                    logging.info(f"Create response for {username}: {create_response.status_code} - {create_response.text}")
            else:
                logging.error(f"Failed to check user {username}. HTTP Status Code: {response.status_code} - {response.text}")
        except Exception as e:
            logging.error(f"Error processing user {username}: {e}")

# Main execution
if __name__ == '__main__':
    try:
        logging.info("Starting script.")
        display_employee_details(file_path)
        session = authenticate()
        if session:
            send_data_to_api(file_path, session)
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
    finally:
        logging.info("Script finished.")
