import os
import sys
import logging
import pandas as pd
import requests
from tkinter import Tk, Label, Entry, Button, filedialog, Text, Checkbutton, IntVar, StringVar
from tkinter import Frame


class EmployeeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Employee Import")
        self.root.geometry("800x500")  # Adjusted window size

        # Set up logging
        self.setup_logging()

        # Authentication session variable
        self.auth_session = None

        # Create the UI
        self.create_ui()

    def setup_logging(self):
        base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
        log_file_path = os.path.join(base_dir, "script.log")
        logging.basicConfig(
            level=logging.DEBUG,
            format="%(asctime)s - %(levelname)s - %(message)s",
            handlers=[
                logging.FileHandler(log_file_path),
                logging.StreamHandler(sys.stdout)
            ]
        )
        logging.info(f"Log file path: {log_file_path}")

    def create_ui(self):
        # Section 1: Import
        import_frame = Frame(self.root, bd=2, relief="groove", padx=10, pady=10)
        import_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        Label(import_frame, text="Import File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.file_path_var = StringVar()
        self.file_entry = Entry(import_frame, textvariable=self.file_path_var, width=50)
        self.file_entry.grid(row=0, column=1, padx=5, pady=5)
        Button(import_frame, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)

        self.overwrite_var = IntVar()
        Checkbutton(import_frame, text="Overwrite", variable=self.overwrite_var).grid(row=1, column=1, sticky="w", padx=5, pady=5)

        Button(import_frame, text="Import", command=self.process_data).grid(row=2, column=1, padx=5, pady=10)

        # Section 2: Log
        log_frame = Frame(self.root, bd=2, relief="groove", padx=10, pady=10)
        log_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        Label(log_frame, text="Log:").grid(row=0, column=0, padx=5, pady=5, sticky="nw")
        self.log_text = Text(log_frame, width=80, height=15, state="disabled")
        self.log_text.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

    def log_message(self, message):
        logging.info(message)
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.configure(state="disabled")
        self.log_text.see("end")

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.file_path_var.set(file_path)

    def authenticate(self):
        try:
            auth_url = "http://rpsupport1003/v1/rest/auth"
            auth_response = requests.get(auth_url)

            if auth_response.status_code != 200:
                self.log_message(f"Authentication failed. Status Code: {auth_response.status_code}")
                return False

            auth_nonce = int(auth_response.headers.get('Auth-Nonce', 0))
            if not auth_nonce:
                self.log_message("Auth-Nonce not found in the response headers.")
                return False

            auth_nonce_response = (auth_nonce // 13) % 99999 * 17
            auth_with_cred_url = "http://rpsupport1003/v1/rest/auth?usr=sysadmin&pwd=sysadmin"
            headers = {
                "Auth-Nonce": str(auth_nonce),
                "Auth-Nonce-Response": str(auth_nonce_response)
            }
            response = requests.get(auth_with_cred_url, headers=headers)

            if response.status_code != 200:
                self.log_message(f"Authentication failed. Status Code: {response.status_code}")
                return False

            self.auth_session = response.headers.get('Auth-Session')
            if not self.auth_session:
                self.log_message("Auth-Session not found in the response headers.")
                return False

            self.log_message("Authentication successful.")
            return True
        except Exception as e:
            self.log_message(f"Error during authentication: {e}")
            return False

    def process_data(self):
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            self.log_message("Invalid file path. Please select a valid Excel file.")
            return

        # Authenticate automatically if not already authenticated
        if not self.auth_session:
            self.log_message("Authenticating...")
            if not self.authenticate():
                self.log_message("Authentication failed. Cannot proceed with import.")
                return

        self.send_data_to_api(file_path, self.auth_session)

    def send_data_to_api(self, file_path, auth_session):
        try:
            df = pd.read_excel(file_path)
            base_api_url = "http://rpsupport1003/api/common/employee"
            headers = {
                "Auth-Session": auth_session,
                "Accept": "application/json,text/plain, version=2",
                "Content-Type": "application/json"
            }

            for index, row in df.iterrows():
                username = row['USERNAME']
                fullname = row['FULLNAME']
                emplno1 = row['EMPLNO1']
                emplno2 = row['EMPLNO2']
                maxdiscperc = row['MAXDISCPERC']
                newusername = row['NEWUSERNAME']

                name_parts = fullname.split()
                firstname = name_parts[0]
                lastname = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""

                check_url = f"{base_api_url}list?filter=(username,lk,{username})AND(active,eq,true)"
                self.log_message(f"Checking if user {username} exists...")
                response = requests.get(check_url, headers=headers, timeout=10)

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
                        update_url = f"{base_api_url}/{sid}"
                        self.log_message(f"Updating user {username}...")
                        update_response = requests.put(update_url, json=update_payload, headers=headers, timeout=10)
                        if update_response.ok:
                            self.log_message(f"Successfully updated user {username}.")
                        else:
                            self.log_message(f"Failed to update user {username}. Status Code: {update_response.status_code}")
                    else:
                        create_payload = {
                            "data": [{
                                "originapplication": "RProPrismWeb",
                                "active": True,
                                "useractive": True,
                                "exempt": 0,
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
                        self.log_message(f"Creating new user {username}...")
                        create_response = requests.post(base_api_url, json=create_payload, headers=headers, timeout=10)
                        if create_response.ok:
                            self.log_message(f"Successfully created new user {username}.")
                        else:
                            self.log_message(f"Failed to create user {username}. Status Code: {create_response.status_code}")
                else:
                    self.log_message(f"Failed to check user {username}. Status Code: {response.status_code}")

        except Exception as e:
            self.log_message(f"Error in send_data_to_api: {e}")


if __name__ == "__main__":
    root = Tk()
    app = EmployeeApp(root)
    root.mainloop()
