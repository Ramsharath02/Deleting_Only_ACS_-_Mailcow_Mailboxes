import csv
import requests
import subprocess
import pandas as pd
from datetime import datetime

# === Configuration ===
MAILCOW_API_KEY = "134C38-0CBB59-49C117-219FB7-A7BC5B"
MAILCOW_URL = "https://mail6.atozwriter.com"
ACS_EMAIL_SERVICE = "acs-mailcow-domain01"
ACS_RESOURCE_GROUP = "acs-mailcow-rg"

# === For Logging ===
logs = []

# === Functions ===

def delete_mailcow_mailbox(email):
    url = f"{MAILCOW_URL}/api/v1/delete/mailbox"
    headers = {
        "Content-Type": "application/json",
        "X-API-Key": MAILCOW_API_KEY
    }
    payload = [email]  # List of mailbox emails to delete
    
    response = requests.post(url, json=payload, headers=headers)
    if response.status_code == 200:
        return True, None
    else:
        return False, response.text

def delete_acs_sender(domain_name, sender_username):
    try:
        delete_command = [
            "az", "communication", "email", "domain", "sender-username", "delete",
            "--domain-name", domain_name,
            "--email-service-name", ACS_EMAIL_SERVICE,
            "--name", sender_username,
            "--resource-group", ACS_RESOURCE_GROUP,
            "--yes"
        ]
        subprocess.run(delete_command, check=True, timeout=60)
        return True, None
    except subprocess.CalledProcessError as e:
        return False, str(e)
    except subprocess.TimeoutExpired:
        return False, "Command timed out."

# === Main Process ===
def main():
    try:
        with open("emails.csv", newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                email = row["Email"].strip()
                domain = row["Domain"].strip()
                username = row["Username"].strip()

                # --- Delete Mailcow Mailbox ---
                mailbox_success, mailbox_error = delete_mailcow_mailbox(email)

                # --- Delete ACS Sender Username ---
                sender_success, sender_error = delete_acs_sender(domain, username)

                # --- Log the result ---
                logs.append({
                    "Email": email,
                    "Username": username,
                    "Mailbox Deletion Status": "Success" if mailbox_success else f"Failed: {mailbox_error}",
                    "ACS Sender Deletion Status": "Success" if sender_success else f"Failed: {sender_error}",
                    "Timestamp": datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
                })

        # === Save Logs to Excel ===
        df = pd.DataFrame(logs)
        df.to_excel("deletion_log_mailcow_acs.xlsx", index=False)
        print("✅ Deletion report saved to: deletion_log_mailcow_acs.xlsx")

    except Exception as e:
        print(f"❌ Error occurred during deletion: {str(e)}")

# === Entrypoint ===
if __name__ == "__main__":
    main()
