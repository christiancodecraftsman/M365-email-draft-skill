# /// script
# dependencies = [
#   "msal",
#   "requests",
#   "python-dotenv",
# ]
# ///

import os
import json
import msal
import requests

from dotenv import load_dotenv
import argparse
import logging

# Configure the logger
logging.basicConfig(
    level=logging.ERROR,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("m365_email_draft_skill.log"), # Writes to a file
        logging.StreamHandler()                # Writes to your terminal
    ]
)

class M365Client:
    # NOTE: cache_path: defaults to the same folder as the skill, so multiple applications w/ cache won't overwrite each other
    def __init__(self, client_id, tenant_id, scopes, cache_path="./m365_cache.bin"):
        self.client_id = client_id
        self.scopes = scopes
        self.cache_path = cache_path

        # Initialize Logger
        self.logger = logging.getLogger("M365EmailDraftClient")

        # Initialize and reload MSAL cache
        self.cache = msal.SerializableTokenCache()
        if os.path.exists(self.cache_path):
            with open(self.cache_path, "r") as f:
                self.cache.deserialize(f.read())

        # Initialize MSAL
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.app = msal.PublicClientApplication(
            self.client_id, 
            authority=authority, 
            token_cache=self.cache
        )

    def _save_cache(self):
        """Writes the in-memory cache back to the disk only if it has been updated."""
        if self.cache.has_state_changed:
            # Write cache to file and mark it as read only
            with open(self.cache_path, "w") as f:
                f.write(self.cache.serialize())
            os.chmod(self.cache_path, 0o600)
            
            # Reset the cache changed flag
            self.cache.has_state_changed = False
            self.logger.info(f"cache saved to /{self.cache_path}")

    def get_token(self):
        # Load accounts from cache if it exists, w/ offline_access this should be 90-day from previous use
        accounts = self.app.get_accounts()
        result = None
        if accounts:
            result = self.app.acquire_token_silent(self.scopes, account=accounts[0])
            self._save_cache()
            self.logger.info("User authenticated from cache")
            if result:
                return result["access_token"]
            else:
                self.logger.info("Result was empty?")

        # Fail w/ a message for the LLM to pass along to the user
        print(json.dumps({
            "status": "error",
            "error_type": "AUTH_REQUIRED",
            "message": "Your login session has expired. Please run 'main.py --login' on the VM to reconnect me."
        }))
        self.logger.error("User not authenticated, directed to login tool")
        exit()

    def launch_auth_flow(self):
        flow = self.app.initiate_device_flow(self.scopes)
        print(flow['message'])
        self.app.acquire_token_by_device_flow(flow)
        self._save_cache()
        print("Congrats you authenticated. Your token should be valid for 90-days from the last time it's used.")
        self.logger.info("User authenticated from flow")
        exit()

    def create_draft(self, subject, body, to, cc, bcc):
        token = self.get_token()
        url = "https://graph.microsoft.com/v1.0/me/messages"
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

        payload = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": format_recipients(to),
        }

        if cc:
            payload["ccRecipients"] = format_recipients(cc)
        if bcc:
            payload["bccRecipients"] = format_recipients(bcc)

        self.logger.info(f"Email draft payload generated:\npayload: /{payload}")
        
        response = requests.post(url, headers=headers, json=payload)
        json = response.json()
        self.logger.info(f"Graph API Response: /{json}")

        return json

def format_recipients(to_input: str | list[str]) -> list[dict]:
    """
    Normalizes email input (string or list) and returns the 
    nested JSON structure required by Microsoft Graph API.
    """
    # 1. Ensure we are working with a list
    if isinstance(to_input, str):
        to_input = [to_input]
    
    # 2. Extract and clean every individual email
    clean_emails = []
    for entry in to_input:
        # Replace commas with spaces, then split into individual words
        parts = entry.replace(',', ' ').split()
        clean_emails.extend(parts)
    
    # 3. Construct the nested objects
    # Result: [{"emailAddress": {"address": "email@example.com"}}, ...]
    return [
        {"emailAddress": {"address": email.strip()}}
        for email in clean_emails
        if "@" in email  # Basic validation to skip junk
    ]

if __name__ == "__main__":
    # Load environment from .env file
    load_dotenv()
    client_id = os.environ.get("AZURE_CLIENT_ID")
    tenant_id = os.environ.get("AZURE_TENANT_ID")

    # NOTE: offline_access should not be included in the scope (it's automatic if it's configured on the app)
    scopes = ["User.Read", "Mail.ReadWrite"]
    client = M365Client(client_id, tenant_id, scopes)

    # Setup command line arg parser
    parser = argparse.ArgumentParser(description="M365 Agent Tool")
    parser.add_argument("--login", action="store_true", help="Manual interactive login")
    parser.add_argument("--to", nargs='+', help="One or more recipient email addresses") # space seperate list
    parser.add_argument("--cc", nargs='+', help="carbon copy") # space seperate list
    parser.add_argument("--bcc", nargs='+', help="blind carbon copy") # space seperate list
    parser.add_argument("--subject", type=str, default="New Draft from AI", help="Email subject")
    parser.add_argument("--body", type=str, help="Email body (HTML supported)")
    args = parser.parse_args()

    # Seperate path for login
    if args.login:
        client.launch_auth_flow()

    # Seperate path for creating a draft (requires to and body)
    if args.to and args.body:
        res = client.create_draft(
            args.subject, 
            args.body, 
            args.to,
            args.cc,
            args.bcc
        )

        print("Draft successfully created!")
        print(json.dumps(res))
    else:
        print("Error: Missing required arguments --to or --body")
        parser.print_help()
