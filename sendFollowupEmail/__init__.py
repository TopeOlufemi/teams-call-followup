
import logging
import azure.functions as func
import requests
from msal import ConfidentialClientApplication
import os

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Processing follow-up email request")

    recipient_email = req.params.get("email")
    if not recipient_email:
        return func.HttpResponse("Missing 'email' parameter", status_code=400)

    client_id = os.environ["CLIENT_ID"]
    client_secret = os.environ["CLIENT_SECRET"]
    tenant_id = os.environ["TENANT_ID"]
    sender_email = os.environ["SENDER_EMAIL"]

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scope = ["https://graph.microsoft.com/.default"]

    app = ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )

    token_result = app.acquire_token_for_client(scopes=scope)

    if "access_token" not in token_result:
        return func.HttpResponse(f"Auth failed: {token_result.get('error_description')}", status_code=500)

    access_token = token_result["access_token"]

    email_body = """
Hi,

I just tried calling you on Teams, but we missed each other.

Here’s what I wanted to share:
- [Insert your message here]

Let me know when’s a good time to connect.

Best,  
[Your Name]
    """

    message = {
        "message": {
            "subject": "Tried to Reach You via Teams",
            "body": {
                "contentType": "Text",
                "content": email_body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient_email
                    }
                }
            ]
        }
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail"
    response = requests.post(url, headers=headers, json=message)

    if response.status_code == 202:
        return func.HttpResponse("Email sent successfully.")
    else:
        return func.HttpResponse(f"Error: {response.text}", status_code=500)
