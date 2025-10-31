import os
import requests
from msal import PublicClientApplication

# CONFIG
CLIENT_ID = "81dcae7c-55a9-4e4e-86e1-e364b25fc990"
TENANT_ID = "common"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Mail.Read"]


OUTPUT_DIR = "./data/outlook_emails"


def get_access_token():
    app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

    # Try silent login first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    # Device flow for first-time login
    flow = app.initiate_device_flow(scopes=SCOPES)
    print(flow["message"])  # Shows code + link to log in
    result = app.acquire_token_by_device_flow(flow)
    return result["access_token"]


def export_outlook_web_emails(limit=50):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    messages = []
    next_url = f"https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages?$top={limit}&$select=subject,bodyPreview,body"


    # Fetch multiple pages if necessary
    while next_url and len(messages) < limit:
        res = requests.get(next_url, headers=headers)
        res.raise_for_status()
        data = res.json()
        messages.extend(data.get("value", []))
        next_url = data.get("@odata.nextLink")

    count = 0
    for msg in messages:
        subject = msg.get("subject", "No Subject")
        body_info = msg.get("body", {})
        body_content = body_info.get("content", "")
        content_type = body_info.get("contentType", "")

        if content_type.lower() == "html":
            import re
            body_content = re.sub(r"<[^>]+>", "", body_content)

        if not body_content.strip():
            body_content = msg.get("bodyPreview", "")

        filename = os.path.join(OUTPUT_DIR, f"email_{count+1:04}.txt")
        with open(filename, "w", encoding="utf-8", errors="ignore") as f:
            f.write(f"Subject: {subject}\n\n{body_content.strip()}")
        count += 1

    print(f" Exported {count} Outlook Web emails to {OUTPUT_DIR}")



if __name__ == "__main__":
    export_outlook_web_emails(limit=50)
