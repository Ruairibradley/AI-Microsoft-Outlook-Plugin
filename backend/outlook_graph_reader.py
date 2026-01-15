# import os
# import requests
# from msal import PublicClientApplication

# CLIENT_ID = "81dcae7c-55a9-4e4e-86e1-e364b25fc990"
# TENANT_ID = "common"
# AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
# SCOPES = ["User.Read", "Mail.Read"]

# def get_access_token():
#     app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

#     accounts = app.get_accounts()
#     if accounts:
#         result = app.acquire_token_silent(SCOPES, account=accounts[0])
#         if result and "access_token" in result:
#             return result["access_token"]

#     flow = app.initiate_device_flow(scopes=SCOPES)
#     print(flow["message"])
#     result = app.acquire_token_by_device_flow(flow)
#     return result["access_token"]

# def fetch_outlook_messages(limit=50):
#     token = get_access_token()
#     headers = {"Authorization": f"Bearer {token}"}

#     messages = []
#     select_fields = "subject,bodyPreview,body,webLink,conversationId,receivedDateTime,from"
#     next_url = (
#         "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages"
#         f"?$top={min(limit,50)}&$select={select_fields}"
#     )

#     while next_url and len(messages) < limit:
#         res = requests.get(next_url, headers=headers)
#         res.raise_for_status()
#         data = res.json()
#         messages.extend(data.get("value", []))
#         next_url = data.get("@odata.nextLink")

#     return messages[:limit]
