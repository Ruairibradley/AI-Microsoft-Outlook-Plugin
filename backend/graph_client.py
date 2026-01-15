import requests
from typing import Any, Dict, List

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

def _headers(access_token: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {access_token}"}

def list_folders(access_token: str) -> List[Dict[str, Any]]:
    url = f"{GRAPH_ROOT}/me/mailFolders?$top=200&$select=id,displayName,parentFolderId,childFolderCount,totalItemCount"
    r = requests.get(url, headers=_headers(access_token), timeout=30)
    r.raise_for_status()
    return r.json().get("value", [])

def list_messages(access_token: str, folder_id: str, top: int = 25) -> Dict[str, Any]:
    select_fields = "id,subject,bodyPreview,webLink,receivedDateTime,from"
    url = (
        f"{GRAPH_ROOT}/me/mailFolders/{folder_id}/messages"
        f"?$top={top}&$select={select_fields}&$orderby=receivedDateTime desc"
    )
    r = requests.get(url, headers=_headers(access_token), timeout=30)
    r.raise_for_status()
    return r.json()

def get_messages_by_ids(access_token: str, ids: List[str]) -> List[Dict[str, Any]]:
    out = []
    for mid in ids:
        url = f"{GRAPH_ROOT}/me/messages/{mid}?$select=id,subject,bodyPreview,webLink,receivedDateTime,from"
        r = requests.get(url, headers=_headers(access_token), timeout=30)
        r.raise_for_status()
        out.append(r.json())
    return out
