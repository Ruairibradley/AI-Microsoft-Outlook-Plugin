import requests
from typing import Any, Dict, List, Optional
from urllib.parse import urlparse

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


def list_messages_page(access_token: str, folder_id: Optional[str], top: int = 25, next_link: Optional[str] = None) -> Dict[str, Any]:
    """
    Returns a single page of messages plus @odata.nextLink if present.
    If next_link is provided, it must be a Microsoft Graph URL (SSRF-safe check).
    """
    if next_link:
        parsed = urlparse(next_link)
        if parsed.scheme != "https" or parsed.netloc.lower() != "graph.microsoft.com":
            raise ValueError("Invalid next_link host.")
        url = next_link
    else:
        if not folder_id:
            raise ValueError("folder_id is required when next_link is not provided.")
        select_fields = "id,subject,bodyPreview,webLink,receivedDateTime,from"
        url = (
            f"{GRAPH_ROOT}/me/mailFolders/{folder_id}/messages"
            f"?$top={top}&$select={select_fields}&$orderby=receivedDateTime desc"
        )

    r = requests.get(url, headers=_headers(access_token), timeout=30)
    r.raise_for_status()
    return r.json()


def get_message_weblink(access_token: str, message_id: str) -> str:
    """
    Fetches a fresh webLink for a message id.
    This addresses cases where stored webLink becomes stale after moves.
    """
    url = f"{GRAPH_ROOT}/me/messages/{message_id}?$select=webLink"
    r = requests.get(url, headers=_headers(access_token), timeout=30)
    r.raise_for_status()
    data = r.json()
    return data.get("webLink") or ""
