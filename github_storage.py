"""
Luu va tai file tu GitHub repository thay vi Google Drive.
"""
import os
import base64
import logging
import requests
from urllib.parse import quote

GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO = "dobangnhi-gif/dobangnhi"
GITHUB_BRANCH = "main"
GITHUB_DATA_PATH = "data"
API_BASE = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_DATA_PATH}"


def _headers():
    return {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }


def upload_file(local_path, filename=None):
    """Upload file len GitHub."""
    if not GITHUB_TOKEN:
        return False
    if filename is None:
        filename = os.path.basename(local_path)
    url = f"{API_BASE}/{quote(filename)}"
    with open(local_path, "rb") as f:
        content = base64.b64encode(f.read()).decode()
    # Lay SHA neu file da ton tai
    sha = None
    r = requests.get(url, headers=_headers())
    if r.status_code == 200:
        sha = r.json().get("sha")
    payload = {
        "message": f"Update {filename}",
        "content": content,
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    r = requests.put(url, json=payload, headers=_headers())
    if r.status_code in [200, 201]:
        logging.info(f"GitHub upload OK: {filename}")
        return True
    else:
        logging.error(f"GitHub upload ERROR {r.status_code}: {r.text[:200]}")
        return False


def download_file(filename, local_path):
    """Download file tu GitHub ve local. Tra ve True neu thanh cong."""
    if not GITHUB_TOKEN:
        return False
    url = f"{API_BASE}/{quote(filename)}"
    r = requests.get(url, headers=_headers())
    if r.status_code != 200:
        return False
    content = base64.b64decode(r.json()["content"])
    os.makedirs(os.path.dirname(local_path) if os.path.dirname(local_path) else ".", exist_ok=True)
    with open(local_path, "wb") as f:
        f.write(content)
    logging.info(f"GitHub download OK: {filename}")
    return True
