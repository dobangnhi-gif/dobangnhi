"""
Helper module de luu va doc file tu Google Drive.
Can co GOOGLE_CREDENTIALS env variable chua JSON key cua service account.
"""

import os
import io
import json
import logging
from datetime import datetime

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]
FOLDER_NAME = "IndonBot_Data"


def get_service():
    """Tao Google Drive service tu credentials."""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise ValueError("Chua co GOOGLE_CREDENTIALS trong environment variables")
    creds_data = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(creds_data, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def get_or_create_folder(service, folder_name=FOLDER_NAME):
    """Lay hoac tao thu muc tren Drive."""
    query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]
    # Tao moi
    meta = {"name": folder_name, "mimeType": "application/vnd.google-apps.folder"}
    folder = service.files().create(body=meta, fields="id").execute()
    return folder["id"]


def upload_file(local_path, drive_filename=None, folder_id=None):
    """Upload file len Drive, tra ve file_id."""
    service = get_service()
    if folder_id is None:
        folder_id = get_or_create_folder(service)
    if drive_filename is None:
        drive_filename = os.path.basename(local_path)

    # Xoa file cu cung ten neu co
    query = f"name='{drive_filename}' and '{folder_id}' in parents and trashed=false"
    old = service.files().list(q=query, fields="files(id)").execute().get("files", [])
    for f in old:
        service.files().delete(fileId=f["id"]).execute()

    meta = {"name": drive_filename, "parents": [folder_id]}
    with open(local_path, "rb") as fh:
        media = MediaIoBaseUpload(fh, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        f = service.files().create(body=meta, media_body=media, fields="id").execute()
    logging.info(f"Uploaded {drive_filename} -> Drive id={f['id']}")
    return f["id"]


def download_file(drive_filename, local_path, folder_id=None):
    """Download file tu Drive ve local. Tra ve True neu thanh cong."""
    service = get_service()
    if folder_id is None:
        folder_id = get_or_create_folder(service)

    query = f"name='{drive_filename}' and '{folder_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    if not files:
        return False

    file_id = files[0]["id"]
    request = service.files().get_media(fileId=file_id)
    os.makedirs(os.path.dirname(local_path) if os.path.dirname(local_path) else ".", exist_ok=True)
    with open(local_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    logging.info(f"Downloaded {drive_filename} from Drive")
    return True


def file_exists_on_drive(drive_filename, folder_id=None):
    """Kiem tra file co ton tai tren Drive khong."""
    service = get_service()
    if folder_id is None:
        folder_id = get_or_create_folder(service)
    query = f"name='{drive_filename}' and '{folder_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    return len(results.get("files", [])) > 0


def list_files_on_drive(folder_id=None):
    """Liet ke tat ca file trong thu muc Drive."""
    service = get_service()
    if folder_id is None:
        folder_id = get_or_create_folder(service)
    query = f"'{folder_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id, name, modifiedTime)").execute()
    return results.get("files", [])
