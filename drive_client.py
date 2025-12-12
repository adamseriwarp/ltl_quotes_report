"""
Google Drive client for accessing LTL quote folders.
Supports both local development (credentials.json/token.pickle) and
Streamlit Cloud deployment (secrets).
"""
import os
import io
import pickle
import json
from pathlib import Path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# If modifying these scopes, delete the token.pickle file
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']


def _get_credentials_from_streamlit_secrets():
    """Load credentials from Streamlit secrets (for cloud deployment)."""
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            # Using service account
            from google.oauth2 import service_account
            return service_account.Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=SCOPES
            )
        elif "google_oauth" in st.secrets:
            # Using OAuth token stored in secrets
            token_info = st.secrets["google_oauth"]
            return Credentials(
                token=token_info.get("token"),
                refresh_token=token_info.get("refresh_token"),
                token_uri=token_info.get("token_uri", "https://oauth2.googleapis.com/token"),
                client_id=token_info.get("client_id"),
                client_secret=token_info.get("client_secret"),
                scopes=SCOPES
            )
    except Exception:
        pass
    return None


class DriveClient:
    def __init__(self, credentials_path: str = 'credentials.json', token_path: str = 'token.pickle'):
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.service = None
        self._authenticate()

    def _authenticate(self):
        """Authenticate with Google Drive API."""
        creds = None

        # Try Streamlit secrets first (for cloud deployment)
        creds = _get_credentials_from_streamlit_secrets()

        if creds is None:
            # Fall back to local file-based auth
            if os.path.exists(self.token_path):
                with open(self.token_path, 'rb') as token:
                    creds = pickle.load(token)

            # If no valid credentials, authenticate
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.credentials_path, SCOPES
                    )
                    creds = flow.run_local_server(port=0)

                # Save the credentials for next run
                with open(self.token_path, 'wb') as token:
                    pickle.dump(creds, token)

        self.service = build('drive', 'v3', credentials=creds)
        print("✓ Connected to Google Drive")
    
    def list_folders(self, parent_id: str = None, name_contains: str = None) -> list:
        """List folders, optionally filtering by parent or name."""
        query_parts = ["mimeType='application/vnd.google-apps.folder'"]
        
        if parent_id:
            query_parts.append(f"'{parent_id}' in parents")
        
        if name_contains:
            query_parts.append(f"name contains '{name_contains}'")
        
        query = " and ".join(query_parts)
        
        results = self.service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name)',
            includeItemsFromAllDrives=True,
            supportsAllDrives=True
        ).execute()
        
        return results.get('files', [])
    
    def list_files_in_folder(self, folder_id: str, file_type: str = None) -> list:
        """List files in a specific folder."""
        query = f"'{folder_id}' in parents and trashed=false"
        
        if file_type:
            query += f" and mimeType contains '{file_type}'"
        
        results = self.service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name, mimeType)',
            includeItemsFromAllDrives=True,
            supportsAllDrives=True
        ).execute()
        
        return results.get('files', [])
    
    def download_file_content(self, file_id: str) -> bytes:
        """Download a file's content as bytes."""
        request = self.service.files().get_media(fileId=file_id)
        file_content = io.BytesIO()
        downloader = MediaIoBaseDownload(file_content, request)
        
        done = False
        while not done:
            status, done = downloader.next_chunk()
        
        return file_content.getvalue()
    
    def search_folders(self, name_pattern: str) -> list:
        """Search for folders by name pattern."""
        query = f"mimeType='application/vnd.google-apps.folder' and name contains '{name_pattern}'"
        
        results = self.service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name)',
            includeItemsFromAllDrives=True,
            supportsAllDrives=True
        ).execute()
        
        return results.get('files', [])


if __name__ == "__main__":
    # Test the connection
    client = DriveClient()
    
    # Search for quote folders
    print("\nSearching for 'Quotes' folders...")
    folders = client.search_folders("Quotes")
    
    for folder in folders[:10]:  # Show first 10
        print(f"  📁 {folder['name']} (ID: {folder['id']})")

