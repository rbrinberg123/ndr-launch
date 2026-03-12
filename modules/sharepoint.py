import os
import requests


class SharePointClient:
    def __init__(self):
        self.tenant_id    = os.environ['AZURE_TENANT_ID']
        self.client_id    = os.environ['AZURE_CLIENT_ID']
        self.client_secret = os.environ['AZURE_CLIENT_SECRET']
        self.site_id      = os.environ['SHAREPOINT_SITE_ID']
        self.folder_path  = os.environ.get('SHAREPOINT_FOLDER', '/NDR Launch')
        self._token       = None

    def _get_token(self):
        if self._token:
            return self._token
        url = f'https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token'
        resp = requests.post(url, data={
            'grant_type':    'client_credentials',
            'client_id':     self.client_id,
            'client_secret': self.client_secret,
            'scope':         'https://graph.microsoft.com/.default',
        })
        resp.raise_for_status()
        self._token = resp.json()['access_token']
        return self._token

    def upload_file(self, file_bytes, filename):
        token = self._get_token()
        headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/octet-stream'}

        folder = self.folder_path.strip('/')
        upload_url = (
            f'https://graph.microsoft.com/v1.0/sites/{self.site_id}'
            f'/drive/root:/{folder}/{filename}:/content'
        )
        resp = requests.put(upload_url, headers=headers, data=file_bytes)
        resp.raise_for_status()

        data = resp.json()
        return data.get('webUrl', '')

    def is_configured(self):
        return all([
            os.environ.get('AZURE_TENANT_ID'),
            os.environ.get('AZURE_CLIENT_ID'),
            os.environ.get('AZURE_CLIENT_SECRET'),
            os.environ.get('SHAREPOINT_SITE_ID'),
        ])
