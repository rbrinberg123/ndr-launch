import os, requests

class SharePointClient:
    def __init__(self):
        self.tenant_id     = os.environ.get('AZURE_TENANT_ID','')
        self.client_id     = os.environ.get('AZURE_CLIENT_ID','')
        self.client_secret = os.environ.get('AZURE_CLIENT_SECRET','')
        self.site_id       = os.environ.get('SHAREPOINT_SITE_ID','')
        self.folder_path   = os.environ.get('SHAREPOINT_FOLDER','/NDR Launch')
        self._token        = None

    def is_configured(self):
        return all([self.tenant_id, self.client_id, self.client_secret, self.site_id])

    def _get_token(self):
        if self._token: return self._token
        r = requests.post(
            f'https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token',
            data={'grant_type':'client_credentials','client_id':self.client_id,
                  'client_secret':self.client_secret,
                  'scope':'https://graph.microsoft.com/.default'})
        r.raise_for_status()
        self._token = r.json()['access_token']
        return self._token

    def upload_file(self, file_bytes, filename):
        token  = self._get_token()
        folder = self.folder_path.strip('/')
        r = requests.put(
            f'https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root:/{folder}/{filename}:/content',
            headers={'Authorization':f'Bearer {token}','Content-Type':'application/octet-stream'},
            data=file_bytes)
        r.raise_for_status()
        return r.json().get('webUrl','')
