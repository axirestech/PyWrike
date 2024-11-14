from PyWrike.gateways.basegateway1 import APIGateway
import requests
import json
import webbrowser
import threading
import socketserver
import http.server
import re

class OAuth2CodeServer(http.server.BaseHTTPRequestHandler):
    # Server setup remains as-is

class QuickSocketServer(socketserver.TCPServer):
    # Socket server setup remains as-is

class ServerThread(threading.Thread):
    # Thread setup remains as-is

class OAuth2Gateway1(APIGateway):
    def __init__(self, client_id=None, client_secret=None, redirect_uri=None, tokens_updater=None, wait_for_redirect=False):
        super().__init__()
        self._client_id = client_id
        self._client_secret = client_secret
        self._redirect_uri = redirect_uri
        self._tokens_updater = tokens_updater
        self._wait_for_redirect = wait_for_redirect
        self._oauth2_url = 'https://www.wrike.com/oauth2/token'
        self._authorization_url = 'https://www.wrike.com/oauth2/authorize'
        
        self._protocol_status.append(401)

    def authenticate(self, access_token=None):
        if access_token:
            return access_token
        elif self._client_id and self._client_secret and self._redirect_uri:
            return self._get_access_token()
        else:
            raise ValueError("Client credentials are missing.")

    def _get_access_token(self):
        # Generates a new access token if none is passed
        scopes = 'Default,wsReadWrite,amReadOnlyUser,amReadWriteUser,wsReadOnly'
        webbrowser.open(f'{self._authorization_url}?client_id={self._client_id}&response_type=code&redirect_uri={self._redirect_uri}&scope={scopes}')

        # Start the authorization server
        httpd = QuickSocketServer(self._wait_for_redirect)
        server_thread = ServerThread(httpd)
        server_thread.start()
        while httpd.authentication_code is None:
            pass
        code = httpd.authentication_code
        
        # Use the authorization code to obtain an access token
        response = requests.post(self._oauth2_url, data={
            'client_id': self._client_id,
            'client_secret': self._client_secret,
            'code': code,
            'redirect_uri': self._redirect_uri,
            'grant_type': 'authorization_code'
        })
        
        if response.status_code == 200:
            return response.json().get('access_token')
        else:
            raise ValueError("Failed to obtain access token.")

    def call_api(self, api, **args):
        access_token = self.authenticate()  # Calls authenticate to get or refresh the token
        headers = {'Authorization': f'Bearer {access_token}'}
        result, status = super().call(api, headers=headers, **args)
        return result, status
