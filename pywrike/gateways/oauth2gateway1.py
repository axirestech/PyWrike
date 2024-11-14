from PyWrike.gateways.basegateway1 import APIGateway

import os
import json
import webbrowser
import socketserver
import socket
import http.server 
import re
import threading
import pandas as pd

class OAuth2CodeServer(http.server.BaseHTTPRequestHandler):
    def __init__(self, *args):
        self.code = None
        http.server.BaseHTTPRequestHandler.__init__(self, *args)

    def do_GET(self):
        match = re.search('code=([\w|\-]+)', self.path)
        if match is not None:
            self.server.authentication_code = match.group(1)
            while self.server.wait_for_redirect and self.server.redirect is None: pass
            if self.server.wait_for_redirect:
                self.send_response(301)
                self.send_header('Location', self.server.redirect)
                self.end_headers()
            else:
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b"Thank you, you can now close this window.")
        else:
            self.server.authentication_code = 0
            self.send_response(406)
            self.end_headers()

class QuickSocketServer(socketserver.TCPServer):
    def __init__(self, wait_for_redirect=False):
        self.authentication_code = None
        self.redirect = None
        self.wait_for_redirect = wait_for_redirect
        socketserver.TCPServer.__init__(self, ("", 19877), OAuth2CodeServer)

    def server_bind(self):
        self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.socket.bind(self.server_address)

class ServerThread(threading.Thread):
    def __init__(self, httpd, **args):
        self._httpd = httpd
        threading.Thread.__init__(self, **args)

    def run(self):
        self._httpd.handle_request()

class _OAuth2Gateway1(APIGateway):
    def __init__(self, oauth2_url):
        APIGateway.__init__(self)
        self._host_url = oauth2_url
        self._common_params = {}
        self._common_headers = {}
        self._api = {
            'refresh_token': {
                'path': '',
                'method': 'POST',
                'params': {
                    'grant_type': 'refresh_token'
                }
            },
            'get_token': {
                'path': '',
                'method': 'POST',
                'params': {
                    'grant_type': 'authorization_code'
                },
                'valid_status': [200]
            }
        }

class OAuth2Gateway1(APIGateway):
    def __init__(self, data_filepath=None, auth_info=None, tokens_updater=None, excel_filepath=None, wait_for_redirect=False):
        APIGateway.__init__(self)
        self._common_params = {}
        self._common_headers = {}
        self._oauth2_gateway = None
        self._protocol_status.append(401)
        self._httpd = None
        self._serverthread = None
        self._data_filepath = data_filepath
        self._auth_info = None
        self._tokens_updater = tokens_updater
        self._excel_filepath = excel_filepath  # New field to store Excel filepath
        self._wait_for_redirect = wait_for_redirect  # Initialize wait_for_redirect

        if auth_info is not None:
            self._set_auth_info(auth_info)

        if self._excel_filepath:
            self._load_credentials_from_excel()

    # New method to load credentials from Excel
    def _load_credentials_from_excel(self):
        try:
            # Assuming 'Config' sheet has 'Client ID', 'Client Secret', and 'Redirect URI' in the first row
            config_df = pd.read_excel(self._excel_filepath, sheet_name='Config', header=1)
            self._oauth2_client_id = config_df.at[0, 'Client ID']
            self._oauth2_client_secret = config_df.at[0, 'Client Secret']
            self._oauth2_redirect_url = config_df.at[0, 'Redirect URI']
        except Exception as e:
            print(f"Error loading credentials from Excel: {e}")
            raise

    def call(self, api, **args):
        self._authenticate_client()
        result, status = super(OAuth2Gateway1, self).call(api, **args)
        if status == 401 and result['error'] == 'not_authorized':
            self._refresh_client_authentication()
            result, status = super(OAuth2Gateway1, self).call(api, **args)
        return result, status

    def update_common_headers(self, data):
        self._common_headers = {
            'Authorization': 'bearer {0}'.format(data['access_token'])
        }

    def redirect(self, redirect="http://www.google.com"):
        if redirect is not None:
            if self._httpd is not None:
                self._httpd.redirect = redirect
                self._httpd = None

            if self._serverthread is not None:
                self._serverthread.join()
                self._serverthread = None

    def get_auth_info(self):
        data = self._auth_info
        if (data is None) and (self._data_filepath is not None) and os.path.isfile(self._data_filepath):
            with open(self._data_filepath, 'r') as data_file:
                data = json.load(data_file)
        return data

    def _get_oauth2_gateway(self):
        if self._oauth2_gateway is None:
            self._oauth2_url = 'https://www.wrike.com/oauth2/token'
            self._oauth2_gateway = _OAuth2Gateway1(self._oauth2_url)
        return self._oauth2_gateway

    def _authenticate_client(self):
        auth_info = self.get_auth_info()
        if auth_info is None:
            auth_info = self._create_auth_info()
        self._set_auth_info(auth_info)

    def _set_auth_info(self, new_auth_info):
        if self._auth_info != new_auth_info:
            self._dump_auth_info_to_file(new_auth_info)
            self._auth_info = new_auth_info
            self.update_common_headers(new_auth_info)
            if self._tokens_updater is not None:
                self._tokens_updater.new_tokens(refresh_token=new_auth_info.get('refresh_token'), access_token=new_auth_info.get('access_token'))

    def _create_auth_info(self):
        # Set the OAuth2 authorization URL
        self._oauth2_authorization_url = 'https://www.wrike.com/oauth2/authorize'
        
        # Use self._oauth2_redirect_url instead of redirect_uri
        if not self._oauth2_client_id or not self._oauth2_redirect_url:
            raise ValueError("OAuth2 Client ID or Redirect URI is not set.")
        
        scopes = 'Default,wsReadWrite,amReadOnlyUser,amReadWriteUser,wsReadOnly,amReadOnlyWorkflow,amReadWriteWorkflow,wsReadOnly'

        # Open the authorization URL in the web browser
        webbrowser.open(self._oauth2_authorization_url + 
                        f'?client_id={self._oauth2_client_id}&response_type=code&redirect_uri={self._oauth2_redirect_url}&scope={scopes}')

        self._httpd = QuickSocketServer(self._wait_for_redirect)
        self._serverthread = ServerThread(self._httpd)
        self._serverthread.start()
        while self._httpd.authentication_code is None: pass
        authentication_code = self._httpd.authentication_code

        # Use the authentication code to get the access token
        auth_info = self._get_oauth2_gateway().call('get_token', params={
            'client_id': self._oauth2_client_id,
            'client_secret': self._oauth2_client_secret,
            'code': authentication_code,
            'redirect_uri': self._oauth2_redirect_url  # Use self._oauth2_redirect_url here as well
        })[0]

        # Extract the access token only
        access_token = auth_info.get('access_token')
        return access_token

    def _refresh_client_authentication(self):
        auth_info = self._get_oauth2_gateway().call('refresh_token', params={
            'client_id': self._oauth2_client_id,
            'client_secret': self._oauth2_client_secret,
            'refresh_token': self.get_auth_info()['refresh_token']
        })[0]
        self._set_auth_info(auth_info)

    def _dump_auth_info_to_file(self, auth_info):
        if self._data_filepath is not None:
            with open(self._data_filepath, 'w') as outfile:
                json.dump(auth_info, outfile)
