from PyWrike.gateways.basegateway1 import APIGateway
import requests
import json
import webbrowser
import threading
import socketserver
import http.server
import re

class OAuth2CodeServer(BaseHTTPServer.BaseHTTPRequestHandler):
  def __init__(self, *args):
    self.code = None
    BaseHTTPServer.BaseHTTPRequestHandler.__init__(self, *args)

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
        self.wfile.write("Thank you, you can now close this window.")

    else:
      self.server.authentication_code = 0
      self.send_response(406)
      self.end_headers()

class QuickSocketServer(SocketServer.TCPServer):
  def __init__(self, wait_for_redirect=False):
    self.authentication_code = None
    self.redirect = None
    self.wait_for_redirect = wait_for_redirect
    SocketServer.TCPServer.__init__(self, ("", 19877), OAuth2CodeServer)

  def server_bind(self):
    self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    self.socket.bind(self.server_address)

class ServerThread(threading.Thread):
  def __init__(self, httpd, **args):
    self._httpd = httpd
    threading.Thread.__init__(self, **args)

  def run(self):
    self._httpd.handle_request()


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
