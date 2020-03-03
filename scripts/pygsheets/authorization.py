# -*- coding: utf-8 -*-.
import os
import json
import warnings
import webbrowser
import Google
from tkinter import *

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow

from pygsheets.client import Client

userInput = ""
authInput = ""

try:
    input = raw_input
except NameError:
    pass

def getEntry():
    global userInput 
    userInput = authInput.get()
    return

def _get_user_authentication_credentials(client_secret_file, scopes, credential_directory=None):
    """Returns user credentials."""
    if credential_directory is None:
        credential_directory = os.getcwd()
    elif credential_directory == 'global':
        home_dir = os.path.expanduser('~')
        credential_directory = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_directory):
            os.makedirs(credential_directory)
    else:
        pass

    credentials_path = os.path.join(credential_directory, 'access_token.json')  # TODO Change hardcoded name?

    if os.path.exists(credentials_path):
        # expect these to be valid. may expire at some point, but should be refreshed by google api client...
        return Credentials.from_authorized_user_file(credentials_path, scopes=scopes)

    flow = Flow.from_client_secrets_file(client_secret_file, scopes=scopes,
                                         redirect_uri='urn:ietf:wg:oauth:2.0:oob')

    auth_url, _ = flow.authorization_url(prompt='consent')
    webbrowser.open_new(format(auth_url))
    
    global userInput
    global authInput
    inputWindow = Tk()
    inputWindow.title("Planner Authorisation")
    inputWindow.config(bg="#EEE8AA",relief="raised",borderwidth="16")
    Label(inputWindow,text="FIAP Distinctions Planner", font="Arial 40 bold",borderwidth=8,relief="raised",bg="#EEE8AA",fg="#2F4F4F",padx="150").pack(side="top",fill="x")
    Label(inputWindow,text="Your web browser should have opened.\n\nFollow Google's instructions there to get your authorisation code.\n\nOnce you have it, paste it into the box below and press the 'Authorise' button",bg="#EEE8AA",fg="#2F4F4F",font="Ariel 14 bold",relief="groove").pack(fill="x",padx="5",pady="5")
    authInput = Entry(inputWindow,font="Ariel 14 bold",borderwidth=6)
    authInput.pack(fill="x",padx="30")
    Button(inputWindow,text="Authorise",font="Arial 20 bold",borderwidth="4",bg="#838B8B",fg="#FFFFFF",command=lambda:[getEntry(),inputWindow.destroy()]).pack(fill="x",pady="10",padx="10")
    inputWindow.mainloop()
    
    code = str(userInput)
    flow.fetch_token(code=code)

    credentials = flow.credentials

    credentials_as_dict = {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'id_token': credentials.id_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret
    }

    with open(credentials_path, 'w') as file:
        file.write(json.dumps(credentials_as_dict))

    return credentials


_SCOPES = ('https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive')

_deprecated_keyword_mapping = {
    'outh_file': 'client_secret',
    'outh_creds_store': 'credentials_directory',
    'service_file': 'service_account_file',
    'credentials': 'custom_credentials'
}


def authorize(client_secret='client_secret.json',
              service_account_file=None,
              credentials_directory='',
              scopes=_SCOPES,
              custom_credentials=None,
              **kwargs):

    """Authenticate this application with a google account.

    See general authorization documentation for details on how to attain the necessary files.

    :param client_secret:           Location of the oauth2 credentials file.
    :param service_account_file:    Location of a service account file.
    :param credentials_directory:   Location of the token file created by the OAuth2 process. Use 'global' to store in
                                    global location, which is OS dependent. Default None will store token file in
                                    current working directory. Please note that this is override your client secret.
    :param custom_credentials:      A custom or pre-made credentials object. Will ignore all other params.
    :param scopes:                  The scopes for which the authentication applies.
    :param kwargs:                  Parameters to be handed into the client constructor.
    :returns:                       :class:`Client`

    .. warning::
        The `credentials_directory` overrides `client_secrest`. So you might be accidently using a different credntial
        than intended, if you are using global `credentials_directory` in more than one script.

    """

    for key in kwargs:
        if key in ['outh_file', 'outh_creds_store', 'service_file', 'credentials']:
            warnings.warn('The argument {} is deprecated. Use {} instead.'.format(key, _deprecated_keyword_mapping[key])
                          , category=DeprecationWarning)
    client_secret = kwargs.get('outh_file', client_secret)
    service_account_file = kwargs.get('service_file', service_account_file)
    credentials_directory = kwargs.get('outh_creds_store', credentials_directory)
    custom_credentials = kwargs.get('credentials', custom_credentials)

    if custom_credentials is not None:
        credentials = custom_credentials
    elif service_account_file is not None:
        credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    else:
        credentials = _get_user_authentication_credentials(client_secret, scopes, credentials_directory)

    return Client(credentials)
