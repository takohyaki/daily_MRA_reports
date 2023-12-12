import requests
from datetime import datetime, timedelta

def obtain_access_token(tenant_id, client_id, client_secret):
    """
    Obtains an access token from Azure AD.
    """
    url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    response = requests.post(url, headers=headers, data=data)
    response.raise_for_status()
    token_data = response.json()
    return token_data['access_token'], token_data['expires_in']

def get_access_token_info(tenant_id, client_id, client_secret):
    """
    Obtain a new access token and its expiry time.
    """
    access_token, expires_in = obtain_access_token(tenant_id, client_id, client_secret)
    expiry_time = datetime.now() + timedelta(seconds=expires_in - 300)  # Subtract 5 minutes for buffer
    return {'access_token': access_token, 'expiry_time': expiry_time}

def is_token_valid(access_token_info):
    """
    Check if the access token is valid.
    """
    if not access_token_info or 'expiry_time' not in access_token_info:
        return False

    expiry_time = access_token_info['expiry_time']
    return datetime.now() < expiry_time

# Global variable to store the access token info
access_token_info = None

def get_access_token(tenant_id, client_id, client_secret):
    """
    Returns a valid access token, obtaining a new one if necessary.
    """
    global access_token_info
    if not is_token_valid(access_token_info):
        access_token_info = get_access_token_info(tenant_id, client_id, client_secret)
    
    return access_token_info['access_token']

# Usage
tenant_id = '7950fddb-1cbe-4f8b-a278-7ed5cd483ac4'
client_id = 'df4f7567-dfd2-4f9d-92cd-7dbede616fc0'
client_secret = 'Nnx8Q~XEXNZ.GEVhY8NY3sUPGIH8uVOxoCxxvaNM'
