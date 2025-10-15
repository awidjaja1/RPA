import requests
import json
from boxsdk import Client, OAuth2

BOX_CREDENTIALS = {
    'client_id':'zf6cdykvot8apgss1zg2tyxq9qt9nbae',
    'client_secret':'Nlva14zzsQRx6VeAe0HnviADU7FXI9V8',
    'grant_type':'client_credentials',
    'box_subject_type':'enterprise',
    'box_subject_id':'2384924'
    }


def auth():
    """
    Gets a requests session with the Bearer token for the Risk Box system account.
    Uses the Client Credentials Grant authentcation:

        https://developer.box.com/guides/authentication/client-credentials//

    Raises
    ------
    ValueError if bad status code is returned

    Returns
    -------
    s : requests Session with the Bearer token set for Box API use

    """
    url = 'https://ucop.app.box.com/oauth2/token'
    data = BOX_CREDENTIALS
    headers = {'content-type':'application/x-www-form-urlencoded'}
    resp = requests.post(url, data=data, headers=headers)
    if resp.status_code != 200:
        raise ValueError("Error during Box login. Error code: " + str(resp.status_code))
    token = json.loads(resp.text)['access_token']
    #s = requests.Session()
    #s.headers['authorization'] = f'Bearer {token}'
    return(token)

def getfiles():
    accessToken= auth()
    print("accessToken:", accessToken)
    headers={"Authorization":"Bearer "+accessToken,
             "content-type":"application/json",
             "as-user":"19415905383"
             }
    files_id ='1945863800792'
    url=f"https://ucop.app.box.com/files/{files_id}/content"
    print("url: ",url)
    try:
        resp2 = requests.get(url, headers=headers)
        if resp2.status_code != 200:
            raise ValueError("Error during Box login. Error code: " + str(resp2.status_code))
    except Exception as e:
        print(e)
    
    # try:
    #     oauth = OAuth2(
    #         client_id=BOX_CREDENTIALS['client_id'],
    #         client_secret=BOX_CREDENTIALS['client_secret'],
    #         access_token=accessToken
    #     )
    #     client= Client(oauth)
    #     folder_id ='333629015296'
    #     folder_items= client.folder(folder_id=folder_id).get_items()
    #     for f in folder_items:
    #         print(f)
    # except Exception as e:
    #     print(e)


getfiles()


