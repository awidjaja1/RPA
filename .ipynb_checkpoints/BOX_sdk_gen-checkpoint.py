from box_sdk_gen import BoxClient, BoxDeveloperTokenAuth,BoxCCGAuth, CCGConfig, JWTConfig, BoxJWTAuth

# credentials in jwt_config are from TestAppAriJWT Box app
"""
jwt_config = JWTConfig(
    client_id="x2h6cgq83zn73tz0e5sc4u2qp1qyn7r3",
    client_secret="1CdKUjGjsCki9Vcv0xG4spEGCcf4lxDP",
    jwt_key_id="47hnuyos",
    private_key= "-----BEGIN ENCRYPTED PRIVATE KEY-----\nMIIFNTBfBgkqhkiG9w0BBQ0wUjAxBgkqhkiG9w0BBQwwJAQQ7WPa/3CnZMQ7dbzP\nDfZ37wICCAAwDAYIKoZIhvcNAgkFADAdBglghkgBZQMEASoEEPZwjeYWDuGyOYZ7\nAVqNwHMEggTQcdF6BPqObfYv0jxz2wT4Tvqwa9He3EngNlr8nMZmNvOJ/JKmp+u2\n8S0VA9/BGR1BPI8g1qdlN3eWmhrEw3+HrIjaCI9XGh05w4rk4JkAcD0WGUVISP3q\nRl+2kN7KY/2NnQlRZC0StZYHDeJFOw8PfiDQ/EN0lLBpByPEJ++INkxudBqBPdV6\n1wc9BLSZzdn0pvqzpl8Bo0gIpDVFdI1++X2C1pNM6je8ZKmQdsew9reCXEZSD780\n2stkb1B5Bdy672bDxkGxDmCnbYGgURHp9gVd89pTOuNUVtHSxJvu5W0dDFF2FDcF\nVFeZGL72vfnd8UQR0P0RN1X6T2md22lt6U40vKUx1jZUj73VkQGIAtLa1c9EZ2Oi\n489NIl7ifBgBChP81C4xKYlyPWDAQJ3CFmD4qqxR5CVaw7GpRrc21AQf/mHoXzP2\nzN7kSX7TplcItbzPn9jdTJZ016EG9e6AVtbRqEw5rHWiMYfJN5exuxiTcKff5oSi\noRgDEW8A1GIijaK/C3AI/vTKfjanci2nItZpIujkB1+gZ93CPGTLBkJLJfpHDh9G\nCIOMgQQ+1z3GWcC1t2uoELmtMIcYLiCeSiVsuSddn5B/V3T6wtypsTqlYN2kHChL\n3JEgjeifx4oknytDO+EHwTNw/IF5LcylRpKWzlqIU7bOaFQw2rgeCD7kWLz1bISp\nsWeYxnvkTxncZZlgWJPLhgYz4Kl8Y/cfSQlNT4zzU+sCBSnoNuu1QAcrsDo4luEj\n7tMqYZwQywyO0WDZ3AgRnAyhlcd7j5zEmHypPlZQvOhVTLTEgAZpBBuwLBzkP8a7\noh7M1gtoK2xTLJlhPanYfCIWU+erPJDtYu8wZU03flCMRz11uO44ODW8ZuOBV5jj\nCwkk5W2zU2xKjl9oZmo8aodaEsilALHIc9BCeSpbJ00DsMqn349PNrzZ/9jI3HGw\nZtuUPHyW92Zt+csCo+zQOvi5OaFSapLRvMolKXiC8hShoSGsHhxHfXHGU1lfbggH\nNu+p5aZis+q9ZhRx5RaKKFr1RSZzwlCfomm6InsWDenux0g0MqYPo7PQ/heLkL+S\n4Q10/u9AWveA5Ju1ppMveUYX28GfQ7socQFGJ01geYJPrR+iL74byoGtfKlsDakG\nxFa0gFDnS1J2f1xeBjtkq7HTpSXO3APmDQwTlZCZuulC+pI6jz+eMpuBBBq07612\n1e6xpzT9VYEew6k1+iOIUvzgksCOrMGW8dG3D07gZhHT4HlO0cnCC2yidUyQcQaC\nwxarVXQu9xlOwEA2j+NB5vPMUeAN8PcchenBVFW5WusRzLQ+ONK52Gf3F6J/Bf96\njdrdtY+UUGIrdYKOPaUCdfWb52CHAo2Xwk7J8Ew2WJfhaSMFswcDbweE8yUtMX7O\nc5VFKf3YUUnhachCUyjgo7jIrvs1VhgNT306hmbqE43kP/OhuOoXkujIquvs3m8d\nbOL02X6ZACxahOXYZuoUNfQWkXeCyWY4KFvSy3Yl3N71hDP+CKcFhIcrv6K48ch4\ntSkdcZZS/zCBZ5onKZ0sBTYyyvnZLUMQfNbabY6JqYNsUmdlEjXE4NPLNs3WuvPl\nWewc14a1p1fEpr33vQFF0iiBM8eDtdfvy0I7idXf/6AJ9cpoEv7kn54=\n-----END ENCRYPTED PRIVATE KEY-----\n",
    private_key_passphrase="45f2a6955d231279f38436433b4deb14",
    enterprise_id="2384924"
    #user_id="19415905383"
) 
"""
#jwt_config = JWTConfig.from_config_file(config_file_path='./2384924__config.json')
# credentials in BOX_CREDENTIALS or ccg_config are derived from TestAppAri Box app
BOX_CREDENTIALS = {
    'client_id':'zf6cdykvot8apgss1zg2tyxq9qt9nbae',
    'client_secret':'Nlva14zzsQRx6VeAe0HnviADU7FXI9V8',
    'grant_type':'client_credentials',
    'box_subject_type':'enterprise',
    'box_subject_id':'2384924',
    'developer_token':'7r9nqhX1bbzMLV4yeY9FPTaiYcPgPsUl'
    }
ccg_config = CCGConfig(
    client_id="zf6cdykvot8apgss1zg2tyxq9qt9nbae",
    client_secret="Nlva14zzsQRx6VeAe0HnviADU7FXI9V8",
    enterprise_id="2384924"
)
files_id ='1945863800792'
folder_id ='333629015296'
lisa_folder_id='336882472003'
Ari_id="19415905383"
def main():
    #token=BOX_CREDENTIALS['developer_token']
    #auth = BoxDeveloperTokenAuth(token=token) #with developer token
    auth = BoxCCGAuth(config=ccg_config) #with CCG
    #auth = BoxJWTAuth(config=jwt_config) #with JWT
    #user_auth = auth.with_user_subject(Ari_id)
    client = BoxClient(auth=auth)
    service_account = client.users.get_user_me()
    print(f"Service Account user ID is {service_account.id}; name: {service_account.name}")
    #print(service_account)
    listFilesId =[]
    
    for item in client.folders.get_folder_items(folder_id).entries:
        print(f'{item.name}:  {item.id}')
        #print(item)
        listFilesId.append(item.id)
        print('\n')

    try:
        print("Now downloading....")
        for file in listFilesId:
            print('file id: ',file)
            client.downloads.download_file(file)
    except Exception:
        print(Exception)
    print('finish downloading')
if __name__ == '__main__':
    main()