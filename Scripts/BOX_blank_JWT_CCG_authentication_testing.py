from box_sdk_gen import BoxClient,BoxCCGAuth, CCGConfig, JWTConfig, BoxJWTAuth


jwt_config = JWTConfig(
    client_id=['CLIENT ID GOES HERE'],
    client_secret=['CLIENT SECRET GOES HERE'],
    jwt_key_id=['JWT KEY ID GOES HERE'],
    private_key=['PRIVATE KEY GOES HERE']
    private_key_passphrase=['PRIVATE KEY PASSPHRASE GOES HERE'],
    enterprise_id=['ENTERPRISE ID GOES HERE']
) 
ccg_config = CCGConfig(
    client_id=['CLIENT ID GOES HERE'],
    client_secret=['CLIENT SECRET GOES HERE'],
    enterprise_id=['ENTERPRISE ID GOES HERE']
)
def authenticatewithJWT():
    auth = BoxJWTAuth(config=jwt_config)
    client = BoxClient(auth=auth)
    service_account = client.users.get_user_me()
    print(f"Service Account user ID is {service_account.id}; name: {service_account.name}")

def authenticatewithCCG():
    auth = BoxCCGAuth(config=ccg_config)
    client = BoxClient(auth=auth)
    service_account = client.users.get_user_me()
    print(f"Service Account user ID is {service_account.id}; name: {service_account.name}")


if __name__=='__main__':
    authenticatewithJWT()
    authenticatewithCCG()