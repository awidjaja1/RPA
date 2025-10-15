from box_sdk_gen import BoxClient, BoxDeveloperTokenAuth,BoxCCGAuth, CCGConfig, JWTConfig, BoxJWTAuth


ccg_config = CCGConfig(
    client_id="zf6cdykvot8apgss1zg2tyxq9qt9nbae",
    client_secret="Nlva14zzsQRx6VeAe0HnviADU7FXI9V8",
    enterprise_id="2384924"
)

folder_id = "333627267841"

def checkFolder():
    auth = BoxCCGAuth(config=ccg_config) #with CCG
    client = BoxClient(auth=auth)
    listFilesId =[]
    for item in client.folders.get_folder_items(folder_id).entries:
        print(f'{item.name}:  {item.id}')
        #print(item)
        listFilesId.append(f"name:{item.name}; id:{item.id}")
        print('\n')
    if len(listFilesId)>0:
      return listFilesId
    else:
      return None
    
if __name__ == '__main__':
    A = checkFolder()
    print(A)