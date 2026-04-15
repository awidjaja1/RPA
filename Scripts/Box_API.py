from box_sdk_gen import BoxClient, BoxCCGAuth, CCGConfig,FolderMini
from box_sdk_gen.managers.uploads import UploadFileAttributes, UploadFileAttributesParentField, UploadFileVersionAttributes
import os

def uploadFile(credList):
    """
    This function is used to upload a new file to a Box folder and if the file exists, update the version of the file

    Args:
    credList (list):   The list must have the following data in the specific index: 
    # 0 - client ID
    # 1 - client secret
    # 2 - local folder path where the file to be uploaded
    # 3 - id of the Box folder to upload the file
    
    Returns:
    a string "Success" indicating that transaction is successful. Otherwise it returns exception message (string).
    """
    client_id = credList[0]
    client_secret = credList[1]
    folder_path = credList[2]
    folder_id = credList[3]
    ccg_config = CCGConfig(
            client_id=client_id,
            client_secret=client_secret,
            enterprise_id="2384924"
        )
    folder_id = folder_id
    file_dictionary={}

    try:
        auth = BoxCCGAuth(config=ccg_config)
        client = BoxClient(auth=auth)
        file_to_upload=folder_path
        file_name = os.path.basename(file_to_upload)

        attrs=UploadFileAttributes(name=file_name, parent=UploadFileAttributesParentField(id=folder_id))
		
        # generating a dictionary of files in specified Box folder and it's respective id
        for item in client.folders.get_folder_items(folder_id).entries:
            file_dictionary[item.name] = item.id

        # case where file exists and will be updated
        if file_name in file_dictionary:
            with open(file_to_upload, "rb") as f:
                item_id = file_dictionary[file_name]
                new_version_upload = client.uploads.upload_file_version(
                    item_id,
                    UploadFileVersionAttributes(name=file_name),
                    f,
                )
        # case where file does not exist and is new and to be uploaded
        else:
            with open(file_to_upload, "rb") as f:
                result = client.uploads.upload_file(attrs, f)
        return("Success")
    except Exception as e:
        return str(e)



def moveit(credList):
    """
    This function is used to move (or sweep) all files within a Box folder to a destination Box folder.

    Args:
    credList (list):   The list must have the following data in the specific index: 
    # 0 - client ID
    # 1 - client secret
    # 2 - folder id which the files to be moved from
    # 3 - destination folder id which the file to be moved into.
    
    Returns:
    a string "Success" indicating that transaction is successful. Otherwise it returns exception message (string).
    """
    client_id = credList[0]
    client_secret = credList[1]
    folder_id = credList[2]
    destinationFolderId = credList[3]
    ccg_config = CCGConfig(
        client_id=client_id,
        client_secret=client_secret,
        enterprise_id="2384924"
    )
    auth = BoxCCGAuth(config=ccg_config) #with CCG
    client = BoxClient(auth=auth)
    listFilesId =[]
    listFilesName=[]
    folder_id= folder_id
    destinationFolder= destinationFolderId
    #print("here are the list of filename and id:")
    for item in client.folders.get_folder_items(folder_id).entries:
        #print(f'{item.name}:  {item.id}')
        listFilesName.append(item.name)
        listFilesId.append(item.id)
        #print("============")
        #print('\n')

    for ix in listFilesId:
        try:
            filemove(ix,destinationFolder,client)
            return("Success")
        except Exception as e:
            #print("Exception: ",e)
            return(str(e))
        
def filemove(fileId, destinationFolder, client):
    try:
        file_to_move = client.files.get_file_by_id(fileId)
        updated_file = client.files.update_file_by_id(
            file_id=file_to_move.id,
            parent=FolderMini(id=destinationFolder)
        )
        #print("file moved!")
        return("Success")
    except Exception as e:
        #print(e)
        return(str(e))

def checkFiles(credList):
    """
    This function returns a stringified List object. You will need to parse the stringified list object in AA. 
    
    credList is a list of credentials in the following order: Box Client ID, Box Client Secret, Box Folder ID
    Args:
    credList (list):   The list must have the following data in the specific index: 
    # 0 - client ID
    # 1 - client secret
    # 2 - folder id to iterate the list of files
    
    Returns:
    a stringified list of file names.
    
    Notes:
    Box Folder ID - to find Box folder id, you need to login to ucop.app.box.com, then navigate to the folder of choice. The folder ID is within URL path. 
    # For example:https://ucop.app.box.com/folder/349441206902 ; where 349441206902 is the Folder ID.
    """

    listFileNames = []

    ccg_config = CCGConfig(
        client_id=credList[0],
        client_secret=credList[1],
        enterprise_id="2384924"
    )

    try:
        folder_id = credList[2]
        auth = BoxCCGAuth(config=ccg_config)
        client = BoxClient(auth=auth)

        for item in client.folders.get_folder_items(folder_id).entries:
            listFileNames.append(item.name)

    except Exception as e:
        return e

    return listFileNames

