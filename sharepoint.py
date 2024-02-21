from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import logging
import os

logger = logging.getLogger('sharepoint.sharepoint_client')

class SharepointClient:
    def __init__(self,
                 sharepoint_user: str = None,
                 sharepoint_password: str = None,
                 sharepoint_base_url: str = None
                ):
        self.sharepoint_user = sharepoint_user
        self.sharepoint_password = sharepoint_password
        self.sharepoint_base_url = sharepoint_base_url
        self.client_context = None
        
    
    def get_connection(self):
        sharepoint_user = self.sharepoint_user
        sharepoint_password = self.sharepoint_password
        sharepoint_base_url = self.sharepoint_base_url
        
        # authenticate user
        auth = AuthenticationContext(sharepoint_base_url)
        auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
        
        client_context = ClientContext(sharepoint_base_url, auth)
        web = client_context.web 
        client_context.load(web)
        client_context.execute_query()
        print('Connected to SharePoint: ',web.properties['Title'])
        logger.info('Connected to SharePoint: ',web.properties['Title'])
        return client_context
        
        
    def list_folder(self, folder_path):
        if not self.client_context:
            client_context = self.get_connection()
        
        # get folder details
        folder = client_context.web.get_folder_by_server_relative_url(folder_path)
        folder_names = []
        
        # get sub folders
        sub_folders = folder.folders
        client_context.load(sub_folders)
        client_context.execute_query()
        for sub_folder in sub_folders:
            folder_names.append(sub_folder.properties["Name"])
        logger.info(f'Folders present: {folder_names}')
        return folder_names
    
    
    def download_from_sharepoint(self, file_path, file_name):
        if not self.client_context:
            client_context = self.get_connection()
        file_response = File.open_binary(client_context, file_path)
        print(file_response)
        logger.info(f'File response present: {file_response}')
        
        # save file locally
        with open(file_name, 'wb') as output_file:  
            output_file.write(file_response.content)


    def upload_to_sharepoint(self, file_path, target_path):
        if not self.client_context:
            client_context = self.get_connection()
        
        target_folder = client_context.web.get_folder_by_server_relative_url(target_path)
        with open(file_path, 'rb') as input_file:
            file_content = input_file.read()
        
        file_name = os.path.basename(file_path)
        target_file = target_folder.upload_file(file_name, file_content).execute_query()
        logger.info(f"File has been uploaded to url: {target_file.serverRelativeUrl}")


if __name__=='__main__':
    share = SharepointClient('user', 'pwd', 'https://xxx.sharepoint.com/sites/Team/')
    
    file_list = share.list_folder(folder_path='/sites/Team/Shared%20Documents/')
    print(file_list)
    
    file_response = share.download_from_sharepoint(file_name='test.xlsx', file_path='/sites/Team/Shared%20Documents/Controlling/file.xlsx')
    print(file_response)
    
    
    upload = share.upload_to_sharepoint(file_path='./test.xlsx', target_path='/sites/Team/Shared%20Documents/')
