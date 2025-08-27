import requests
import msal
import json
import os
from urllib.parse import quote

GRAPH_URL = 'https://graph.microsoft.com/v1.0/'
class SharePointGraphql:

    def __init__(self, site_url, tenant_id, client_id, client_secret):
        """
        Acquire token via MSAL
        """
        authority_url = f'https://login.microsoftonline.com/{tenant_id}'
        app = msal.ConfidentialClientApplication(
            authority=authority_url,
            client_id=f'{client_id}',
            client_credential=f'{client_secret}'
        )
        token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        # Check if the URL starts with "https://"
        if not site_url.startswith("https://"):
            return None

        # Split the URL into parts
        parts = site_url.split("/")

        # Rebuild the URL in Graph API format
        site_url = f"{parts[2]}:/{parts[3]}/{parts[4]}:/"
        try:
            self.access_token = token['access_token']
        except KeyError:
            print("Error: Access token not found, please check your credentials")
            return None
        headers = {"Authorization": f"Bearer {self.access_token}"}

        self.site_url = site_url
        url = f'{GRAPH_URL}sites/' + site_url
        res = json.loads(requests.get(url, headers=headers).text)
        self.site_id = res['id']

        # Get share documents path
        url = f'{GRAPH_URL}sites/{self.site_id}/drive/'
        doc_res = json.loads(requests.get(url, headers=headers).text)
        res = json.loads(requests.get(url, headers=headers).text)

        if 'error' in res:
            print(f"Error: {res['error']['message']}")
            return None
        self.documents_id = res['id']


    def _encode_file_path(self, file_path):
        """
        Helper method to encode file path for Graph API requests.
        Encodes only the filename part while preserving the directory structure.
        
        Args:
            file_path: The file path to encode
            
        Returns:
            Encoded file path with filename properly URL-encoded
        """
        if not file_path:
            return file_path
            
        # Split the path into directory and filename
        directory = os.path.dirname(file_path)
        filename = os.path.basename(file_path)
        
        # Encode only the filename
        encoded_filename = quote(filename)
        
        # Reconstruct the path
        if directory:
            return f"{directory}/{encoded_filename}"
        else:
            return encoded_filename


    def list_files(self, folder_path, next_link=None, files=[]):
        """
        Lists files within a specific folder on the SharePoint site. (Max 5000 files)

        Args:
            folder_path: The server-relative path of the folder (e.g., "/sites/your-site/Shared Documents/subfolder").

        Returns:
            A list of dictionaries representing files, each containing properties like name, id, and downloadUrl.
            An empty list if there are no files or an error occurs.
        """

        url = f"{GRAPH_URL}drives/{self.documents_id}/root:/{folder_path}:/children"
        headers = {"Authorization": f"Bearer {self.access_token}"}

        if len(files) > 5000:
            raise Exception("Too many files (Try to create subfolder)")

        try:
            if next_link is not None:
                #Replace url with next link
                url = next_link
            response = requests.get(url, headers=headers)
            response.raise_for_status()  # Raise exception for non-200 status codes
            data = response.json()
            files += data.get("value", [])
            if '@odata.nextLink' in data:
                next_link = data['@odata.nextLink']
                return self.list_files(folder_path=folder_path, next_link=next_link, files=files)

            return files  # Extract "value" array containing files
        except requests.exceptions.RequestException as e:
            print(f"Error listing files: {e}")
            return []

    def download_file_by_relative_path(self, remote_path, local_path):
        """
        Downloads a file by its relative path from the SharePoint site.

        Args:
            remote_path: The file path of the file to download. (Relative path start after Documents)
            local_path: The file path of the destination your will save

        Returns:
            True if download file successful, False otherwise.
        """

        # Encode the remote path for Graph API
        encoded_remote_path = self._encode_file_path(remote_path)
        url = f"{GRAPH_URL}/sites/{self.site_id}/drive/root:/{encoded_remote_path}"

        headers = {"Authorization": f"Bearer {self.access_token}"}

        try:
            response = requests.get(url, headers=headers, stream=True)
            response.raise_for_status()
            data = response.json()

            return self.download_file(data['@microsoft.graph.downloadUrl'], local_path)
        except (requests.exceptions.RequestException, KeyError) as e:
            print(f"Error downloading file: {e}")
            return False

    def upload_file_by_relative_path(self, remote_path, local_path):
        """
        Upload a file by its relative path from the SharePoint site.

        Args:
            remote_path: The file path of the file to upload. (Relative path start after Documents)
            local_path: The file path of the local file

        Returns:
            True if upload file successful, False otherwise.
        """

        # Encode the remote path for Graph API
        encoded_remote_path = self._encode_file_path(remote_path)
        url = f"{GRAPH_URL}/sites/{self.site_id}/drive/root:/{encoded_remote_path}:/content"

        headers = {"Authorization": f"Bearer {self.access_token}"}

        try:
            with open(local_path, "rb") as f:
                response = requests.put(url, headers=headers, stream=True, data=f.read())
            response.raise_for_status()
            data = response.json()

            return True
        except (requests.exceptions.RequestException, KeyError, OSError) as e:
            print(f"Error Uploading file: {e}")
            return False

    def move_file(self, remote_src_path, remote_des_path, replace=False):
        """
        Move a file by its source path to the destination from the SharePoint site.

        Args:
            remote_src_path: The remote file path of the source file
            remote_des_path: The remote file path of the destination file
            replace: Whether to replace the destination file if it already exists (default: False)

        Returns:
            dict: A dictionary containing the result:
                - success (bool): True if move was successful, False otherwise
                - error_code (int): HTTP status code if move failed, None if successful
                - error_details (dict): Detailed error information if move failed, None if successful
                - file_metadata (dict): File metadata if available, None if not available
        """

        new_filename = os.path.basename(remote_des_path)
        path = os.path.dirname(remote_des_path)

        # Construct the path reference
        path_reference = f"drives/{self.documents_id}/root:/{path}"

        # Encode the source path for Graph API
        encoded_remote_src_path = self._encode_file_path(remote_src_path)
        url = f"{GRAPH_URL}/sites/{self.site_id}/drive/root:/{encoded_remote_src_path}"

        headers = {"Authorization": f"Bearer {self.access_token}"}

        # Payload for the move request
        payload = {
            "parentReference": {
                'path': path_reference
            },
            "name": new_filename
        }

        # Add replace parameter if specified
        if replace:
            payload["@microsoft.graph.conflictBehavior"] = "replace"

        try:
            response = requests.patch(url, headers=headers, stream=True, json=payload)
            response.raise_for_status()
            data = response.json()

            return {
                'success': True,
                'error_code': None,
                'error_details': None,
                'file_metadata': None
            }
        except requests.exceptions.HTTPError as e:
            # Get detailed error information
            error_details = {
                'error': str(e),
                'status_code': e.response.status_code,
                'source_path': remote_src_path,
                'destination_path': remote_des_path
            }
            
            # Try to get file metadata for additional context
            file_metadata = None
            try:
                metadata = self.get_file_metadata_by_relative_path(remote_src_path)
                if metadata:
                    file_metadata = {
                        'name': metadata.get('name'),
                        'size': metadata.get('size'),
                        'created_by': metadata.get('createdBy', {}).get('user', {}).get('displayName'),
                        'last_modified_by': metadata.get('lastModifiedBy', {}).get('user', {}).get('displayName'),
                        'created_date': metadata.get('createdDateTime'),
                        'modified_date': metadata.get('lastModifiedDateTime'),
                        'web_url': metadata.get('webUrl'),
                        'file_type': metadata.get('file', {}).get('mimeType'),
                        'parent_folder': metadata.get('parentReference', {}).get('name')
                    }
            except Exception as metadata_error:
                error_details['metadata_error'] = str(metadata_error)
            
            # Provide specific error messages based on status code
            if e.response.status_code == 409:
                error_details['error_type'] = 'Conflict'
                error_details['message'] = 'A file with the same name already exists at the destination. This could be due to:'
                error_details['possible_causes'] = [
                    'Destination file already exists and replace=False',
                    'File is being synchronized',
                    'File has pending changes',
                    'Destination folder has conflicting permissions'
                ]
                error_details['suggestion'] = 'Set replace=True to overwrite the existing file'
            elif e.response.status_code == 423:
                error_details['error_type'] = 'File Locked'
                error_details['message'] = 'File is currently locked and cannot be moved. This could be due to:'
                error_details['possible_causes'] = [
                    'File is being edited by another user',
                    'File is checked out',
                    'File has active sharing permissions',
                    'File is in a protected library or folder'
                ]
            elif e.response.status_code == 403:
                error_details['error_type'] = 'Permission Denied'
                error_details['message'] = 'You do not have permission to move this file. This could be due to:'
                error_details['possible_causes'] = [
                    'Insufficient permissions on the source file',
                    'Insufficient permissions on the destination folder',
                    'File is in a protected folder',
                    'Your account lacks move permissions'
                ]
            elif e.response.status_code == 404:
                error_details['error_type'] = 'File Not Found'
                error_details['message'] = 'The specified source file does not exist or cannot be found.'
            elif e.response.status_code == 400:
                error_details['error_type'] = 'Bad Request'
                error_details['message'] = 'Invalid request parameters. This could be due to:'
                error_details['possible_causes'] = [
                    'Invalid file path format',
                    'Invalid destination path',
                    'Source and destination are the same',
                    'Invalid file name characters'
                ]
            else:
                error_details['error_type'] = 'Unknown Error'
                error_details['message'] = f'Unexpected error occurred (Status: {e.response.status_code})'
            
            return {
                'success': False,
                'error_code': e.response.status_code,
                'error_details': error_details,
                'file_metadata': file_metadata
            }
        except (requests.exceptions.RequestException, KeyError, OSError) as e:
            return {
                'success': False,
                'error_code': None,
                'error_details': {
                    'error': str(e),
                    'error_type': 'Request Exception',
                    'message': 'An unexpected error occurred during the move operation.',
                    'source_path': remote_src_path,
                    'destination_path': remote_des_path
                },
                'file_metadata': None
            }

    def delete_file_by_relative_path(self, remote_path):
        """
        Delete a file by its relative path from the SharePoint site.

        Args:
            remote_path: The file path of the file to delete. (Relative path start after Documents)

        Returns:
            dict: A dictionary containing the result:
                - success (bool): True if deletion was successful, False otherwise
                - error_code (int): HTTP status code if deletion failed, None if successful
                - error_details (dict): Detailed error information if deletion failed, None if successful
                - file_metadata (dict): File metadata if available, None if not available
        """

        # Encode the remote path for Graph API
        encoded_remote_path = self._encode_file_path(remote_path)
        url = f"{GRAPH_URL}/sites/{self.site_id}/drive/root:/{encoded_remote_path}"

        headers = {"Authorization": f"Bearer {self.access_token}"}

        try:
            response = requests.delete(url, headers=headers, stream=True)
            response.raise_for_status()

            return {
                'success': True,
                'error_code': None,
                'error_details': None,
                'file_metadata': None
            }
        except requests.exceptions.HTTPError as e:
            # Get detailed error information
            error_details = {
                'error': str(e),
                'status_code': e.response.status_code,
                'file_path': remote_path
            }
            
            # Try to get file metadata for additional context
            file_metadata = None
            try:
                metadata = self.get_file_metadata_by_relative_path(remote_path)
                if metadata:
                    file_metadata = {
                        'name': metadata.get('name'),
                        'size': metadata.get('size'),
                        'created_by': metadata.get('createdBy', {}).get('user', {}).get('displayName'),
                        'last_modified_by': metadata.get('lastModifiedBy', {}).get('user', {}).get('displayName'),
                        'created_date': metadata.get('createdDateTime'),
                        'modified_date': metadata.get('lastModifiedDateTime'),
                        'web_url': metadata.get('webUrl'),
                        'file_type': metadata.get('file', {}).get('mimeType'),
                        'parent_folder': metadata.get('parentReference', {}).get('name')
                    }
            except Exception as metadata_error:
                error_details['metadata_error'] = str(metadata_error)
            
            # Provide specific error messages based on status code
            if e.response.status_code == 423:
                error_details['error_type'] = 'File Locked'
                error_details['message'] = 'File is currently locked and cannot be deleted. This could be due to:'
                error_details['possible_causes'] = [
                    'File is being edited by another user',
                    'File is checked out',
                    'File has active sharing permissions',
                    'File is in a protected library or folder'
                ]
            elif e.response.status_code == 403:
                error_details['error_type'] = 'Permission Denied'
                error_details['message'] = 'You do not have permission to delete this file. This could be due to:'
                error_details['possible_causes'] = [
                    'Insufficient permissions on the file',
                    'File is in a protected folder',
                    'File has special permissions',
                    'Your account lacks delete permissions'
                ]
            elif e.response.status_code == 404:
                error_details['error_type'] = 'File Not Found'
                error_details['message'] = 'The specified file does not exist or cannot be found.'
            elif e.response.status_code == 409:
                error_details['error_type'] = 'Conflict'
                error_details['message'] = 'There is a conflict preventing file deletion. This could be due to:'
                error_details['possible_causes'] = [
                    'File is being synchronized',
                    'File has pending changes',
                    'File is in a state that prevents deletion'
                ]
            else:
                error_details['error_type'] = 'Unknown Error'
                error_details['message'] = f'Unexpected error occurred (Status: {e.response.status_code})'
            
            return {
                'success': False,
                'error_code': e.response.status_code,
                'error_details': error_details,
                'file_metadata': file_metadata
            }
        except (requests.exceptions.RequestException, KeyError, OSError) as e:
            return {
                'success': False,
                'error_code': None,
                'error_details': {
                    'error': str(e),
                    'error_type': 'Request Exception',
                    'message': 'An unexpected error occurred during the delete operation.',
                    'file_path': remote_path
                },
                'file_metadata': None
            }

    def get_file_metadata_by_relative_path(self, remote_path):
        """
        Retrieve metadata for a file by its relative path from the SharePoint site.

        Args:
            remote_path: The file path of the file to get metadata for. (Relative path start after Documents)

        Returns:
            dict: A dictionary containing the file metadata including properties like:
                - id: File ID
                - name: File name
                - size: File size in bytes
                - createdDateTime: Creation timestamp
                - lastModifiedDateTime: Last modified timestamp
                - webUrl: Web URL to access the file
                - downloadUrl: Direct download URL
                - @microsoft.graph.downloadUrl: Microsoft Graph download URL
                - file: File-specific properties (for files)
                - folder: Folder-specific properties (for folders)
                - parentReference: Parent folder information
                - createdBy: Information about who created the file
                - lastModifiedBy: Information about who last modified the file
            None if the file doesn't exist or an error occurs.
        """

        # Encode the remote path for Graph API
        encoded_remote_path = self._encode_file_path(remote_path)
        url = f"{GRAPH_URL}/sites/{self.site_id}/drive/root:/{encoded_remote_path}"

        headers = {"Authorization": f"Bearer {self.access_token}"}

        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            data = response.json()

            return data
        except requests.exceptions.RequestException as e:
            print(f"Error retrieving file metadata: {e}")
            return None



    def download_file(self, url, output_path):
        """
        Downloads a file from a URL and saves it to the specified path.

        Args:
            url (str): The absolute URL of the file to download.
            output_path (str): The absolute path where the file will be saved.

        Returns:
            file: The file object of the downloaded file,
                or None if there was an error.

        Raises:
            OSError: If there's an issue creating the output directory or file.
            requests.exceptions.RequestException: If there's an error downloading the file.
        """

        # Get absolute path based on current working directory (for relative paths)
        if not os.path.isabs(output_path):
            output_path = os.path.join(os.getcwd(), output_path)

        # Check if output directory exists, create it if necessary
        output_dir = os.path.dirname(output_path)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Get the filename from the URL (consider using a library for robust extraction)
        filename = os.path.basename(url)

        # Download the file using requests
        try:
            response = requests.get(url, stream=True)
            response.raise_for_status()  # Raise an exception for non-2xx status codes

            # Open the output file in binary write mode
            with open(output_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:  # filter out keep-alive new chunks
                        f.write(chunk)

            return True  # Return the opened file object

        except (OSError, requests.exceptions.RequestException) as e:
            print(f"Error downloading file: {e}")
            return False
