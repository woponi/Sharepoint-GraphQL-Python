import requests
import msal
import json
import os

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
        self.access_token = token['access_token']
        headers = {"Authorization": f"Bearer {self.access_token}"}

        self.site_url = site_url
        url = f'{GRAPH_URL}sites/' + site_url
        res = json.loads(requests.get(url, headers=headers).text)
        self.site_id = res['id']

        # Get share documents path
        url = f'{GRAPH_URL}sites/{self.site_id}/drive/'
        doc_res = json.loads(requests.get(url, headers=headers).text)
        res = json.loads(requests.get(url, headers=headers).text)

        self.documents_id = res['id']


    def list_files(self, folder_path):
        """
        Lists files within a specific folder on the SharePoint site.

        Args:
            folder_path: The server-relative path of the folder (e.g., "/sites/your-site/Shared Documents/subfolder").

        Returns:
            A list of dictionaries representing files, each containing properties like name, id, and downloadUrl.
            An empty list if there are no files or an error occurs.
        """

        url = f"{GRAPH_URL}drives/{self.documents_id}/root:/{folder_path}:/children"
        headers = {"Authorization": f"Bearer {self.access_token}"}

        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()  # Raise exception for non-200 status codes
            data = response.json()
            return data.get("value", [])  # Extract "value" array containing files
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

        url = f"{GRAPH_URL}/sites/{self.site_id}/drive/root:/{remote_path}"

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

        url = f"{GRAPH_URL}/sites/{self.site_id}/drive/root:/{remote_path}:/content"

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

    def move_file(self, remote_src_path, remote_des_path):
        """
        Move a file by its source path to the destination from the SharePoint site.

        Args:
            remote_src_path: The remote file path of the source file
            remote_des_path: The remote file path of the destination file

        Returns:
            True if move file successful, False otherwise.
        """

        new_filename = os.path.basename(remote_des_path)
        path = os.path.dirname(remote_des_path)

        # Construct the path reference
        path_reference = f"drives/{self.documents_id}/root:/{path}"

        url = f"{GRAPH_URL}/sites/{self.site_id}/drive/root:/{remote_src_path}"

        headers = {"Authorization": f"Bearer {self.access_token}"}

        # Payload for the move request
        payload = {
            "parentReference": {
                'path': path_reference
            },
            "name": new_filename
        }

        try:
            response = requests.patch(url, headers=headers, stream=True, json=payload)
            response.raise_for_status()
            data = response.json()

            return True
        except (requests.exceptions.RequestException, KeyError) as e:
            print(f"Error downloading file: {e}")
            return None

    def delete_file_by_relative_path(self, remote_path):
        """
        Delete a file by its relative path from the SharePoint site.

        Args:
            remote_path: The file path of the file to delete. (Relative path start after Documents)

        Returns:
            True if delete file successful, False otherwise.
        """

        url = f"{GRAPH_URL}/sites/{self.site_id}/drive/root:/{remote_path}"

        headers = {"Authorization": f"Bearer {self.access_token}"}

        try:
            response = requests.delete(url, headers=headers, stream=True)
            response.raise_for_status()

            return True
        except (requests.exceptions.RequestException, KeyError, OSError) as e:
            print(f"Error deleteing file: {e}")
            return False

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
