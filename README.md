### SharePoint GraphQL 

#### Overview:
This Python utility enables users to interact with SharePoint sites via Microsoft Graph API, facilitating tasks such as listing, downloading, uploading, moving, and deleting files.
The motivation behind creating this repository stems from the absence of a SharePoint GraphQL interface in the existing Office365 Python library client (Office365-REST-Python-Client)
This README provides instructions for setting up and using the utility.

#### Prerequisites:
- Python installed on your machine.
- Microsoft Azure AD tenant with necessary permissions to access SharePoint resources. 
- Registered Azure AD application with appropriate API permissions. (Delegated permission)

#### Installation:
1. Clone or download the repository containing the utility code.
2. Navigate to the project directory in your terminal or command prompt.

#### Configuration:
1. Update the following variables with your Azure AD and SharePoint details:
   - `TENANT_ID`: Your Azure AD tenant ID.
   - `CLIENT_ID`: Client ID of your Azure AD application.
   - `CLIENT_SECRET`: Client secret of your Azure AD application.
   - `SITE_URL`: URL of the SharePoint site you want to interact with.

#### Usage:
1. Import the `SharePointGraphql` class from `sharepoint_graphql.py` into your Python script.
2. Create an instance of the `SharePointGraphql` class by passing the required parameters: `site_url`, `tenant_id`, `client_id`, and `client_secret`.
3. Use the instance methods to perform various tasks:
   - `list_files(folder_path)`: List files within a specific folder.
   - `download_file_by_relative_path(remote_path, local_path)`: Download a file by its relative path.
   - `upload_file_by_relative_path(remote_path, local_path)`: Upload a file by its relative path.
   - `move_file(remote_src_path, remote_des_path)`: Move a file from source to destination.
   - `delete_file_by_relative_path(remote_path)`: Delete a file by its relative path.

#### Example:
```python
from sharepoint_graphql import SharePointGraphql

# Initialize SharePointGraphql instance
sp_graphql = SharePointGraphql(site_url, tenant_id, client_id, client_secret)

# List files in a folder
files = sp_graphql.list_files("/Shared Documents/Subfolder")

# Download a file
sp_graphql.download_file_by_relative_path("/Shared Documents/Folder/file.txt", "local_path/file.txt")

# Upload a file
sp_graphql.upload_file_by_relative_path("/Shared Documents/Folder/file.txt", "local_path/file.txt")

# Move a file
sp_graphql.move_file("/Shared Documents/Folder/file.txt", "/Shared Documents/NewFolder/file.txt")

# Delete a file
sp_graphql.delete_file_by_relative_path("/Shared Documents/Folder/file.txt")
```

#### Notes:
- Ensure your Azure AD application has the necessary permissions configured in Azure Portal.
- Handle exceptions and errors appropriately in your script for robustness.
- Refer to Microsoft Graph API documentation for additional functionalities and parameters.


#### License:
This project is licensed under the [MIT License](LICENSE).