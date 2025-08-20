### SharePoint GraphQL 

### Project on PyPI
https://pypi.org/project/sharepoint-graphql/

#### Overview:
This Python utility enables users to interact with SharePoint sites via Microsoft Graph API, facilitating tasks such as listing, downloading, uploading, moving, and deleting files.
The motivation behind creating this repository stems from the absence of a SharePoint GraphQL interface in the existing Office365 Python library client (Office365-REST-Python-Client)
This README provides instructions for setting up and using the utility.

#### Prerequisites:
- Python installed on your machine.
- Microsoft Azure AD tenant with necessary permissions to access SharePoint resources. 
- Registered Azure AD application with appropriate API permissions. (Delegated permission)

#### Installation:
1. Use pip
```shell
pip install sharepoint-graphql
```


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
   - `delete_file_by_relative_path(remote_path)`: Delete a file by its relative path with enhanced error handling.
   - `get_file_metadata_by_relative_path(remote_path)`: Retrieve comprehensive metadata for a file.

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

# Delete a file with enhanced error handling
result = sp_graphql.delete_file_by_relative_path("/Shared Documents/Folder/file.txt")
if result['success']:
    print("File deleted successfully!")
else:
    print(f"Deletion failed: {result['error_details']['error_type']}")
    print(f"Error code: {result['error_code']}")
    if result['file_metadata']:
        print(f"File was last modified by: {result['file_metadata']['last_modified_by']}")

# Get file metadata
metadata = sp_graphql.get_file_metadata_by_relative_path("/Shared Documents/Folder/file.txt")
if metadata:
    print(f"File: {metadata['name']}")
    print(f"Size: {metadata['size']} bytes")
    print(f"Modified by: {metadata['lastModifiedBy']['user']['displayName']}")
```

#### Enhanced Error Handling

The `delete_file_by_relative_path()` method now returns a comprehensive result dictionary instead of a simple boolean:

```python
result = sp_graphql.delete_file_by_relative_path("/path/to/file.xlsx")

# Result structure:
{
    'success': bool,           # True if deletion succeeded, False if failed
    'error_code': int,         # HTTP status code (423, 403, 404, etc.) or None
    'error_details': dict,     # Detailed error information if failed, None if succeeded
    'file_metadata': dict      # File metadata if available, None if not available
}
```

##### Common Error Codes:
- **423 (Locked)**: File is currently locked or being edited
- **403 (Forbidden)**: Insufficient permissions to delete the file
- **404 (Not Found)**: File does not exist or cannot be found
- **409 (Conflict)**: File operation conflict (e.g., during synchronization)

##### Error Handling Example:
```python
result = sp_graphql.delete_file_by_relative_path("/Documents/locked_file.xlsx")

if result['success']:
    print("✅ File deleted successfully!")
else:
    error_code = result['error_code']
    error_details = result['error_details']
    
    if error_code == 423:
        print("🔒 File is locked")
        print(f"Likely locked by: {result['file_metadata']['last_modified_by']}")
        print("Suggestion: Contact the user to unlock the file")
    elif error_code == 403:
        print("🚫 Permission denied")
        print("Suggestion: Check your access permissions")
    elif error_code == 404:
        print("❓ File not found")
        print("Suggestion: Verify the file path")
    else:
        print(f"❌ Unexpected error: {error_details['error_type']}")
```

#### File Metadata

The `get_file_metadata_by_relative_path()` method returns comprehensive file information:

```python
metadata = sp_graphql.get_file_metadata_by_relative_path("/Documents/file.xlsx")
```

##### Return Format:
```python
{
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#sites(...)/drive/root/$entity",
    "id": "01DJ3JZXU4GYM7ZKFZXVCI5J7H6EM6W4ZW",
    "name": "document.xlsx",
    "size": 166575,
    "createdDateTime": "2025-08-19T08:23:18Z",
    "lastModifiedDateTime": "2025-08-19T08:22:47Z",
    "webUrl": "https://sharepoint.com/sites/site/_layouts/15/Doc.aspx?sourcedoc=...",
    "@microsoft.graph.downloadUrl": "https://sharepoint.com/sites/site/_layouts/15/download.aspx?UniqueId=...",
    "createdBy": {
        "user": {
            "displayName": "John Doe",
            "email": "john.doe@company.com"
        }
    },
    "lastModifiedBy": {
        "user": {
            "displayName": "Jane Smith", 
            "email": "jane.smith@company.com"
        }
    },
    "parentReference": {
        "driveId": "b!...",
        "driveType": "documentLibrary",
        "id": "01DJ3JZXU4GYM7ZKFZXVCI5J7H6EM6W4ZW",
        "name": "Shared Documents",
        "path": "/drive/root:/Documents"
    },
    "file": {
        "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "hashes": {
            "quickXorHash": "0cN4uzzESpARSqDKVB9qEwvvETA="
        }
    },
    "fileSystemInfo": {
        "createdDateTime": "2025-08-19T08:23:18Z",
        "lastModifiedDateTime": "2025-08-19T08:22:47Z"
    },
    "shared": {
        "scope": "anonymous"
    },
    "eTag": "\"{FC19369C-B9A8-44BD-8EA7-E7F119EB7336},1\"",
    "cTag": "\"c:{FC19369C-B9A8-44BD-8EA7-E7F119EB7336},0\""
}
```

##### Key Metadata Properties:

| Property | Type | Description |
|----------|------|-------------|
| `id` | string | Unique file identifier |
| `name` | string | File name with extension |
| `size` | integer | File size in bytes |
| `createdDateTime` | string | ISO 8601 timestamp when file was created |
| `lastModifiedDateTime` | string | ISO 8601 timestamp when file was last modified |
| `webUrl` | string | Direct link to view file in SharePoint |
| `@microsoft.graph.downloadUrl` | string | Direct download URL with authentication |
| `createdBy.user.displayName` | string | Name of user who created the file |
| `createdBy.user.email` | string | Email of user who created the file |
| `lastModifiedBy.user.displayName` | string | Name of user who last modified the file |
| `lastModifiedBy.user.email` | string | Email of user who last modified the file |
| `parentReference.name` | string | Name of the parent folder |
| `parentReference.path` | string | Path to the parent folder |
| `file.mimeType` | string | MIME type of the file |
| `file.hashes.quickXorHash` | string | File integrity hash |
| `shared.scope` | string | Sharing scope (anonymous, users, etc.) |
| `eTag` | string | Entity tag for caching |
| `cTag` | string | Change tag for synchronization |

##### Usage Examples:

```python
# Get basic file information
metadata = sp_graphql.get_file_metadata_by_relative_path("/Documents/report.xlsx")
if metadata:
    print(f"File: {metadata['name']}")
    print(f"Size: {metadata['size']} bytes")
    print(f"Type: {metadata['file']['mimeType']}")

# Get user information
if metadata.get('createdBy', {}).get('user'):
    creator = metadata['createdBy']['user']
    print(f"Created by: {creator['displayName']} ({creator['email']})")

if metadata.get('lastModifiedBy', {}).get('user'):
    modifier = metadata['lastModifiedBy']['user']
    print(f"Modified by: {modifier['displayName']} ({modifier['email']})")

# Get download URL
if '@microsoft.graph.downloadUrl' in metadata:
    download_url = metadata['@microsoft.graph.downloadUrl']
    print(f"Download URL: {download_url}")

# Get parent folder information
if metadata.get('parentReference'):
    parent = metadata['parentReference']
    print(f"Parent folder: {parent['name']}")
    print(f"Parent path: {parent['path']}")

# Check file integrity
if metadata.get('file', {}).get('hashes', {}).get('quickXorHash'):
    file_hash = metadata['file']['hashes']['quickXorHash']
    print(f"File hash: {file_hash}")

# Get sharing information
if metadata.get('shared'):
    shared_info = metadata['shared']
    print(f"Sharing scope: {shared_info.get('scope', 'Not shared')}")
```

#### Notes:
- Ensure your Azure AD application has the necessary permissions configured in Azure Portal.
- Handle exceptions and errors appropriately in your script for robustness.
- Refer to Microsoft Graph API documentation for additional functionalities and parameters.


#### License:
This project is licensed under the [MIT License](LICENSE).