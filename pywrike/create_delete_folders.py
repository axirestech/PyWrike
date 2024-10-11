import requests
import pandas as pd
from PyWrike.gateways import OAuth2Gateway1

# Function to validate the access token
def validate_token(access_token):
    endpoint = 'https://www.wrike.com/api/v4/contacts'
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    
    response = requests.get(endpoint, headers=headers)
    if response.status_code == 200:
        print("Access token is valid.")
        return True
    else:
        print(f"Access token is invalid. Status code: {response.status_code}")
        return False

# Function to authenticate using OAuth2 if the token is invalid
def authenticate_with_oauth2(client_id, client_secret, redirect_url):
    wrike = OAuth2Gateway1(client_id=client_id, client_secret=client_secret)
    
    # Start the OAuth2 authentication process
    auth_info = {
        'redirect_uri': redirect_url
    }
    
    # Perform OAuth2 authentication and retrieve the access token
    access_token = wrike.authenticate(auth_info=auth_info)
    
    print(f"New access token obtained: {access_token}")
    return access_token

# Function to get the ID of a folder by its name
def get_folder_id_by_name(folder_name, access_token):
    endpoint = 'https://www.wrike.com/api/v4/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve folders. Status code: {response.status_code}")
        print(response.text)
        return None

    folders = response.json().get('data', [])
    for folder in folders:
        if folder['title'] == folder_name:
            return folder['id']

    print(f"Folder with name '{folder_name}' not found.")
    return None

# Function to get the ID of a subfolder by its name and parent folder ID
def get_subfolder_id_by_name(parent_folder_id, subfolder_name, access_token):
    endpoint = f'https://www.wrike.com/api/v4/folders/{parent_folder_id}/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve subfolders. Status code: {response.status_code}")
        print(response.text)
        return None

    subfolders = response.json().get('data', [])
    for subfolder in subfolders:
        if subfolder['title'] == subfolder_name:
            return subfolder['id']

    return None  # Return None if subfolder not found

# Function to create a new project in Wrike
def create_wrike_project(access_token, parent_folder_id, project_title):
    endpoint = f'https://www.wrike.com/api/v4/folders/{parent_folder_id}/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    data = {
        'title': project_title,
    }

    response = requests.post(endpoint, headers=headers, json=data)

    if response.status_code == 200:
        project_id = response.json()['data'][0]['id']
        print(f"Project '{project_title}' created successfully with ID: {project_id}!")
        return project_id
    else:
        print(f"Failed to create project '{project_title}'. Status code: {response.status_code}")
        print(response.text)
        return None

# Function to create a new folder in a project in Wrike
def create_wrike_folder(access_token, parent_folder_id, folder_title):
    endpoint = f'https://www.wrike.com/api/v4/folders/{parent_folder_id}/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    data = {
        'title': folder_title,
    }

    response = requests.post(endpoint, headers=headers, json=data)

    if response.status_code == 200:
        print(f"Folder '{folder_title}' created successfully!")
    else:
        print(f"Failed to create folder '{folder_title}'. Status code: {response.status_code}")
        print(response.text)

# Function to delete a folder in Wrike
def delete_wrike_folder(access_token, parent_folder_id, folder_title):
    folder_id = get_subfolder_id_by_name(parent_folder_id, folder_title, access_token)
    if not folder_id:
        print(f"Folder '{folder_title}' not found in project.")
        return

    endpoint = f'https://www.wrike.com/api/v4/folders/{folder_id}'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.delete(endpoint, headers=headers)

    if response.status_code == 200:
        print(f"Folder '{folder_title}' deleted successfully!")
    else:
        print(f"Failed to delete folder '{folder_title}'. Status code: {response.status_code}")
        print(response.text)

# Function to delete a project in Wrike
def delete_wrike_project(access_token, parent_folder_id, project_title):
    project_id = get_subfolder_id_by_name(parent_folder_id, project_title, access_token)
    if not project_id:
        print(f"Project '{project_title}' not found.")
        return

    endpoint = f'https://www.wrike.com/api/v4/folders/{project_id}'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.delete(endpoint, headers=headers)

    if response.status_code == 200:
        print(f"Project '{project_title}' deleted successfully!")
    else:
        print(f"Failed to delete project '{project_title}'. Status code: {response.status_code}")
        print(response.text)

# Main function to handle project and folder creation and deletion
def create_delete_folders_main():
    excel_file = input("Enter the path to the Excel file: ")
    try:
        config_df = pd.read_excel(excel_file, sheet_name="Config", header=1)
        project_df = pd.read_excel(excel_file, sheet_name="Projects")
    except Exception as e:
        print(f"Failed to read Excel file: {e}")
        return 

    access_token = config_df.at[0, "Token"]
    folder_path = config_df.at[0, "Project Folder Path"]
        # Validate the token
    if not validate_token(access_token):
        # If the token is invalid, authenticate using OAuth2 and update the access_token
        wrike = OAuth2Gateway1(excel_filepath=excel_file)
        access_token = wrike._create_auth_info()  # Perform OAuth2 authentication
        print(f"New access token obtained: {access_token}")
        
    
    parent_folder_id = get_folder_id_by_name(folder_path, access_token)
    if not parent_folder_id:
        print("Parent folder ID not found.")
        return
    print(f"Parent folder ID: {parent_folder_id}")

    # Group rows by project name
    create_projects = project_df[["Create Project Title", "Create Folders"]]
    delete_projects = project_df[["Delete Project Title", "Delete Folders"]]

    # Create projects and folders
    for project_name, group in create_projects.groupby("Create Project Title"):
        print(f"Processing project: {project_name}")
        
        project_id = get_subfolder_id_by_name(parent_folder_id, project_name.strip(), access_token)
        
        if not project_id:
            print(f"Project '{project_name.strip()}' not found. Creating it.")
            project_id = create_wrike_project(access_token, parent_folder_id, project_name.strip())
        else:
            print(f"Project '{project_name.strip()}' found with ID: {project_id}")

        if project_id:
            folders = group["Create Folders"].dropna().tolist()
            if not folders:
                print(f"No folders to create in project '{project_name.strip()}'")
            for folder_name in folders:
                if folder_name.strip():
                    print(f"Creating folder '{folder_name.strip()}' in project '{project_name.strip()}'")
                    create_wrike_folder(access_token, project_id, folder_name.strip())

    # Delete projects and folders
    for project_name, group in delete_projects.groupby("Delete Project Title"):
        print(f"Processing project: {project_name}")

        project_id = get_subfolder_id_by_name(parent_folder_id, project_name.strip(), access_token)
        
        if project_id:
            folders = group["Delete Folders"].dropna().tolist()
            if not folders:
                print(f"No folders to delete in project '{project_name.strip()}'")
            for folder_name in folders:
                if folder_name.strip():
                    print(f"Deleting folder '{folder_name.strip()}' from project '{project_name.strip()}'")
                    delete_wrike_folder(access_token, project_id, folder_name.strip())
            
            # Delete the project itself if folders are not specified or all are deleted
            print(f"Deleting project '{project_name.strip()}' itself")
            delete_wrike_project(access_token, parent_folder_id, project_name.strip())
        else:
            print(f"Project '{project_name.strip()}' not found, skipping deletion.")

if __name__ == '__create_delete_folders_main__':
    create_delete_folders_main()