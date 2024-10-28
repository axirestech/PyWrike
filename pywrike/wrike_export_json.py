import requests
import json
import pandas as pd
import os
from PyWrike.gateways import OAuth2Gateway1

BASE_URL = 'https://www.wrike.com/api/v4'

# Function to validate the access token
def validate_token(wrike_api_token):
    endpoint = f'{BASE_URL}/contacts'
    headers = {
        'Authorization': f'Bearer {wrike_api_token}'
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
    auth_info = {'redirect_uri': redirect_url}
    wrike_api_token = wrike.authenticate(auth_info=auth_info)
    print(f"New access token obtained: {wrike_api_token}")
    return wrike_api_token

# Function to get all spaces
def get_all_spaces(wrike_api_token):
    url = f'{BASE_URL}/spaces'
    headers = {'Authorization': f'Bearer {wrike_api_token}'}
    response = requests.get(url, headers=headers)
    try:
        return response.json()["data"]
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON response: {e}")
        print(f"Response content: {response.content}")
        raise

# Function to get the space ID from space name
def get_space_id_from_name(space_name, spaces):
    for space in spaces:
        if space["title"] == space_name:
            return space["id"]
    return None

# Function to get details of subtasks
def get_subtask_details(subtask_ids, wrike_api_token):
    headers = {'Authorization': f'Bearer {wrike_api_token}'}
    subtasks = []
    for subtask_id in subtask_ids:
        url = f'{BASE_URL}/tasks/{subtask_id}'
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            subtask_data = response.json()['data'][0]
            subtasks.append(subtask_data)
        else:
            print(f"Failed to get subtask details for subtask {subtask_id}. Status Code: {response.status_code}")
    return subtasks

# Function to get tasks in a folder
def get_tasks_in_folder(folder_id, wrike_api_token):
    headers = {'Authorization': f'Bearer {wrike_api_token}'}
    url = f'{BASE_URL}/folders/{folder_id}/tasks?fields=["subTaskIds","effortAllocation","authorIds","customItemTypeId","responsibleIds","description","hasAttachments","dependencyIds","superParentIds","superTaskIds","metadata","customFields","parentIds","sharedIds","recurrent","briefDescription","attachmentCount"]'
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        tasks = response.json()['data']
        for task in tasks:
            if 'subTaskIds' in task and task['subTaskIds']:
                task['subtasks'] = get_subtask_details(task['subTaskIds'], wrike_api_token)
        return tasks
    else:
        print(f"Failed to get tasks for folder {folder_id}. Status Code: {response.status_code}")
        return []

# Function to get all folders in a workspace
def get_all_folders(workspace_id, wrike_api_token):
    headers = {'Authorization': f'Bearer {wrike_api_token}'}
    url = f'{BASE_URL}/spaces/{workspace_id}/folders'
    response = requests.get(url, headers=headers)
    workspace_data = {'workspace_id': workspace_id, 'folders': []}
    if response.status_code == 200:
        folders = response.json()['data']
        for folder in folders:
            folder_data = folder
            folder_data['tasks'] = get_tasks_in_folder(folder['id'], wrike_api_token)
            workspace_data['folders'].append(folder_data)
    else:
        print(f"Failed to get folders for workspace {workspace_id}. Status Code: {response.status_code}")
    return workspace_data

# Function to save data to JSON
def save_to_json(data, filename='workspace_data.json'):
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

# Main function to execute the export
def wrike_export_json_main():
    # Load configuration and delete task details from Excel file
    excel_file = input("Enter the path to the Excel file: ")
    if not os.path.isfile(excel_file):
        print("File does not exist. Please check the path.")
        exit()
    
    try:
        config_df = pd.read_excel(excel_file, sheet_name="Config", header=1)
    except Exception as e:
        print(f"Failed to read Excel file: {e}")
        exit()

    # Extract token and space name from the Config sheet
    wrike_api_token = config_df.at[0, "Token"]
    space_name = config_df.at[0, "Space to extract data from"]

    # Validate token
    if not validate_token(wrike_api_token):
        client_id = config_df.at[0, "Client ID"]
        client_secret = config_df.at[0, "Client Secret"]
        redirect_url = config_df.at[0, "Redirect URI"]
        wrike_api_token = authenticate_with_oauth2(client_id, client_secret, redirect_url)

    # Get spaces and the ID for the specified space
    spaces_response = get_all_spaces(wrike_api_token)
    workspace_id = get_space_id_from_name(space_name, spaces_response)

    if not workspace_id:
        print(f"Space with name '{space_name}' not found.")
        exit()

    workspace_data = get_all_folders(workspace_id, wrike_api_token)
    save_to_json(workspace_data)

# Execute main function if file is run as a script
if __name__ == "__main__":
    wrike_export_json_main()
