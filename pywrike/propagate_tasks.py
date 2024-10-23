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


# Function to create a folder in a given space and parent folder
def create_folder_in_space(folder_name, parent_folder_id, access_token):
    endpoint = f'https://www.wrike.com/api/v4/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    data = {
        'title': folder_name,
        'parents': [parent_folder_id]
    }

    response = requests.post(endpoint, headers=headers, json=data)
    if response.status_code == 200:
        new_folder = response.json().get('data', [])[0]
        print(f"Folder '{folder_name}' created successfully in parent folder '{parent_folder_id}'.")
        return new_folder['id']
    else:
        print(f"Failed to create folder '{folder_name}'. Status code: {response.status_code}")
        print(response.text)
        return None

# Function to get the space ID by space name
def get_space_id_by_name(space_name, access_token):
    endpoint = f'https://www.wrike.com/api/v4/spaces'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve spaces. Status code: {response.status_code}")
        print(response.text)
        return None

    spaces = response.json().get('data', [])
    for space in spaces:
        if space['title'] == space_name:
            return space['id']

    print(f"Space with name '{space_name}' not found.")
    return None

# Function to get the ID of a folder by its path within a specific space
def get_folder_id_by_path(folder_path, space_id, access_token):
    folder_names = folder_path.split('\\')  # Split the folder path into individual folders
    parent_folder_id = None  # Start with no parent folder

    # Iterate through folder names and get each folder's ID in the hierarchy
    for folder_name in folder_names:
        if parent_folder_id:
            # Fetch subfolder of the current parent folder
            parent_folder_id = get_subfolder_id_by_name(parent_folder_id, folder_name, access_token)
        else:
            # Fetch top-level folder in the given space
            parent_folder_id = get_folder_in_space_by_name(folder_name, space_id, access_token)
        
        if not parent_folder_id:
            print(f"Folder '{folder_name}' not found.")
            return None

    return parent_folder_id


# Function to get a folder within a space by its name
def get_folder_in_space_by_name(folder_name, space_id, access_token):
    endpoint = f'https://www.wrike.com/api/v4/spaces/{space_id}/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve folders in space {space_id}. Status code: {response.status_code}")
        print(response.json())  # Print the full error response
        return None

    folders = response.json().get('data', [])
    for folder in folders:
        if folder['title'] == folder_name:
            return folder['id']

    print(f"Folder with name '{folder_name}' not found in space {space_id}.")
    return None


# Function to get subfolder ID within a parent folder by name
def get_subfolder_id_by_name(parent_folder_id, subfolder_name, access_token):
    endpoint = f'https://www.wrike.com/api/v4/folders/{parent_folder_id}/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve subfolders in folder {parent_folder_id}. Status code: {response.status_code}")
        print(response.json())
        return None

    subfolders = response.json().get('data', [])
    for subfolder in subfolders:
        if subfolder['title'] == subfolder_name:
            return subfolder['id']

    print(f"Subfolder '{subfolder_name}' not found in folder {parent_folder_id}.")
    return None


# Function to get the IDs of all tasks in a folder
def get_all_tasks_in_folder(folder_id, access_token):
    endpoint = f'https://www.wrike.com/api/v4/folders/{folder_id}/tasks'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve tasks. Status code: {response.status_code}")
        print(response.text)
        return None

    tasks = response.json().get('data', [])
    return tasks

# Function to get the ID of a task by its title in a folder
def get_task_id_by_title(folder_id, task_title, access_token):
    tasks = get_all_tasks_in_folder(folder_id, access_token)
    for task in tasks:
        if task['title'] == task_title:
            return task['id']
    print(f"Task with title '{task_title}' not found.")
    return None

def create_task(folder_id, task_data, access_token):
    url = f'https://www.wrike.com/api/v4/folders/{folder_id}/tasks'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    task_dates = task_data.get('dates', {})
    start_date = task_dates.get('start', "")
    due_date = task_dates.get('due', "")
    type_date = task_dates.get('type', "")
    duration_date = task_dates.get("duration", "")
        
    dates = {}
    if start_date:
        dates["start"] = start_date
    if due_date:
        dates["due"] = due_date
    if type_date:
        dates["type"] = type_date
    if duration_date:
        dates["duration"] = duration_date
  
    effortAllocation = task_data.get('effortAllocation', {})
    
    effort_allocation_payload = {}
    if effortAllocation.get('mode') in ['Basic', 'Flexible', 'None', 'FullTime']:  # Check valid modes
        effort_allocation_payload['mode'] = effortAllocation.get('mode')
        if 'totalEffort' in effortAllocation:
            effort_allocation_payload['totalEffort'] = effortAllocation['totalEffort']
        if 'allocatedEffort' in effortAllocation:
            effort_allocation_payload['allocatedEffort'] = effortAllocation['allocatedEffort']
        if 'dailyAllocationPercentage' in effortAllocation:
            effort_allocation_payload['dailyAllocationPercentage'] = effortAllocation['dailyAllocationPercentage']

    payload = {
        "title": task_data.get("title", ""),
        "description": task_data.get("description", ""),
        "responsibles": task_data.get("responsibleIds", []),        
        "customStatus": task_data.get("customStatusId", ""),
        "importance": task_data.get("importance", ""),
        "metadata": task_data.get("metadata", []),
        "customFields": task_data.get("customFields", []) 
    }
    
    if dates:
        payload["dates"] = dates
    
    if effortAllocation:
        payload["effortAllocation"] = effort_allocation_payload
            
    response = requests.post(url, headers=headers, json=payload)
        
    response.raise_for_status()
    return response.json()['data']

# Function to get task details by task ID
def get_task_details(task_id, access_token):
    endpoint = f'https://www.wrike.com/api/v4/tasks/{task_id}'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve task details. Status code: {response.status_code}")
        print(response.text)
        return None

    task = response.json().get('data', [])[0]
    return task

# Modified main function to create new tasks instead of updating existing ones
def propagate_tasks_main():
    # Ask the user to input the path of the Excel file
    excel_file = input("Please enter the full path to the Excel file: ")

    # Load the access token from the 'Config' sheet
    config_df = pd.read_excel(excel_file, sheet_name='Config', header=1)
    access_token = config_df.at[0, 'Token']

    # Validate the token
    if not validate_token(access_token):
        # If the token is invalid, authenticate using OAuth2 and update the access_token
        wrike = OAuth2Gateway1(excel_filepath=excel_file)
        access_token = wrike._create_auth_info()  # Perform OAuth2 authentication
        print(f"New access token obtained: {access_token}")
        
    # Load the propagation details from the file (assuming 'Space Name', 'Source Path', 'Task Title', 'Destination Path' columns)
    propagate_df = pd.read_excel(excel_file, sheet_name='Propagate')

    # Iterate over the rows in the 'Propagate' sheet
    for index, row in propagate_df.iterrows():
        space_name = row['Space Name']
        source_path = row['Source Path']
        task_title = row['Task Title']

        # Check if 'Destination Path' is valid (non-NaN)
        if pd.isna(row['Destination Path']):
            print(f"Skipping row {index} due to missing 'Destination Path'.")
            continue

        destination_path = row['Destination Path'].strip()

        # Get the space ID by space name
        space_id = get_space_id_by_name(space_name, access_token)
        
        # Get the ID of the source folder
        source_folder_id = get_folder_id_by_path(source_path, space_id, access_token)
        if not source_folder_id:
            continue

        # Get the destination folder ID
        destination_folder_id = get_folder_id_by_path(destination_path, space_id, access_token)
        if not destination_folder_id:
            # Create the destination folder if it doesn't exist
            destination_folder_name = destination_path.split('\\')[-1]
            parent_folder_id = get_folder_id_by_path('\\'.join(destination_path.split('\\')[:-1]), space_id, access_token)
            if not parent_folder_id:
                print(f"Parent folder for '{destination_folder_name}' not found. Skipping.")
                continue
            destination_folder_id = create_folder_in_space(destination_folder_name, parent_folder_id, access_token)

        if pd.isna(task_title):
            # Create subfolder and replicate the tasks
            new_subfolder_id = create_folder_in_space(source_path.split('\\')[-1], destination_folder_id, access_token)
            tasks = get_all_tasks_in_folder(source_folder_id, access_token)
            for task in tasks:
                create_task(new_subfolder_id, task, access_token)  # Pass the new subfolder ID here
                
        else:
            # Create a new task instead of updating an existing one
            task_id = get_task_id_by_title(source_folder_id, task_title, access_token)
            if task_id:
                task_details = get_task_details(task_id, access_token)
                create_task(destination_folder_id, task_details, access_token)  