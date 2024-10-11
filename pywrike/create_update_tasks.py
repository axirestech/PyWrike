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
    folder_names = folder_path.split('\\')
    parent_folder_id = get_folder_id_by_name(space_id, folder_names[0], access_token)
    if not parent_folder_id:
        return None

    for folder_name in folder_names[1:]:
        parent_folder_id = get_or_create_subfolder(parent_folder_id, folder_name, access_token)
        if not parent_folder_id:
            print(f"Subfolder '{folder_name}' not found in space '{space_id}'")
            return None

    return parent_folder_id

# Function to get the ID of a folder by its name within a specific space
def get_folder_id_by_name(space_id, folder_name, access_token):
    endpoint = f'https://www.wrike.com/api/v4/spaces/{space_id}/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve folders in space {space_id}. Status code: {response.status_code}")
        print(response.text)
        return None

    folders = response.json().get('data', [])
    for folder in folders:
        if folder['title'] == folder_name:
            return folder['id']

    print(f"Folder with name '{folder_name}' not found in space {space_id}.")
    return None
       
def get_all_folders_in_space(space_id, access_token):
    all_folders = []
    folders_to_process = [space_id]  # Start with the root space
    processed_folders = set()  # Set to track processed folder IDs

    while folders_to_process:
        parent_folder_id = folders_to_process.pop()

        # Check if the folder has already been processed
        if parent_folder_id in processed_folders:
            continue
        
        print(f"[DEBUG] Fetching folders for parent folder ID: {parent_folder_id}")
        endpoint = f'https://www.wrike.com/api/v4/folders/{parent_folder_id}/folders'
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        response = requests.get(endpoint, headers=headers)
        
        if response.status_code != 200:
            print(f"[DEBUG] Failed to retrieve folders. Status code: {response.status_code}")
            print(response.text)
            continue

        folders = response.json().get('data', [])
        print(f"[DEBUG] Found {len(folders)} folders under parent folder ID: {parent_folder_id}")
        all_folders.extend(folders)

        # Mark the folder as processed
        processed_folders.add(parent_folder_id)

        # Add child folders to the list to process them in the next iterations
        for folder in folders:
            folder_id = folder['id']
            if folder_id not in processed_folders:
                folders_to_process.append(folder_id)

    return all_folders

def get_all_tasks_in_space(space_id, access_token):
    folders = get_all_folders_in_space(space_id, access_token)
    all_tasks = []

    for folder in folders:
        folder_id = folder['id']
        print(f"[DEBUG] Fetching tasks for folder ID: {folder_id}")
        tasks = get_tasks_by_folder_id(folder_id, access_token)
        print(f"[DEBUG] Found {len(tasks)} tasks in folder ID: {folder_id}")
        all_tasks.extend(tasks)

    return all_tasks

# Function to get or create a subfolder by its name and parent folder ID
def get_or_create_subfolder(parent_folder_id, subfolder_name, access_token):
    subfolder_id = get_subfolder_id_by_name(parent_folder_id, subfolder_name, access_token)
    if not subfolder_id:
        subfolder_id = create_subfolder(parent_folder_id, subfolder_name, access_token)
    return subfolder_id

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

# Function to create a subfolder in the parent folder
def create_subfolder(parent_folder_id, subfolder_name, access_token):
    endpoint = f'https://www.wrike.com/api/v4/folders/{parent_folder_id}/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    payload = {
        "title": subfolder_name,
        "shareds": []  # Adjust shared settings as needed
    }

    response = requests.post(endpoint, headers=headers, json=payload)
    if response.status_code == 200:
        subfolder_id = response.json().get('data', [])[0].get('id')
        print(f"Subfolder '{subfolder_name}' created successfully in parent folder '{parent_folder_id}'")
        return subfolder_id
    else:
        print(f"Failed to create subfolder '{subfolder_name}' in parent folder '{parent_folder_id}'. Status code: {response.status_code}")
        print(response.text)
        return None

def get_tasks_in_space(space_id, access_token):
    endpoint = f'https://www.wrike.com/api/v4/spaces/{space_id}/tasks'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"[DEBUG] Failed to retrieve tasks in space {space_id}. Status code: {response.status_code}")
        print(response.text)
        return []

    tasks = response.json().get('data', [])
    print(f"[DEBUG] Retrieved {len(tasks)} tasks in space {space_id}.")
    for task in tasks:
        print(f"[DEBUG] Task ID: {task['id']}, Title: '{task['title']}', Parent Folders: {task.get('parentIds', [])}")
    return tasks

# Function to get all tasks by folder ID
def get_tasks_by_folder_id(folder_id, access_token):
    endpoint = f'https://www.wrike.com/api/v4/folders/{folder_id}/tasks/?fields=["subTaskIds", "effortAllocation","authorIds","customItemTypeId","responsibleIds","description","hasAttachments","dependencyIds","superParentIds","superTaskIds","subTaskIds","metadata","customFields","parentIds","sharedIds","recurrent","briefDescription","attachmentCount"]'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve tasks in folder {folder_id}. Status code: {response.status_code}")
        print(response.text)
        return []

    return response.json().get('data', [])

# Function to get the ID of a task by its title and folder ID
def get_task_id_by_title(task_title, folder_id, access_token):
    tasks = get_tasks_by_folder_id(folder_id, access_token)
    for task in tasks:
        if task['title'] == task_title:
            return task['id']
    print(f"Task with title '{task_title}' not found in folder '{folder_id}'.")
    return None

# Function to lookup the responsible ID by first name, last name, and email
def get_responsible_id_by_name_and_email(first_name, last_name, email, access_token):
    endpoint = f'https://www.wrike.com/api/v4/contacts'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve contacts. Status code: {response.status_code}")
        print(response.text)
        return None

    contacts = response.json().get('data', [])
    for contact in contacts:
        if contact.get('firstName', '') == first_name and contact.get('lastName', '') == last_name and contact.get('profiles', [])[0].get('email', '') == email:
            return contact['id']

    return None

def cache_subtasks_from_tasks(cached_tasks, access_token):
    new_subtasks = []

    # Loop through all cached tasks to find those with 'subtaskIds'
    for task in cached_tasks:
        subtask_ids = task.get('subTaskIds')
        if subtask_ids:
            if isinstance(subtask_ids, list):
                for subtask_id in subtask_ids:
                    print(f"[DEBUG] Found subtaskId '{subtask_id}' in task '{task['title']}'. Fetching subtask details.")
                    
                    # Fetch subtask details
                    subtask_response = get_task_by_id(subtask_id, access_token)
                    
                    # Print the entire response for debugging
                    print(f"[DEBUG] Subtask response fetched: {subtask_response}")
                    
                    # Extract the subtask details from the response
                    if 'data' in subtask_response and len(subtask_response['data']) > 0:
                        subtask_details = subtask_response['data'][0]
                        new_subtasks.append(subtask_details)
                        print(f"[DEBUG] Subtask '{subtask_details.get('title', 'Unknown Title')}' added to cache.")
                    else:
                        print(f"[DEBUG] No subtask details found for subtaskId '{subtask_id}'.")
            else:
                print(f"[DEBUG] Unexpected type for 'subtaskIds': {type(subtask_ids)}. Expected a list.")
    
    # Add the new subtasks to the global cached_tasks list
    cached_tasks.extend(new_subtasks)
    print(f"[DEBUG] Cached {len(new_subtasks)} new subtasks.")

def get_custom_fields(access_token):
    endpoint = 'https://www.wrike.com/api/v4/customfields'
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    
    response = requests.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        custom_fields_data = response.json()
        return {field['title']: field['id'] for field in custom_fields_data['data']}  # Map custom field names to their IDs
    else:
        print(f"Failed to fetch custom fields. Status code: {response.status_code}")
        print(response.text)
        return {}

def map_excel_headings_to_custom_fields(headings, wrike_custom_fields):
    mapped_custom_fields = {}

    for heading in headings:
        clean_heading = heading.strip()  # Remove leading/trailing spaces
        if clean_heading in wrike_custom_fields:
            mapped_custom_fields[clean_heading] = wrike_custom_fields[clean_heading]
        else:
            print(f"[WARNING] No match found for Excel heading '{heading}' in Wrike custom fields")
    return mapped_custom_fields

    
def get_task_by_id(task_id, access_token):
    endpoint = f'https://www.wrike.com/api/v4/tasks/{task_id}'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to fetch task with ID '{task_id}': {response.status_code} {response.text}")
        return None

#Function to update task
def update_task_with_tags(task_id, new_folder_id, access_token):
    endpoint = f'https://www.wrike.com/api/v4/tasks/{task_id}'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    # Retrieve current task details to get existing tags
    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve task details for task '{task_id}'. Status code: {response.status_code}")
        print(response.text)
        return

    task_data = response.json().get('data', [])[0]
    existing_tags = task_data.get('parentIds', [])

    # Add the new folder ID if it's not already tagged
    if new_folder_id not in existing_tags:
        existing_tags.append(new_folder_id)

    # Prepare the payload with updated tags
    payload = {
        "addParents": [new_folder_id]  # Update to add only new folder as tag
    }

    # Update the task with new tags
    response = requests.put(endpoint, headers=headers, json=payload)
    if response.status_code == 200:
        print(f"Task '{task_data['title']}' updated successfully with new folder tags.")
    else:
        print(f"Failed to update task '{task_data['title']}'. Status code: {response.status_code}")
        print(response.text)

def update_subtask_with_parent(subtask_id, new_parent_task_id, access_token):
    endpoint = f'https://www.wrike.com/api/v4/tasks/{subtask_id}'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        print(f"Failed to retrieve subtask details for '{subtask_id}'. Status code: {response.status_code}")
        print(response.text)
        return

    subtask_data = response.json().get('data', [])[0]
    existing_parents = subtask_data.get('parentIds', [])

    if new_parent_task_id not in existing_parents:
        existing_parents.append(new_parent_task_id)

    payload = {
        "addSuperTasks": [new_parent_task_id]
    }

    response = requests.put(endpoint, headers=headers, json=payload)
    if response.status_code == 200:
        print(f"Subtask '{subtask_data['title']}' updated successfully with parent task.")
    else:
        print(f"Failed to update subtask '{subtask_data['title']}'. Status code: {response.status_code}")
        print(response.text)

def create_task_in_folder(folder_id, space_id, task_data, access_token):
    global cached_tasks  # Use the global variable for caching tasks
    print(f"[DEBUG] Starting to create/update task '{task_data['title']}' in folder '{folder_id}' within space '{space_id}'.")

    responsible_ids = []
    for first_name, last_name, email in zip(task_data.get("first_names", []), task_data.get("last_names", []), task_data.get("emails", [])):
        responsible_id = get_responsible_id_by_name_and_email(first_name, last_name, email, access_token)
        if responsible_id:
            responsible_ids.append(responsible_id)
        else:
            print(f"[DEBUG] Responsible user '{first_name} {last_name}' with email '{email}' not found.")
            user_input = input(f"User '{first_name} {last_name}' with email '{email}' not found. Would you like to (1) Correct the information, or (2) Proceed without assigning this user? (Enter 1/2): ").strip()
            if user_input == '1':
                first_name = input("Enter the correct first name: ").strip()
                last_name = input("Enter the correct last name: ").strip()
                email = input("Enter the correct email: ").strip()
                responsible_id = get_responsible_id_by_name_and_email(first_name, last_name, email, access_token)
                if responsible_id:
                    responsible_ids.append(responsible_id)
                else:
                    print(f"[DEBUG] User '{first_name} {last_name}' with email '{email}' still not found. Creating the task without assignee.")
            elif user_input == '2':
                print(f"[DEBUG] Proceeding without assigning user '{first_name} {last_name}'.")

    existing_tasks = get_tasks_by_folder_id(folder_id, access_token)
    print(f"[DEBUG] Retrieved {len(existing_tasks)} tasks in folder '{folder_id}'.")

    existing_task = next((task for task in existing_tasks if task['title'].strip().lower() == task_data['title'].strip().lower()), None)
    if existing_task:
        print(f"[DEBUG] Task '{task_data['title']}' already exists in the folder '{folder_id}'.")
        return  # Task already exists in the folder, do nothing

    existing_tasks_space = cached_tasks
    print(f"[DEBUG] Checking for task '{task_data['title']}' in entire space '{space_id}'.")

    existing_task_space = next((task for task in existing_tasks_space if task['title'].strip().lower() == task_data['title'].strip().lower()), None)
    if existing_task_space:
        print(f"[DEBUG] Task '{task_data['title']}' found in another folder in the space.")
        existing_task_id = existing_task_space['id']
        update_task_with_tags(existing_task_id, folder_id, access_token)
        print(f"[DEBUG] Updated task '{task_data['title']}' with new folder tag '{folder_id}'.")
    else:
        print(f"[DEBUG] Task '{task_data['title']}' does not exist in space '{space_id}'. Creating a new task.")
        new_task = create_task(folder_id, task_data, responsible_ids, access_token)
        # Update the cache with the newly created task
        # Ensure the new task is not None and has an ID
        if new_task and 'id' in new_task:
            cached_tasks.append(new_task)
            print(f"[DEBUG] Added newly created task '{new_task['title']}' with ID '{new_task['id']}' to cache.")
        else:
            print(f"[DEBUG] Failed to create the task or retrieve task ID.")

def get_subtasks_by_task_id(parent_task_id, access_token):
    endpoint = f'https://www.wrike.com/api/v4/tasks/{parent_task_id}'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    try:
        response = requests.get(endpoint, headers=headers)
        response.raise_for_status()
        data = response.json()
        return data.get('subTaskIds', [])  # Return the list of subtasks
    except requests.exceptions.RequestException as e:
        print(f"Failed to retrieve subtasks for parent task '{parent_task_id}': {e}")
        return []

def create_subtask_in_parent_task(parent_task_id, space_id, subtask_data, access_token):
    global cached_tasks  # Use the global cache for tasks and subtasks
    print(f"[DEBUG] Starting to create/update subtask '{subtask_data['title']}' under parent task '{parent_task_id}' within space '{space_id}'.")

    responsible_ids = []
    for first_name, last_name, email in zip(subtask_data.get("first_names", []), subtask_data.get("last_names", []), subtask_data.get("emails", [])):
        responsible_id = get_responsible_id_by_name_and_email(first_name, last_name, email, access_token)
        if responsible_id:
            responsible_ids.append(responsible_id)
        else:
            print(f"[DEBUG] Responsible user '{first_name} {last_name}' with email '{email}' not found.")
            user_input = input(f"User '{first_name} {last_name}' with email '{email}' not found. Would you like to (1) Correct the information, or (2) Proceed without assigning this user? (Enter 1/2): ").strip()
            if user_input == '1':
                first_name = input("Enter the correct first name: ").strip()
                last_name = input("Enter the correct last name: ").strip()
                email = input("Enter the correct email: ").strip()
                responsible_id = get_responsible_id_by_name_and_email(first_name, last_name, email, access_token)
                if responsible_id:
                    responsible_ids.append(responsible_id)
                else:
                    print(f"[DEBUG] User '{first_name} {last_name}' with email '{email}' still not found. Creating the subtask without assignee.")
            elif user_input == '2':
                print(f"[DEBUG] Proceeding without assigning user '{first_name} {last_name}'.")

    # Check cached tasks for the subtask under the parent task
    existing_subtask = next((task for task in cached_tasks 
                             if task['title'].strip().lower() == subtask_data['title'].strip().lower() 
                             and task.get('supertaskId') == parent_task_id), None)

    if existing_subtask:
        print(f"[DEBUG] Subtask '{subtask_data['title']}' already exists under parent task '{parent_task_id}'.")
        return  # Subtask already exists, no further action

    # Retrieve all subtasks under the parent task from API
    existing_subtasks = get_subtasks_by_task_id(parent_task_id, access_token)
    print(f"[DEBUG] Retrieved {len(existing_subtasks)} subtasks under parent task '{parent_task_id}'.")

    # Check if the subtask already exists under the parent task
    existing_subtask = next((subtask for subtask in existing_subtasks if subtask['title'].strip().lower() == subtask_data['title'].strip().lower()), None)
    if existing_subtask:
        print(f"[DEBUG] Subtask '{subtask_data['title']}' already exists under the parent task '{parent_task_id}'.")
        return  # Subtask already exists under the parent, do nothing

    # Check for the subtask in the entire space (cached tasks)
    print(f"[DEBUG] Checking for subtask '{subtask_data['title']}' in the entire space '{space_id}'.")
    existing_subtask_space = next((task for task in cached_tasks 
                                   if task['title'].strip().lower() == subtask_data['title'].strip().lower() 
                                   and task.get('supertaskId') != parent_task_id), None)

    if existing_subtask_space:
        print(f"[DEBUG] Subtask '{subtask_data['title']}' found in another parent task within the space.")
        existing_subtask_id = existing_subtask_space['id']
        update_subtask_with_parent(existing_subtask_id, parent_task_id, access_token)
        print(f"[DEBUG] Updated subtask '{subtask_data['title']}' with new parent task '{parent_task_id}'.")
    else:
        print(f"[DEBUG] Subtask '{subtask_data['title']}' does not exist in space '{space_id}'. Creating a new subtask.")
        new_subtask = create_subtask(parent_task_id, subtask_data, responsible_ids, access_token)
        
        # Update the cache with the newly created subtask
        if new_subtask and 'id' in new_subtask:
            cached_tasks.append(new_subtask)
            print(f"[DEBUG] Added newly created subtask '{new_subtask['title']}' with ID '{new_subtask['id']}' to cache.")
        else:
            print(f"[DEBUG] Failed to create the subtask or retrieve subtask ID.")

def create_task(folder_id, task_data, responsible_ids, access_token):
    endpoint = f'https://www.wrike.com/api/v4/folders/{folder_id}/tasks'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    payload = {
        "title": task_data.get("title", ""),
        "responsibles": responsible_ids
    }
    
    if "importance" in task_data and pd.notna(task_data["importance"]) and task_data["importance"]:
        payload["importance"] = task_data["importance"]
    
    if "description" in task_data and pd.notna(task_data["description"]) and task_data["description"]:
        payload["description"] = task_data["description"]
    
    if pd.notna(task_data.get("start_date")) and pd.notna(task_data.get("end_date")):
        payload["dates"] = {
            "start": task_data.get("start_date"),
            "due": task_data.get("end_date")
        }
 
    # Get custom fields from API
    custom_fields = get_custom_fields(access_token)

    # Map Excel headings to Wrike custom fields
    mapped_custom_fields = map_excel_headings_to_custom_fields(task_df.columns, custom_fields)
    print (mapped_custom_fields)

    # Then retrieve the custom fields
    custom_fields_payload = []
    for field_name, field_id in mapped_custom_fields.items():
        field_value = task_data.get(field_name) 
        print(f"[DEBUG] Task Data: {task_data}")
        print(f"Retrieving '{field_name}' from task data: '{field_value}'")  # Debugging line
        if pd.notna(field_value):
            custom_fields_payload.append({
                "id": field_id,
                "value": str(field_value)  # Wrike expects the custom field values as strings
            })
        if custom_fields_payload:
            payload["customFields"] = custom_fields_payload

    
    print(f"[DEBUG] Final payload being sent: {payload}")
    response = requests.post(endpoint, headers=headers, json=payload)
    
    if response.status_code == 200:
        task_data_response = response.json()  # Parse the JSON response to get the task data
        print(f"[DEBUG] Response JSON: {task_data_response}")  # Print out the entire response for inspection

        # Check if the expected data structure is present
        if 'data' in task_data_response and len(task_data_response['data']) > 0:
            task_data = task_data_response['data'][0]
            print(f"Task '{task_data['title']}' created successfully in folder '{folder_id}'")
            return task_data  # Return the first task in the data list
        else:
            print(f"[ERROR] Unexpected response structure: {task_data_response}")
            return None  # Handle the unexpected structure gracefully
    else:
        print(f"Failed to create task '{task_data.get('title', '')}' in folder '{folder_id}'. Status code: {response.status_code}")
        print(response.text)
        return None  # Return None if the task creation fails

def create_subtask(parent_task_id, subtask_data, responsible_ids, access_token):
    endpoint = f'https://www.wrike.com/api/v4/tasks'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    payload = {
        "title": subtask_data.get("title", ""),
        "responsibles": responsible_ids,
        "superTasks": [parent_task_id],
    }

    if "importance" in subtask_data and pd.notna(subtask_data["importance"]) and subtask_data["importance"]:
        payload["importance"] = subtask_data["importance"]

    if "description" in subtask_data and pd.notna(subtask_data["description"]) and subtask_data["description"]:
        payload["description"] = subtask_data["description"]

    if pd.notna(subtask_data.get("start_date")) and pd.notna(subtask_data.get("end_date")):
        payload["dates"] = {
            "start": subtask_data.get("start_date"),
            "due": subtask_data.get("end_date")
        }

    # Add custom fields dynamically
    custom_fields_payload = []
    for field_name, field_value in task_data.items():
        if field_name in custom_fields and pd.notna(field_value):
            custom_fields_payload.append({
                "id": custom_fields[field_name],
                "value": str(field_value)  # Wrike requires custom field values to be strings
            })
    
    if custom_fields_payload:
        payload["customFields"] = custom_fields_payload
   
    response = requests.post(endpoint, headers=headers, json=payload)

    if response.status_code == 200:
        subtask_data_response = response.json()
        print(f"Subtask '{subtask_data['title']}' created successfully under parent task '{parent_task_id}'")
        return subtask_data_response['data'][0] if 'data' in subtask_data_response else None
    else:
        print(f"Failed to create subtask '{subtask_data.get('title', '')}'. Status code: {response.status_code}")
        print(response.text)
        return None

def create_update_tasks_main():
    global task_df
    excel_file = input("Enter the path to the Excel file: ")
    try:
        config_df = pd.read_excel(excel_file, sheet_name="Config", header=1)
        task_df = pd.read_excel(excel_file, sheet_name="Tasks")
    except Exception as e:
        print(f"Failed to read Excel file: {e}")
        return

    headings = list(task_df.columns)

    access_token = config_df.at[0, "Token"]
        # Validate the token
    if not validate_token(access_token):
        # If the token is invalid, authenticate using OAuth2 and update the access_token
        wrike = OAuth2Gateway1(excel_filepath=excel_file)
        access_token = wrike._create_auth_info()  # Perform OAuth2 authentication
        print(f"New access token obtained: {access_token}")

    cached_tasks_by_space = {}  # Dictionary to store cached tasks by space

    if task_df.empty:
        print("No task details provided.")
        return

    for _, row in task_df.iterrows():
        space_name = row.get("Space Name", "")
        folder_path = row.get("Folder Path", "")
        parent_task_title = row.get("Parent Task Title", "")
        parent_task_path = row.get("Folder Path", "")

        if not space_name:
            print(f"Space name is missing for task '{row.get('Title', '')}'. Skipping this task.")
            continue

        if not folder_path:
            print(f"Folder path is missing for task '{row.get('Title', '')}'. Skipping this task.")
            continue
        
        space_id = get_space_id_by_name(space_name, access_token)
        if not space_id:
            print(f"Space '{space_name}' does not exist. Skipping this task.")
            continue

        # Check if tasks for this space have already been cached
        if space_name not in cached_tasks_by_space:
            print(f"[DEBUG] Caching all tasks in the space '{space_name}'.")
            cached_tasks_by_space[space_name] = get_all_tasks_in_space(space_id, access_token)
            cache_subtasks_from_tasks(cached_tasks_by_space[space_name], access_token)

        else:
            print(f"[DEBUG] Using cached tasks for space '{space_name}'.")
            

        # Set the cached tasks for the current space
        global cached_tasks
        cached_tasks = cached_tasks_by_space[space_name]
        

        folder_id = get_folder_id_by_path(folder_path, space_id, access_token)
        if not folder_id:
            user_input = input(f"Folder path '{folder_path}' does not exist. Would you like to create it? (yes/no): ").strip().lower()
            if user_input != 'yes':
                print(f"Task '{row.get('Title', '')}' creation skipped.")
                continue
            folder_id = create_folder_by_path(folder_path, space_id, access_token)
            if not folder_id:
                print(f"Failed to create or locate folder path '{folder_path}'. Skipping this task.")
                continue
        
        task_data = {
            "title": row.get("Title", ""),
            "description": row.get("Description", ""),
            "first_names": row.get("FirstName", "").split(',') if pd.notna(row.get("FirstName")) else [],
            "last_names": row.get("LastName", "").split(',') if pd.notna(row.get("LastName")) else [],
            "emails": row.get("Email", "").split(',') if pd.notna(row.get("Email")) else [],
            "importance": row.get("Priority", ""),
            "start_date": row.get("Start Date"),
            "end_date": row.get("End Date")
        }

                # Dynamically add any custom field columns to task_data
        for col_name in row.index:
            if col_name not in task_data:  # If this column is not already part of task_data
                field_value = row.get(col_name, "")
                if pd.notna(field_value):
                    task_data[col_name] = str(field_value).strip() if isinstance(field_value, str) else str(field_value)
                else:
                    task_data[col_name] = ""


        # Check if parent task title is provided (Subtask scenario)
        if pd.notna(parent_task_title):
            if parent_task_path:
                parent_folder_id = get_folder_id_by_path(parent_task_path, space_id, access_token)
                if parent_folder_id:
                    parent_task_id = get_task_id_by_title(parent_task_title, parent_folder_id, access_token)
                    if parent_task_id:
                        print(f"[DEBUG] Creating subtask '{task_data['title']}' under parent task '{parent_task_title}'.")
                        create_subtask_in_parent_task(parent_task_id, space_id, task_data, access_token)
                    else:
                        print(f"Parent task '{parent_task_title}' not found. Skipping this subtask.")
                else:
                    print(f"Parent task folder path '{parent_task_path}' not found. Skipping this subtask.")
            else:
                print(f"Parent task path is missing for subtask '{row.get('Title', '')}'. Skipping this subtask.")
        else:
            # No parent task title provided, so create a standalone task
            print(f"[DEBUG] Creating standalone task '{task_data['title']}' in folder '{folder_path}'.")
            create_task_in_folder(folder_id, space_id, task_data, access_token)

if __name__ == '__create_update_tasks_main__':
    create_update_tasks_main()