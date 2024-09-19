import requests
import json
import pandas as pd

# Function to read configuration from Excel
def read_config_from_excel(file_path):
    df = pd.read_excel(file_path, sheet_name='Config', header=1)
    config = df.iloc[0].to_dict()  # Convert first row to dictionary
    return config

# Wrike API base URL
BASE_URL = 'https://www.wrike.com/api/v4'

# Function to get the Wrike space ID by name
def get_wrike_space_id(space_name, api_token):
    url = f'{BASE_URL}/spaces'
    headers = {'Authorization': f'Bearer {api_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    spaces = response.json()['data']
    for space in spaces:
        if space['title'].lower() == space_name.lower():
            return space['id']
    return None

# Function to get the details of a space
def get_space_details(space_id, api_token):
    url = f'{BASE_URL}/spaces/{space_id}?fields=["members"]'
    headers = {'Authorization': f'Bearer {api_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()['data'][0]

# Function to create a new Wrike space
def create_new_space(original_space, new_title, api_token):
    url = f'{BASE_URL}/spaces'
    headers = {'Authorization': f'Bearer {api_token}', 'Content-Type': 'application/json'}
    payload = {
        "title": new_title,
        "description": original_space.get("description", ""),
        "accessType": original_space.get("accessType", ""),
        "members": original_space.get("members", [])
    }
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    return response.json()['data'][0]

# Function to get folders in a space
def get_folders_in_space(space_id, api_token):
    url = f'{BASE_URL}/spaces/{space_id}/folders'
    headers = {'Authorization': f'Bearer {api_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()['data']

# Function to get folder by ID from a list of folders
def get_folder_by_id(folder_id, folders):
    for folder in folders:
        if folder['id'] == folder_id:
            return folder
    return None

# Function to get the hierarchy of folder titles
def get_titles_hierarchy(folder_id, folders, path=""):
    folder = get_folder_by_id(folder_id, folders)
    if not folder:
        return []
    current_path = f"{path}/{folder['title']}" if path else folder['title']
    current_entry = {"id": folder_id, "path": current_path, "title": folder["title"]}
    paths = [current_entry]
    for child_id in folder.get("childIds", []):
        child_paths = get_titles_hierarchy(child_id, folders, current_path)
        paths.extend(child_paths)
    return paths

# Function to get tasks in a folder
def get_tasks_in_folder(folder_id, api_token):
    url = f'{BASE_URL}/folders/{folder_id}/tasks?fields=["subTaskIds","effortAllocation","authorIds","customItemTypeId","responsibleIds","description","hasAttachments","dependencyIds","superParentIds","superTaskIds","metadata","customFields","parentIds","sharedIds","recurrent","briefDescription","attachmentCount"]'
    headers = {'Authorization': f'Bearer {api_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()['data']

# Function to get detailed information about a specific task
def get_task_details(task_id, api_token):
    url = f'{BASE_URL}/tasks/{task_id}'
    headers = {'Authorization': f'Bearer {api_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()['data'][0]

# Function to search for the parent task across all folders in the original space
def find_task_across_folders(task_id, folders, api_token):
    for folder in folders:
        tasks = get_tasks_in_folder(folder['id'], api_token)
        for task in tasks:
            if task['id'] == task_id:
                return task
    return None

def create_or_update_task(new_folder_id, task_data, task_map, api_token, folders, folder_mapping, is_subtask=False):
    task_key = task_data['title'] + "|" + str(task_data.get('dates', {}).get('due', ''))

    # Determine if this is a subtask and handle parent task creation first
    super_task_id = None
    if 'superTaskIds' in task_data and task_data['superTaskIds']:
        parent_task_id = task_data['superTaskIds'][0]  # Assuming there's only one parent
        parent_task_key = get_task_key_by_id(parent_task_id, api_token, task_map)

        if parent_task_key not in task_map:
            # Check if the parent task already exists in another folder in the original space
            existing_parent_task = find_task_across_folders(parent_task_id, folders, api_token)
            if existing_parent_task:
                # Parent exists in the original space, so link it to the new space folder
                if existing_parent_task['id'] in folder_mapping:
                    super_task_id = folder_mapping[existing_parent_task['id']]
                else:
                    # Parent exists in the original space but needs to be created in the new space
                    new_parent_folder_id = folder_mapping.get(existing_parent_task['parentIds'][0])
                    parent_task_data = get_task_details(parent_task_id, api_token)
                    parent_task = create_or_update_task(new_parent_folder_id, parent_task_data, task_map, api_token, folders, folder_mapping)
                    super_task_id = parent_task[0]['id'] if parent_task else None
            else:
                # Parent task does not exist anywhere, create it
                parent_task_data = get_task_details(parent_task_id, api_token)
                parent_task = create_or_update_task(new_folder_id, parent_task_data, task_map, api_token, folders, folder_mapping)
                super_task_id = parent_task[0]['id'] if parent_task else None
        else:
            super_task_id = task_map[parent_task_key]

    # Check if the task already exists
    if task_key in task_map:
        existing_task_id = task_map[task_key]

        # Get the details of the existing task to check its current parents
        existing_task_details = get_task_details(existing_task_id, api_token)
        current_parents = existing_task_details.get('parentIds', [])

        # Only update if the new folder is not already a parent
        if not is_subtask and new_folder_id not in current_parents:
            url = f'{BASE_URL}/tasks/{existing_task_id}'
            headers = {
                'Authorization': f'Bearer {api_token}',
                'Content-Type': 'application/json'
            }
            update_payload = {
                "addParents": [new_folder_id]
            }

            print(f"Updating task ID: {existing_task_id} with new folder ID: {new_folder_id}")
            response = requests.put(url, headers=headers, json=update_payload)
            print(f"Response Status: {response.status_code}")
            print(f"Response Data: {response.text}")
            response.raise_for_status()
        else:
            print(f"Task '{task_data['title']}' already exists in the folder. Skipping update.")

    else:
        # Create the task or subtask
        created_task = create_task(
            new_folder_id=new_folder_id if not is_subtask else None,
            task_data=task_data,
            super_task_id=super_task_id,
            api_token=api_token
        )
        task_map[task_key] = created_task[0]['id']

        # Handle subtask creation for the newly created task
        for sub_task_id in task_data.get('subTaskIds', []):
            subtask_data = get_task_details(sub_task_id, api_token)
            create_or_update_task(new_folder_id, subtask_data, task_map, api_token, folders, folder_mapping, is_subtask=True)

        return created_task


# Function to create folders recursively, updating the folder_mapping with original-new folder relationships
def create_folders_recursively(paths, root_folder_id, original_space_name, new_space_name, api_token, folders):
    folder_id_map = {}
    folder_mapping = {}
    new_paths_info = []
    task_map = {}

    for path in paths:
        folder_path = path['path']

        # If the folder_path matches the original space name, skip folder creation but handle tasks
        if folder_path == original_space_name:
            # Process tasks in the root space
            root_tasks = get_tasks_in_folder(path['id'], api_token)
            for task in root_tasks:
                create_or_update_task(new_folder_id=root_folder_id, task_data=task, task_map=task_map, api_token=api_token, folders=folders, folder_mapping=folder_mapping)
            continue

        # If the folder_path is a subfolder, continue with the usual process
        if folder_path.startswith(original_space_name + '/'):
            folder_path = folder_path[len(original_space_name) + 1:]

        path_parts = folder_path.strip('/').split('/')
        parent_id = root_folder_id
        for part in path_parts:
            if part not in folder_id_map:
                new_folder_id = create_folder(part, parent_id, api_token)
                folder_id_map[part] = new_folder_id
                new_path = f"{new_space_name}/{'/'.join(path_parts[:path_parts.index(part)+1])}"
                new_paths_info.append({
                    "original_folder_id": path['id'],
                    "original_folder_title": path['title'],
                    "original_folder_path": path['path'],
                    "new_folder_id": new_folder_id,
                    "new_folder_path": new_path
                })
                folder_mapping[path['id']] = new_folder_id
            parent_id = folder_id_map[part]

        # Process tasks in the current folder
        folder_tasks = get_tasks_in_folder(path['id'], api_token)
        for task in folder_tasks:
            create_or_update_task(new_folder_id=parent_id, task_data=task, task_map=task_map, api_token=api_token, folders=folders, folder_mapping=folder_mapping)

    return new_paths_info







def get_task_key_by_id(task_id, api_token, task_map):
    task_details = get_task_details(task_id, api_token)
    task_key = task_details['title'] + "|" + str(task_details.get('dates', {}).get('due', ''))
    return task_key


# Function to create new tasks or subtasks
def create_task(new_folder_id=None, task_data=None, super_task_id=None, api_token=None):
    url = f'{BASE_URL}/folders/{new_folder_id}/tasks' if new_folder_id else f'{BASE_URL}/tasks'
    headers = {
        'Authorization': f'Bearer {api_token}',
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
    
    if super_task_id:
        payload["superTasks"] = [super_task_id]
    
    print(f"Payload: {json.dumps(payload, indent=2)}")
    
    response = requests.post(url, headers=headers, json=payload)
    print(f"Response status: {response.status_code}")
    print(f"Response data: {response.json()}")
    
    response.raise_for_status()
    return response.json()['data']

# Function to create a folder
def create_folder(title, parent_id, api_token):
    url = f'{BASE_URL}/folders/{parent_id}/folders'
    payload = {'title': title, 'shareds': []}
    headers = {
        'Authorization': f'Bearer {api_token}',
        'Content-Type': 'application/json'
    }
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    return response.json()['data'][0]['id']

def main():
    # Prompt for the Excel file path
    excel_path = input("Please enter the path to the Excel file with configuration: ")
    
    # Read the configuration from the Excel sheet
    config = read_config_from_excel(excel_path)
    
    api_token = config.get('Token')
    space_name = config.get('Space to extract data from')
    new_space_title = config.get('New Space Title')
    
    if not api_token or not space_name or not new_space_title:
        print("Error: Missing necessary configuration details.")
        return
    
    space_id = get_wrike_space_id(space_name, api_token)
    if not space_id:
        print(f"No Wrike space found with the name '{space_name}'.")
        return
    
    original_space = get_space_details(space_id, api_token)
    if not original_space:
        print(f"Could not fetch details for the space '{space_name}'.")
        return
    
    new_space = create_new_space(original_space, new_space_title, api_token)
    if not new_space:
        print(f"Could not create a new space with the title '{new_space_title}'.")
        return
    
    folders = get_folders_in_space(space_id, api_token)
    if not folders:
        print(f"No folders found in the space '{space_name}'.")
        return
    
    all_paths = []
    for folder in folders:
        if "scope" in folder and folder["scope"] == "WsFolder":
            paths = get_titles_hierarchy(folder["id"], folders)
            all_paths.extend(paths)
    
    new_root_folder_id = new_space['id']
    new_paths_info = create_folders_recursively(all_paths, new_root_folder_id, space_name, new_space_title, api_token, folders)


if __name__ == "__main__":
    main()