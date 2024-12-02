import requests
import json
import time
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from bs4 import BeautifulSoup
import pandas as pd
import os
from PyWrike.gateways import OAuth2Gateway1
import numpy as np

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

# Function to create a new project in Wrike
def create_wrike_project(access_token, parent_folder_id, project_title, responsible_id, start_date, end_date):
    if not all([project_title, responsible_id, start_date, end_date]):
        print("Missing required project details.")
        return None

    endpoint = f'https://www.wrike.com/api/v4/folders/{parent_folder_id}/folders'
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}

    data = {
        'title': project_title,
        'project': {
            'ownerIds': [responsible_id],
            'startDate': start_date,
            'endDate': end_date
        }
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
    
# Function to get the ID of a folder by its path within a specific space
def get_folder_id_by_path(folder_path, space_id, access_token):
    folder_names = folder_path.split('\\')
    parent_folder_id = get_folder_id_in_space_by_name(space_id, folder_names[0], access_token)
    if not parent_folder_id:
        return None

    for folder_name in folder_names[1:]:
        parent_folder_id = get_or_create_subfolder(parent_folder_id, folder_name, access_token)
        if not parent_folder_id:
            print(f"Subfolder '{folder_name}' not found in space '{space_id}'")
            return None

    return parent_folder_id

# Function to get the ID of a folder by its name within a specific space
def get_folder_id_in_space_by_name(space_id, folder_name, access_token):
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

# Function to retrieve custom fields and filter by space
def get_custom_fields_by_space(access_token, space_id):
    endpoint = 'https://www.wrike.com/api/v4/customfields'
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    
    response = requests.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        custom_fields_data = response.json()

        # Create a mapping of custom field title to a list of {id, spaces} dicts
        custom_fields = {}
        for field in custom_fields_data['data']:
            field_spaces = field.get('spaceId', [])  # Get the spaces where the custom field is applied
            if space_id in field_spaces:  # Only add custom fields that belong to the specific space
                custom_fields[field['title']] = {'id': field['id'], 'spaces': field_spaces}
        
        return custom_fields
    else:
        print(f"Failed to fetch custom fields. Status code: {response.status_code}")
        print(response.text)
        return {}

# Function to map Excel headings to custom fields by name and space
def map_excel_headings_to_custom_fields(headings, wrike_custom_fields):
    mapped_custom_fields = {}

    for heading in headings:
        clean_heading = heading.strip()  # Remove leading/trailing spaces
        if clean_heading in wrike_custom_fields:
            mapped_custom_fields[clean_heading] = wrike_custom_fields[clean_heading]['id']
        else:
            print(f"[WARNING] No match found for Excel heading '{heading}' in Wrike custom fields")
    
    return mapped_custom_fields

# Task creation function with space-specific custom field mapping
def create_task(folder_id, space_id, task_data, responsible_ids, access_token):
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
            "start": task_data.get("start_date").isoformat() if isinstance(task_data.get("start_date"), pd.Timestamp) else task_data.get("start_date"),
            "due": task_data.get("end_date").isoformat() if isinstance(task_data.get("end_date"), pd.Timestamp) else task_data.get("end_date")
        }


    # Get custom fields from API specific to the space
    custom_fields = get_custom_fields_by_space(access_token, space_id)

    # Map Excel headings to Wrike custom fields
    mapped_custom_fields = map_excel_headings_to_custom_fields(task_data.keys(), custom_fields)
    print(f"[DEBUG] Mapped Custom Fields: {mapped_custom_fields}")

    # Create custom fields payload
    custom_fields_payload = []
    for field_name, field_id in mapped_custom_fields.items():
        field_value = task_data.get(field_name) 
        print(f"[DEBUG] Retrieving '{field_name}' from task data: '{field_value}'") 
        
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

def create_task_in_folder(folder_id, space_id, task_data, access_token, cached_tasks):
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
        new_task = create_task(folder_id, space_id, task_data, responsible_ids, access_token)
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

def create_subtask_in_parent_task(parent_task_id, space_id, subtask_data, access_token, cached_tasks):
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
        new_subtask = create_subtask(parent_task_id, space_id, subtask_data, responsible_ids, access_token)
        
        # Update the cache with the newly created subtask
        if new_subtask and 'id' in new_subtask:
            cached_tasks.append(new_subtask)
            print(f"[DEBUG] Added newly created subtask '{new_subtask['title']}' with ID '{new_subtask['id']}' to cache.")
        else:
            print(f"[DEBUG] Failed to create the subtask or retrieve subtask ID.")

def create_subtask(parent_task_id, space_id, subtask_data, responsible_ids, access_token):
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
            "start": subtask_data.get("start_date").isoformat() if isinstance(subtask_data.get("start_date"), pd.Timestamp) else subtask_data.get("start_date"),
            "due": subtask_data.get("end_date").isoformat() if isinstance(subtask_data.get("end_date"), pd.Timestamp) else subtask_data.get("end_date")
        }

        # Get custom fields from API specific to the space
    custom_fields = get_custom_fields_by_space(access_token, space_id)

    # Map Excel headings to Wrike custom fields
    mapped_custom_fields = map_excel_headings_to_custom_fields(subtask_data.keys(), custom_fields)
    print(f"[DEBUG] Mapped Custom Fields: {mapped_custom_fields}")

    # Create custom fields payload
    custom_fields_payload = []
    for field_name, field_id in mapped_custom_fields.items():
        field_value = subtask_data.get(field_name) 
        print(f"[DEBUG] Retrieving '{field_name}' from task data: '{field_value}'") 
        
        if pd.notna(field_value):
            custom_fields_payload.append({
                "id": field_id,
                "value": str(field_value)  # Wrike expects the custom field values as strings
            })
    
    if custom_fields_payload:
        payload["customFields"] = custom_fields_payload

    # Debugging print statement to see the final payload
    print("Final payload being sent:", payload)   
    response = requests.post(endpoint, headers=headers, json=payload)

    if response.status_code == 200:
        subtask_data_response = response.json()
        print(f"Subtask '{subtask_data['title']}' created successfully under parent task '{parent_task_id}'")
        return subtask_data_response['data'][0] if 'data' in subtask_data_response else None
    else:
        print(f"Failed to create subtask '{subtask_data.get('title', '')}'. Status code: {response.status_code}")
        print(response.text)
        return None

# Function to read configuration from Excel
def read_config_from_excel(file_path):
    df = pd.read_excel(file_path, sheet_name='Config', header=1)
    config = df.iloc[0].to_dict()  # Convert first row to dictionary
    return config

# Function to get the Wrike space ID by name
def get_wrike_space_id(space_name, access_token):
    url = f'https://www.wrike.com/api/v4/spaces'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    spaces = response.json()['data']
    for space in spaces:
        if space['title'].lower() == space_name.lower():
            return space['id']
    return None

# Function to get the details of a space
def get_space_details(space_id, access_token):
    url = f'https://www.wrike.com/api/v4/spaces/{space_id}?fields=["members"]'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()['data'][0]

# Function to create a new Wrike space
def create_new_space(original_space, new_title, access_token):
    url = f'https://www.wrike.com/api/v4/spaces'
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    payload = {
        "title": new_title,
        "description": original_space.get("description", ""),
        "accessType": original_space.get("accessType", ""),
        "members": original_space.get("members", [])
    }
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    return response.json()['data'][0]

# Function to get custom fields in a space
def get_custom_fields(access_token):
    url = f'https://www.wrike.com/api/v4/customfields'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()['data']

# Function to create a new custom field scoped to the new space
def create_custom_field(field_data, new_space_id, access_token):
    url = f'https://www.wrike.com/api/v4/customfields'
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    payload = {
        "title": field_data.get("title"),
        "type": field_data.get("type"),  # e.g., 'Text', 'Numeric', 'DropDown'
        "settings": field_data.get("settings", {}),
        "spaceId": new_space_id  # Set the new space ID as the scope
    }

    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()

    return response.json()['data'][0]

# Function to map custom fields from the original space to the new space
def map_custom_fields(original_fields, original_space_id, new_space_id, access_token):
    field_mapping = {}
           
    # Step 1: Filter original fields by spaceId
    filtered_original_fields = [
        field for field in original_fields
        if field.get('spaceId') == original_space_id
    ]

    for field in filtered_original_fields:
        # Check if the field has a valid 'scope' and belongs to the original space
        field_scope = field.get('spaceId', [])
        
        if field_scope == original_space_id:
            print(f"Creating new custom field: {field['title']}")
            # Create the custom field in the new space
            new_field = create_custom_field(field, new_space_id, access_token)
            # Map the old field ID to the new field ID
            field_mapping[field['id']] = new_field['id']
            print(f"Mapped Custom Field: {field['title']} -> New Field ID: {new_field['id']}")
        else:
            print(f"Field {field['title']} is skipped due to missing or invalid scope.")

    return field_mapping

def map_custom_fields_propagate(original_fields, original_space_id, new_space_id, access_token):
    field_mapping = {}

    # Step 1: Retrieve all custom fields in the destination space
    destination_fields = get_custom_fields(access_token)
    dest_space_fields = [
        field for field in destination_fields
        if field.get('spaceId') == new_space_id
    ]

    for field in original_fields:
        # Check if the field belongs to the original space
        if field.get('spaceId') != original_space_id:
            continue

        # Check if the field already exists in the destination space
        existing_field = next(
            (dest_field for dest_field in dest_space_fields if dest_field['title'] == field['title']),
            None
        )

        if existing_field:
            print(f"Custom field '{field['title']}' already exists in the destination space. Mapping directly.")
            field_mapping[field['id']] = existing_field['id']
        else:
            print(f"Creating new custom field: {field['title']}")
            # Create the custom field in the destination space
            new_field = create_custom_field(field, new_space_id, access_token)
            field_mapping[field['id']] = new_field['id']
            print(f"Mapped Custom Field: {field['title']} -> New Field ID: {new_field['id']}")

    return field_mapping


# Function to get folders in a space
def get_folders_in_space(space_id, access_token):
    url = f'https://www.wrike.com/api/v4/spaces/{space_id}/folders'
    headers = {'Authorization': f'Bearer {access_token}'}
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
def get_tasks_in_folder(folder_id, access_token):
    url = f'https://www.wrike.com/api/v4/folders/{folder_id}/tasks?fields=["subTaskIds","effortAllocation","authorIds","customItemTypeId","responsibleIds","description","hasAttachments","dependencyIds","superParentIds","superTaskIds","metadata","customFields","parentIds","sharedIds","recurrent","briefDescription","attachmentCount"]'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()['data']

# Function to get detailed information about a specific task
def get_task_details(task_id, access_token):
    url = f'https://www.wrike.com/api/v4/tasks/{task_id}'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()['data'][0]

# Function to search for the parent task across all folders in the original space
def find_task_across_folders(task_id, folders, access_token):
    for folder in folders:
        tasks = get_tasks_in_folder(folder['id'], access_token)
        for task in tasks:
            if task['id'] == task_id:
                return task
    return None

def create_or_update_task(new_folder_id, task_data, task_map, access_token, folders, folder_mapping, custom_field_mapping, is_subtask=False):
    task_key = task_data['title'] + "|" + str(task_data.get('dates', {}).get('due', ''))

    # Determine if this is a subtask and handle parent task creation first
    super_task_id = None
    if 'superTaskIds' in task_data and task_data['superTaskIds']:
        parent_task_id = task_data['superTaskIds'][0]  # Assuming there's only one parent
        parent_task_key = get_task_key_by_id(parent_task_id, access_token, task_map)

        if parent_task_key not in task_map:
            # Check if the parent task already exists in another folder in the original space
            existing_parent_task = find_task_across_folders(parent_task_id, folders, access_token)
            if existing_parent_task:
                # Parent exists in the original space, so link it to the new space folder
                if existing_parent_task['id'] in folder_mapping:
                    super_task_id = folder_mapping[existing_parent_task['id']]
                else:
                    # Parent exists in the original space but needs to be created in the new space
                    new_parent_folder_id = folder_mapping.get(existing_parent_task['parentIds'][0])
                    parent_task_data = get_task_details(parent_task_id, access_token)
                    parent_task = create_or_update_task(new_parent_folder_id, parent_task_data, task_map, access_token, folders, folder_mapping)
                    super_task_id = parent_task[0]['id'] if parent_task else None
            else:
                # Parent task does not exist anywhere, create it
                parent_task_data = get_task_details(parent_task_id, access_token)
                parent_task = create_or_update_task(new_folder_id, parent_task_data, task_map, access_token, folders, folder_mapping)
                super_task_id = parent_task[0]['id'] if parent_task else None
        else:
            super_task_id = task_map[parent_task_key]
    # Map custom fields for the task
    mapped_custom_fields = []
    for field in task_data.get('customFields', []):
        field_id = field['id']
        if field_id in custom_field_mapping:
            mapped_custom_fields.append({
                'id': custom_field_mapping[field_id],
                'value': field['value']
            })
    
    # Check if the task already exists
    if task_key in task_map:
        existing_task_id = task_map[task_key]

        # Get the details of the existing task to check its current parents
        existing_task_details = get_task_details(existing_task_id, access_token)
        current_parents = existing_task_details.get('parentIds', [])
      
        # Only update if the new folder is not already a parent
        if not is_subtask and new_folder_id not in current_parents:
            url = f'https://www.wrike.com/api/v4/tasks/{existing_task_id}'
            headers = {
                'Authorization': f'Bearer {access_token}',
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
        created_task = create_tasks(
            new_folder_id=new_folder_id if not is_subtask else None,
            task_data=task_data,
            super_task_id=super_task_id,
            access_token=access_token,
            mapped_custom_fields=mapped_custom_fields  # Pass the mapped custom fields
        )

        task_map[task_key] = created_task[0]['id']

        # Handle subtask creation for the newly created task
        for sub_task_id in task_data.get('subTaskIds', []):
            subtask_data = get_task_details(sub_task_id, access_token)
            create_or_update_task(new_folder_id, subtask_data, task_map, access_token, folders, folder_mapping, custom_field_mapping, is_subtask=True)

        return created_task

# Function to create folders recursively, updating the folder_mapping with original-new folder relationships
def create_folders_recursively(paths, root_folder_id, original_space_name, new_space_name, access_token, folders, custom_field_mapping):
    folder_id_map = {}
    folder_mapping = {}
    new_paths_info = []
    task_map = {}

    for path in paths:
        folder_path = path['path']

        # Skip folder creation for the root space but handle its tasks
        if folder_path == original_space_name:
            root_tasks = get_tasks_in_folder(path['id'], access_token)
            for task in root_tasks:
                create_or_update_task(
                    new_folder_id=root_folder_id,
                    task_data=task,
                    task_map=task_map,
                    access_token=access_token,
                    folders=folders,
                    folder_mapping=folder_mapping,
                    custom_field_mapping=custom_field_mapping
                )
            continue

        # Adjust folder_path for subfolders
        if folder_path.startswith(original_space_name + '/'):
            folder_path = folder_path[len(original_space_name) + 1:]

        path_parts = folder_path.strip('/').split('/')
        parent_id = root_folder_id

        for part in path_parts:
            if part not in folder_id_map:
                # Check if the current folder is a project
                folder_data = next((f for f in folders if f.get('title') == part), None)

                project_details = folder_data.get('project') if folder_data else None

                # Create the folder or project
                new_folder_id = create_folder_or_project(
                    title=part,
                    parent_id=parent_id,
                    access_token=access_token,
                    project_details=project_details
                )

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
        folder_tasks = get_tasks_in_folder(path['id'], access_token)
        for task in folder_tasks:
            create_or_update_task(
                new_folder_id=parent_id,
                task_data=task,
                task_map=task_map,
                access_token=access_token,
                folders=folders,
                folder_mapping=folder_mapping,
                custom_field_mapping=custom_field_mapping
            )

    return new_paths_info

def get_task_key_by_id(task_id, access_token, task_map):
    task_details = get_task_details(task_id, access_token)
    task_key = task_details['title'] + "|" + str(task_details.get('dates', {}).get('due', ''))
    return task_key

def create_tasks(new_folder_id=None, task_data=None, super_task_id=None, access_token=None, mapped_custom_fields=None):
    url = f'https://www.wrike.com/api/v4/folders/{new_folder_id}/tasks' if new_folder_id else f'https://www.wrike.com/api/v4/tasks'
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
        "customFields": mapped_custom_fields or []  # Use the provided mapped custom fields
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
def create_folder(title, parent_id, access_token):
    url = f'https://www.wrike.com/api/v4/folders/{parent_id}/folders'
    payload = {'title': title, 'shareds': []}
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    return response.json()['data'][0]['id']

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
def get_folder_id_by_paths(folder_path, space_id, access_token):
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
def get_task_id_by_titles(folder_id, task_title, access_token):
    tasks = get_all_tasks_in_folder(folder_id, access_token)
    for task in tasks:
        if task['title'] == task_title:
            return task['id']
    print(f"Task with title '{task_title}' not found.")
    return None

def create_task_folder(folder_id, task_data, access_token, mapped_custom_fields=None):
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
        "customFields": mapped_custom_fields or []  
    }
    
    if dates:
        payload["dates"] = dates
    
    if effortAllocation:
        payload["effortAllocation"] = effort_allocation_payload
            
    response = requests.post(url, headers=headers, json=payload)
        
    response.raise_for_status()
    return response.json()['data']

# Function to get task details by task ID
def get_task_detail(task_id, access_token):
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

# Retry mechanism for handling rate limits
def retry_request(url, headers, retries=3, delay=60):
    for _ in range(retries):
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response
        elif response.status_code == 429:
            print("Rate limit exceeded. Sleeping for 60 seconds...")
            time.sleep(delay)
        else:
            response.raise_for_status()
    raise Exception(f"Failed after {retries} retries")

# Function to get all spaces
def get_all_spaces(access_token):
    url = 'https://www.wrike.com/api/v4/spaces'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = retry_request(url, headers=headers)
    print("Fetching all spaces...")
    
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

# Function to get all folders and subfolders in the space
def get_all_folders(space_id, access_token):
    url = f'https://www.wrike.com/api/v4/spaces/{space_id}/folders'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = retry_request(url, headers=headers)
    print("Fetching all folders and subfolders...")
    
    try:
        return response.json()
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON response: {e}")
        print(f"Response content: {response.content}")
        raise

# Function to get task details by ID with custom status mapping
def get_tasks_details(task_id, access_token, custom_status_mapping, custom_field_mapping):
    url = f'https://www.wrike.com/api/v4/tasks/{task_id}?fields=["effortAllocation"]'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = retry_request(url, headers=headers)
    print(f"Fetching details for task {task_id}")
    
    try:
        task_data = response.json()["data"][0]
        custom_status_id = task_data.get("customStatusId")
        task_data["customStatus"] = custom_status_mapping.get(custom_status_id, "Unknown")
        # Process custom fields by mapping ID to name
        custom_fields = task_data.get("customFields", [])
        custom_field_data = {custom_field_mapping.get(cf["id"], "Unknown Field"): cf.get("value", "") for cf in custom_fields}
        task_data["customFields"] = custom_field_data

        return task_data
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON response: {e}")
        print(f"Response content: {response.content}")
        raise

# Function to get tasks for a folder
def get_tasks_for_folder(folder_id, access_token):
    url = f'https://www.wrike.com/api/v4/folders/{folder_id}/tasks?fields=["subTaskIds","effortAllocation","authorIds","customItemTypeId","responsibleIds","description","hasAttachments","dependencyIds","superParentIds","superTaskIds","metadata","customFields","parentIds","sharedIds","recurrent","briefDescription","attachmentCount"]'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = retry_request(url, headers=headers)
    print(f"Fetching tasks for folder {folder_id}")
    
    try:
        return response.json()["data"]
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON response: {e}")
        print(f"Response content: {response.content}")
        raise

# Recursive function to get all subtask IDs
def get_all_subtask_ids(task, token):
    task_ids = [{"id": task["id"], "title": task["title"]}]
    if "subTaskIds" in task:
        for subtask_id in task["subTaskIds"]:
            subtask = get_task_details(subtask_id, token, {})
            task_ids.extend(get_all_subtask_ids(subtask, token))
    return task_ids

# Function to clean HTML content and preserve line breaks
def clean_html(raw_html):
    soup = BeautifulSoup(raw_html, "html.parser")
    lines = soup.stripped_strings
    return "\n".join(lines)

# Function to get user details by ID
def get_user_details(user_id, access_token, user_cache):
    if user_id in user_cache:
        return user_cache[user_id]
    
    url = f"https://www.wrike.com/api/v4/users/{user_id}"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = retry_request(url, headers=headers)
    print(f"Fetching details for user {user_id}")
    
    try:
        user_data = response.json()["data"][0]
        email = user_data["profiles"][0]["email"]
        user_cache[user_id] = email
        return email
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON response: {e}")
        print(f"Response content: {response.content}")
        raise

# Function to get custom statuses
def get_custom_statuses(access_token):
    url = 'https://www.wrike.com/api/v4/workflows'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = retry_request(url, headers=headers)
    print("Fetching custom statuses...")
    
    try:
        return response.json()["data"]
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON response: {e}")
        print(f"Response content: {response.content}")
        raise

# Function to create a mapping from customStatusId to custom status name
def create_custom_status_mapping(workflows):
    custom_status_mapping = {}
    for workflow in workflows:
        for status in workflow.get("customStatuses", []):
            custom_status_mapping[status["id"]] = status["name"]
    return custom_status_mapping

# Create a mapping from customFieldId to customFieldName and customFieldType
def create_custom_field_mapping(custom_fields):
    custom_field_mapping = {}
    for field in custom_fields:
        field_title = field["title"]
        field_type = field["type"]
        # Store both name and type in the mapping
        custom_field_mapping[field["id"]] = f"{field_title} [{field_type}]"
    return custom_field_mapping

# Function to save data to JSON
def save_to_json(data, filename='workspace_data.json'):
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

# Function to create a folder by its path within a specific space
def create_folder_by_path(folder_path, space_id, access_token):
    folder_names = folder_path.split('\\')
    parent_folder_id = get_folder_id_by_name(space_id, folder_names[0], access_token)

    # If the parent folder does not exist, create it
    if not parent_folder_id:
        parent_folder_id = create_folders(space_id, folder_names[0], access_token)
        if not parent_folder_id:
            print(f"Failed to create the parent folder '{folder_names[0]}' in space '{space_id}'")
            return None

    # Iterate over the subfolders and create them if necessary
    for folder_name in folder_names[1:]:
        parent_folder_id = get_or_create_subfolder(parent_folder_id, folder_name, access_token)
        if not parent_folder_id:
            print(f"Failed to create or find subfolder '{folder_name}' in space '{space_id}'")
            return None

    return parent_folder_id

# Helper function to create a folder in a space
def create_folders(space_id, folder_name, access_token):
    url = f"https://www.wrike.com/api/v4/folders"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    data = {
        "title": folder_name,
        "spaceId": space_id
    }
    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 201:
        return response.json().get("data", {}).get("id")
    else:
        print(f"Error creating folder '{folder_name}': {response.text}")
        return None

# Function to create a folder or project
def create_folder_or_project(title, parent_id, access_token, project_details=None):
    url = f'https://www.wrike.com/api/v4/folders/{parent_id}/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    # Common payload for folders
    payload = {
        'title': title,
        'shareds': []
    }
    
    # Add project-specific details if it is a project
    if project_details:
        payload['project'] = {
            'authorId': project_details.get('authorId'),
            'ownerIds': project_details.get('ownerIds', []),
            'customStatusId': project_details.get('customStatusId'),
            'createdDate': project_details.get('createdDate')
        }

    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    
    return response.json()['data'][0]['id']

