import requests
import pandas as pd

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

# Function to update a task with new parent IDs
def update_task_with_parent_id(task_id, parent_ids, access_token):
    endpoint = f'https://www.wrike.com/api/v4/tasks/{task_id}'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    data = {
        'addParents': parent_ids
    }

    response = requests.put(endpoint, headers=headers, json=data)
    if response.status_code == 200:
        print(f"Task '{task_id}' updated successfully with new parent IDs: {parent_ids}")
    else:
        print(f"Failed to update task '{task_id}'. Status code: {response.status_code}")
        print(response.text)


# Main function to handle reading from Excel and updating tasks
def main():
    # Ask the user to input the path of the Excel file
    excel_file = input("Please enter the full path to the Excel file: ")

    # Load the access token from the 'Config' sheet
    config_df = pd.read_excel(excel_file, sheet_name='Config', header=1)
    access_token = config_df.at[0, 'Token']

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

        # Extract the latest subfolder from the source path
        source_subfolder_name = source_path.split('\\')[-1]

        # Get the destination folder ID
        destination_folder_id = get_folder_id_by_path(destination_path, space_id, access_token)
        if not destination_folder_id:
            # Create the destination folder if it doesn't exist
            destination_folder_name = destination_path.split('\\')[-1]  # Get the last part of the path as folder name
            parent_folder_id = get_folder_id_by_path('\\'.join(destination_path.split('\\')[:-1]), space_id, access_token)  # Get the parent folder's path

            if not parent_folder_id:
                print(f"Parent folder for '{destination_folder_name}' not found. Skipping.")
                continue

            destination_folder_id = create_folder_in_space(destination_folder_name, parent_folder_id, access_token)

        # If the task title is empty, replicate the folder structure
        if pd.isna(task_title):
            # Create the subfolder under the destination path
            new_subfolder_id = create_folder_in_space(source_subfolder_name, destination_folder_id, access_token)

            if not new_subfolder_id:
                print(f"Failed to create subfolder '{source_subfolder_name}' in destination '{destination_path}'.")
                continue

            # Copy all tasks from the source subfolder to the newly created subfolder
            tasks = get_all_tasks_in_folder(source_folder_id, access_token)
            for task in tasks:
                update_task_with_parent_id(task['id'], [new_subfolder_id], access_token)
                print(f"Task '{task['id']}' updated successfully with new parent ID: '{new_subfolder_id}'")

        else:
            # Update the specified task
            task_id = get_task_id_by_title(source_folder_id, task_title, access_token)
            if task_id:
                update_task_with_parent_id(task_id, [destination_folder_id], access_token)
                print(f"Task '{task_id}' updated successfully with new parent ID: '{destination_folder_id}'")

# Call the main function
if __name__ == '__main__':
    main()