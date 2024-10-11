import pandas as pd
import requests
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
    
    auth_info = {
        'redirect_uri': redirect_url
    }
    
    access_token = wrike.authenticate(auth_info=auth_info)
    print(f"New access token obtained: {access_token}")
    return access_token

# Function to get all contacts
def get_contacts(access_token):
    endpoint = 'https://www.wrike.com/api/v4/contacts'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(endpoint, headers=headers)

    if response.status_code == 200:
        contacts_data = response.json()  # Parse the JSON response
        print(f"All contacts: {contacts_data}")
        return contacts_data
    else:
        print(f"Failed to get contacts. Status code: {response.status_code}")
        return None

# Function to get specific contact by name
def get_specific_contact(access_token, contact_name):
    # First, get all contacts
    contacts_data = get_contacts(access_token)
    
    if contacts_data:
        contacts_list = contacts_data.get('data', [])
        # Search for the contact by name
        for contact in contacts_list:
            if contact.get('firstName') == contact_name or contact.get('lastName') == contact_name:
                contact_id = contact.get('id')
                print(f"Found contact {contact_name} with ID: {contact_id}")
                
                # Now fetch details for this specific contact using the contact ID
                get_contact_by_id(access_token, contact_id)
                return
            
        print(f"No contact found with the name: {contact_name}")
    else:
        print("Unable to retrieve contacts.")

# Function to get a contact by their contact ID
def get_contact_by_id(access_token, contact_id):
    endpoint = f'https://www.wrike.com/api/v4/contacts/{contact_id}'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(endpoint, headers=headers)

    if response.status_code == 200:
        contact_data = response.json()  # Parse the JSON response
        print(f"Details for contact ID {contact_id}: {contact_data}")
    else:
        print(f"Failed to get contact by ID. Status code: {response.status_code}")

    
# Function to get tasks in current account
def get_tasks(access_token):
    endpoint = f'https://www.wrike.com/api/v4/tasks'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(endpoint, headers=headers)

    if response.status_code == 200:
        task_data = response.json()  # Parse the JSON response
        print(f"Tasks in current account: {task_data}")
    else:
        print(f"Failed to get all tasks in current account. Status code: {response.status_code}")

# Function to get all spaces and retrieve space ID based on the space name
def get_space_id_by_name(access_token, space_name):
    endpoint = 'https://www.wrike.com/api/v4/spaces'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    response = requests.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        spaces_data = response.json()  # Parse the JSON response
        spaces_list = spaces_data.get('data', [])
        
        # Search for the space by name
        for space in spaces_list:
            if space.get('title') == space_name:
                space_id = space.get('id')
                print(f"Found space '{space_name}' with ID: {space_id}")
                return space_id
        
        print(f"No space found with the name: {space_name}")
        return None
    else:
        print(f"Failed to retrieve spaces. Status code: {response.status_code}")
        return None

# Function to get all tasks within a space using space_id
def get_tasks_in_space(access_token, space_id):
    endpoint = f'https://www.wrike.com/api/v4/spaces/{space_id}/tasks'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    response = requests.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        tasks_data = response.json()  # Parse the JSON response
        print(f"Tasks in space with ID {space_id}: {tasks_data}")
    else:
        print(f"Failed to get tasks in space with ID {space_id}. Status code: {response.status_code}")

# Main function to get tasks by space name
def get_tasks_by_space_name(access_token):
    # Prompt user for space name
    space_name = input("Enter the space name: ")
    
    # Retrieve space ID by name
    space_id = get_space_id_by_name(access_token, space_name)
    
    if space_id:
        # Retrieve tasks within the space
        get_tasks_in_space(access_token, space_id)
    else:
        print(f"Cannot retrieve tasks, space with name '{space_name}' not found.")

# Function to get all comments in current account
def get_comments(access_token):
    endpoint = f'https://www.wrike.com/api/v4/comments'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    response = requests.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        comments_data = response.json()  # Parse the JSON response
        print(f"Comments in current account: {comments_data}")
    else:
        print(f"Failed to get comments in current account: Status code: {response.status_code}")

# Function to get folders in current account
def get_folders(access_token):
    endpoint = f'https://www.wrike.com/api/v4/tasks'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(endpoint, headers=headers)

    if response.status_code == 200:
        folders_data = response.json()  # Parse the JSON response
        print(f"Folders in current account: {folders_data}")
    else:
        print(f"Failed to get all folders in current account. Status code: {response.status_code}")

# Function to get all folders within a space using space_id
def get_folders_in_space(access_token, space_id):
    endpoint = f'https://www.wrike.com/api/v4/spaces/{space_id}/folders'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    response = requests.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        folders_data = response.json()  # Parse the JSON response
        print(f"Folders in space with ID {space_id}: {folders_data}")
    else:
        print(f"Failed to get folders in space with ID {space_id}. Status code: {response.status_code}")

# Main function to get folders by space name
def get_folders_by_space_name(access_token):
    # Prompt user for space name
    space_name = input("Enter the space name: ")
    
    # Retrieve space ID by name
    space_id = get_space_id_by_name(access_token, space_name)
    
    if space_id:
        # Retrieve tasks within the space
        get_folders_in_space(access_token, space_id)
    else:
        print(f"Cannot retrieve folders, space with name '{space_name}' not found.")

# Function to display the menu and allow the user to choose which function to call
def display_menu():
    print("\nSelect an option:")
    print("1. Fetch All Tasks")
    print("2. Fetch Tasks by Space Name")
    print("3. Fetch All Contacts")
    print("4. Fetch Specific Contact by Name")
    print("5. Fetch All Comments")
    print("6. Fetch All Folders")
    print("7. Fetch Folders by Space Name")
    print("8. Exit")
    choice = input("Enter the number of your choice: ")
    return choice

# Main function to authenticate and proceed with further actions
def api_calls_main():
    excel_file = input("Enter the path to the Excel file: ")
    try:
        config_df = pd.read_excel(excel_file, sheet_name="Config", header=1)        
    except Exception as e:
        print(f"Failed to read Excel file: {e}")
        return 

    access_token = config_df.at[0, "Token"]
    
        # Validate the token
    if not validate_token(access_token):
        # If the token is invalid, authenticate using OAuth2 and update the access_token
        wrike = OAuth2Gateway1(excel_filepath=excel_file)
        access_token = wrike._create_auth_info()  # Perform OAuth2 authentication
        print(f"New access token obtained: {access_token}")

    # Menu to allow the user to choose an action
    while True:
        choice = display_menu()
        if choice == "1":
            get_tasks(access_token)
        elif choice == "2":
            get_tasks_by_space_name(access_token)
        elif choice == "3":
            get_contacts(access_token)
        elif choice == "4":
            contact_name = input("Enter the contact's first or last name: ")
            get_specific_contact(access_token, contact_name)
        elif choice == "5":
            get_comments(access_token)
        elif choice == "6":
            get_folders(access_token)
        elif choice == "7":
            get_folders_by_space_name(access_token)
        elif choice == "8":
            print("Exiting...")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__api_calls_main__":
    api_calls_main()
