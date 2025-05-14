import os
import Const as C
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askdirectory

#def check_and_update_excel_dir_path():
    #excel_dir_path = check_excel_dir_path()
    #if not isinstance(excel_dir_path,str):
        #if excel_dir_path == 0:
        #else:

def check_excel_dir_path():
    """
    Search for the Excel path file and get the Excel file.

    Returns:
    int: 
        excel_path Config file OK, path OK and accessible,
        0 Config file OK, path detected but not OK
        -1 Config file OK but path not detected
        -2 No config file.
    """
    # Build the absolute path to the configuration file
    current_dir = os.path.dirname(__file__)
    config_file = os.path.abspath(os.path.join(current_dir, '..', C.NAME_FOLDER_CONFIG, C.NAME_FILE_PATH_EXCEL))

    # Ensure that the config file exists, create it if necessary
    if not ensure_config_file_exists(config_file):
        return -2  # If the config file creation fails, return -2

    if not os.path.isdir(os.path.dirname(config_file)):
        print(f"Directory '{os.path.dirname(config_file)}' does not exist.")
        return -2  # Directory where config file is located doesn't exist.

    # Try to retrieve the Excel file path from the config file
    excel_path = get_first_path_after_marker(config_file)

    if isinstance(excel_path,str):
        print(f"Valid Excel file path found: {excel_path}")

        if os.path.isdir(excel_path):  # If dir exists
            print(f"Path '{excel_path}' is accessible.")
            return excel_path  # Folder exists, and the path is accessible.

        else:  # Path exists but is not accessible (non-existent dir or invalid path)
            print(f"Path '{excel_path}' exists but is not accessible.")
            return 0  # Path exists but is not accessible
    else:
        print("No valid Excel path found in the config file.")
        return excel_path  # No path found in config file.

def ensure_config_file_exists(config_file):
    """
    Ensure that the config file exists. If it doesn't, create it with default content.
    Also ensures that the config folder exists before creating the file.

    Args:
    - config_file (str): The path to the configuration file to check.

    Returns:
    - bool: True if the file exists or was successfully created, False if creation failed.
    """
    try:
        # Get the directory of the config file
        config_dir = os.path.dirname(config_file)

        # Check if the directory exists; if not, create it
        if not os.path.isdir(config_dir):
            print(f"Directory '{config_dir}' does not exist. Creating it...")
            os.makedirs(config_dir)  # Create the directory if it doesn't exist
            print(f"Directory '{config_dir}' has been created.")

        # Check if the file exists
        if not os.path.isfile(config_file):
            print(f"Config file '{config_file}' does not exist. Creating it...")
            # Create the config file with default content
            with open(config_file, 'w') as f:
                f.write(C.EXCEL_PATH_BALISE + '\n')
            print(f"Config file '{config_file}' has been created.")
            return True
        else:
            print(f"Config file '{config_file}' already exists.")
            return True
    except Exception as e:
        print(f"Error creating config file '{config_file}': {e}")
        return False

def get_first_path_after_marker(file_path):
    """
    Read a .txt file, find the first occurrence of a line containing only '>',
    then check if the next line is a valid file path.

    Parameters:
    file_path (str): Path to the input .txt file.

    Returns:
    str or 
    0: The first valid file path found after the marker not valid
    -1 if not found.
    """
    try:
        with open(file_path, 'r') as file:
            lines = file.readlines()

        for i, line in enumerate(lines):
            if line.strip() == C.EXCEL_PATH_BALISE:
                if i + 1 < len(lines):
                    potential_path = lines[i + 1].strip()
                    if os.path.isdir(potential_path):
                        return potential_path
                    else:
                        print(f"Path found after '>' is not a valid dir: {potential_path}")
                        return 0
        print("No marker '>' followed by a valid path was found.")
        return -1

    except FileNotFoundError:
        print(f"Input dir not found: {file_path}")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def update_excel_path(new_path):
    """
    Update the Excel path in the configuration file. 
    This function will replace the old path (if present) and write the new path 
    after the marker '>'. If the marker is not found, it will add the marker and path at the end.

    Args:
    - new_path (str): The new path to write in the file.

    Returns:
    1: path updated
    -1: provided path invalid or inaccessible
    -2: Config File KO
    0: unexpected error
    """
    try:
        # Check if the new path is valid
        if not os.path.isdir(new_path):
            print(f"Provided path '{new_path}' is not valid or accessible.")
            return -1  # Invalid path

        # Get the directory of the config file
        current_dir = os.path.dirname(__file__)
        config_file = os.path.abspath(os.path.join(current_dir, '..', C.NAME_FOLDER_CONFIG, C.NAME_FILE_PATH_EXCEL))

        # Ensure the directory exists; create it if necessary
        if not ensure_config_file_exists(config_file):
            return -2

        # Open the file and read the content
        with open(config_file, 'r') as f:
            lines = f.readlines()

        # Find the marker '>' and replace the next line with the new path
        path_updated = 0
        for i, line in enumerate(lines):
            if line.strip() == C.EXCEL_PATH_BALISE:
                # Replace the next line after the marker with the new path
                if i + 1 < len(lines):  # If there is a next line after the marker
                    lines[i + 1] = new_path + '\n'  # Replace the path
                    path_updated = 1
                    print(f"Updated path to '{new_path}' in the config file.")
                else:
                    # If there is no next line, add the new path after the marker
                    lines.append(new_path + '\n')  # Append the new path to the file
                    path_updated = 1
                    print(f"Added new path '{new_path}' after the marker '>' in the config file.")
                break
        # If the marker '>' was not found, add it at the end of the file
        if not path_updated:
            print("Marker '>' not found. Adding marker and new path.")
            lines.append(C.EXCEL_PATH_BALISE + '\n')  # Add the marker at the end
            lines.append(new_path + '\n')  # Add the new path after the marker
            path_updated = 1
            print(f"Added marker '>' and path '{new_path}' to the config file.")

        # Write the updated content back to the config file
        with open(config_file, 'w') as f:
            f.writelines(lines)

        return path_updated  # Path successfully updated

    except Exception as e:
        print(f"Error updating the path in the config file '{config_file}': {e}")
        return 0

def update_excel_path_popup_ask():
    """
    Display a popup message in French asking the user to update the Excel path. 
    Offers the option to browse for a directory (OK) or cancel the action (Cancel).
    
    Returns:
    1: update successfully
    0: discard the update
    -1: error
    """
    try:
        # Display the popup message with 'OK' and 'Cancel' buttons
        response = messagebox.askquestion(
            "Mise à jour du chemin Excel",  # Title of the popup
            "Le chemin Excel doit être mis à jour.\nSouhaitez-vous rechercher un dossier ?",  # Message
            icon='info'
        )

        if response == 'yes':  # If user clicked 'OK'
            # Open the file dialog to choose a directory
            folder_path = askdirectory(title="Choisissez le dossier Excel")
            if folder_path:  # If a directory is selected
                update_excel_path(folder_path)
                print(f"Le chemin Excel a été mis à jour avec le dossier : {folder_path}")
                return 1  # Return the selected folder path
            else:
                print("Aucun dossier sélectionné.")
                return 0  # Return None if no folder was selected
        else:
            print("Action annulée.")
            return 0  # If user clicked 'Cancel', return None
    
    except Exception as e:
        # Handle any exceptions and return an error code
        print(f"Erreur lors de la mise à jour du chemin Excel : {e}")
        return -1  # Return error code
    
def update_excel_path_popup_mandatory():
    """
    Display a popup message in French asking the user to update the Excel path. 
    Offers the option to browse for a directory (OK) or cancel the action (Cancel).
    
    Returns:
    1: update successfully
    0: discard the update
    -1: error
    """
    try:
        # Display the popup message with 'OK' and 'Cancel' buttons
        response = messagebox.showerror(
            "Mise à jour du chemin Excel",  # Title of the popup
            "Le chemin Excel doit être mis à jour.",  # Message
            icon='warning'
        )
        # Open the file dialog to choose a directory
        folder_path = askdirectory(title="Choisissez le dossier Excel")
        if folder_path:  # If a directory is selected
            update_excel_path(folder_path)
            print(f"Le chemin Excel a été mis à jour avec le dossier : {folder_path}")
            return 1  # Return the selected folder path
        else:
            print("Aucun dossier sélectionné.")
            return 0  # Return None if no folder was selected

    except Exception as e:
        # Handle any exceptions and return an error code
        print(f"Erreur lors de la mise à jour du chemin Excel : {e}")
        return -1  # Return error code
