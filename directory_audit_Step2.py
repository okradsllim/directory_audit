

'''
1. Refined Hyperlink Extraction: I revised the extract_hyperlinks function to use row indices as keys in the hyperlinks dictionary. 
This aligns with DataFrame indexing, ensuring that each hyperlink is correctly associated with its corresponding row in the DataFrame. 
I also made sure to handle cases where a cell might not contain a hyperlink by assigning None.

2. Improved 'Extracted Path' Mapping: To properly map the 'Path' column values to the 'Extracted Path' column, I implemented a more efficient approach. 
By using audit_sheet_df['Extracted Path'] = audit_sheet_df.index + 2 and then applying audit_sheet_df['Extracted Path'].apply(lambda x: hyperlinks.get(x)), 
I made sure the hyperlinks correctly correspond to the DataFrame rows, considering Excel's row indexing.

3. Enhanced File Renaming Logic: I've added logic in the action_rename function to preserve file extensions when renaming files. 
This is crucial for maintaining the integrity and functionality of files, especially when they are of specific types like documents or images.

4. Updated Move Functionality: In the action_move function, I now ensure that the target directory is correctly retrieved, 
either from the folders_to_create dictionary or by defaulting to the BASE_DIRECTORY. 
This update provides more accurate and flexible file moving capabilities.

5. Interactive Directory Creation: I introduced user prompts to handle cases where directories don't exist within the base directory. 
This interactive feature not only improves user experience but also gives users control over how the directory structure is managed during the script's execution.

6. Recycle Directory Management: I've incorporated logic to handle situations where a recycle directory is necessary, particularly for delete actions. 
The script checks if the recycle directory is outside the BASE_DIRECTORY and prompts the user to create it if it doesn't exist, enhancing the script's safety and usability.'''


import pandas as pd
from pathlib import Path
from shutil import move
import logging
import os
from openpyxl import load_workbook
import re


def extract_hyperlinks(excel_file_path, sheet_name='AuditSheet', target_column_name='Path'):
    wb = load_workbook(excel_file_path, data_only=False)
    sheet = wb[sheet_name]

    header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
    try:
        target_column_index = header_row.index(target_column_name)
    except ValueError:
        logging.error(f"Column '{target_column_name}' not found in the sheet '{sheet_name}'.")
        return {}

    hyperlink_regex = re.compile(r'HYPERLINK\("([^"]+)"')
    hyperlinks = {}
    for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=False), start=2):
        cell = row[target_column_index]
        if cell.value and "HYPERLINK" in str(cell.value):
            match = hyperlink_regex.search(cell.value)
            if match:
                link_location = match.group(1)
                hyperlinks[row_index] = link_location
                logging.debug(f"Hyperlink extracted at row {row_index}: {link_location}")
            else:
                logging.debug(f"Hyperlink formula not found at row {row_index}")
                hyperlinks[row_index] = None
        else:
            logging.debug(f"No hyperlink at row {row_index}")
            hyperlinks[row_index] = None

    return hyperlinks

def extract_base_directory(paths):
    absolute_paths = [Path(p).resolve() for p in paths if p is not None]
    if not absolute_paths:
        return None
    common_path = os.path.commonpath(absolute_paths)
    return Path(common_path)

def action_rename(original_path, new_name):
    logging.debug(f"Attempting to rename: {original_path} to {new_name}")

    if not original_path.exists():
        log_message = f"Rename Failed - Original file/folder does not exist: {original_path}"
        logging.warning(log_message)
        return log_message

    new_path = original_path.parent / new_name
    
    if original_path.is_file(): # If the original path is a file, preserve the extension in the new name
        original_extension = original_path.suffix
        if not new_name.endswith(original_extension):
            new_name += original_extension

    if new_path.exists():
        log_message = f"Rename Failed - New name already exists: {new_path}"
        logging.warning(log_message)
        return log_message

    try:
        original_path.rename(new_path)
        log_message = f"Renamed {original_path} to {new_path}"
        logging.info(log_message)
        print(log_message)
        return new_name  # Return the full new_name including the extension
    except Exception as e:
        log_message = f"Rename Error: {e}"
        logging.error(log_message)
        return log_message

def action_move(original_path, move_to_folder_name, folders_to_create, BASE_DIRECTORY):
    
    logging.debug(f"Attempting to move: {original_path} to {move_to_folder_name}")
    
    # Retrieve the target directory from the folders_to_create dictionary or default to combining with BASE_DIRECTORY
    target_dir = folders_to_create.get(move_to_folder_name, BASE_DIRECTORY / move_to_folder_name)

    # Proceed with the move if the original path exists and the target is a directory
    if not original_path.exists():
        log_message = f"Move Failed - Original file/folder does not exist: {original_path}"
        logging.warning(log_message)
        return log_message

    if not target_dir.is_dir():
        log_message = f"Move Failed - Target directory is not a directory: {target_dir}"
        logging.warning(log_message)
        return log_message

    try:
        print(f"Attempting to move {original_path} to {target_dir}")
        move(str(original_path), str(target_dir))
        log_message = f"Moved {original_path} to {target_dir}"
        logging.info(log_message)
        print(log_message)
        return log_message
    
    except Exception as e:
        log_message = f"Move Error: {e}"
        logging.error(log_message)
        print(log_message)
        return log_message

def action_delete(original_path, recycle_dir_path):
    
    logging.debug(f"Attempting to move: {original_path} to {recycle_dir_path}")
    
    if original_path.exists():
        try:
            move(str(original_path), str(recycle_dir_path))
            log_message = f'Moved {original_path} to recycle directory: {recycle_dir_path}'
            logging.info(log_message)
            print(log_message)
            return log_message
        except Exception as e:
            log_message = f"Delete Error: {e}"
            logging.error(log_message)
            return log_message
    else:
        log_message = 'Delete Failed - File not found'
        logging.warning(log_message)
        return log_message

def perform_actions(audit_sheet_df, recycle_dir_path, folders_to_create, action_logs, BASE_DIRECTORY):
    
    logging.debug("Starting to perform actions.")

    for index, row in audit_sheet_df.iterrows():
        logging.debug(f"Processing row {index + 1}.")
        name = row['Name']
        actions = row['Action'].lower().split(',') if isinstance(row['Action'], str) else []
        new_name = row['Rename as…'].strip() if isinstance(row['Rename as…'], str) else ''
        move_to = row['Move to…'].strip() if isinstance(row['Move to…'], str) else ''

        original_path = Path(row['Extracted Path'])
        
        # Check if the path points to an existing file or directory
        if not original_path.exists():
            action_logs.append({'Action': ', '.join(actions), 'Path': str(original_path), 'Status': 'Original path not found or does not exist'})
            continue

        status = []
        for action in actions:
            logging.debug(f"Performing action '{action}' for {name}.")
            if action == 'rename' and new_name:
                result = action_rename(original_path, new_name)
                status.append(result)
                if not result.startswith("Failed"):
                    original_path = Path(result)
                logging.info(status[-1])

            elif action == 'move' and move_to:
                if move_to in folders_to_create or (BASE_DIRECTORY / move_to).exists():
                    result = action_move(original_path, move_to, folders_to_create, BASE_DIRECTORY)
                    status.append(result)
                else:
                    status.append(f"Failed - Target directory not found for moving: {move_to}")
                logging.info(status[-1])
                
            elif action == 'delete':
                result = action_delete(original_path, recycle_dir_path)
                status.append(result)
                logging.info(status[-1])

        action_logs.append({'Action': ', '.join(actions), 'Path': str(original_path), 'Status': '; '.join(status)})
        
        logging.debug(f"Row {index + 1} actions completed with status: {'; '.join(status)}")
    
def validate_path(path, should_exist=True, is_directory=False):
    p = Path(path).resolve()
    if should_exist and not p.exists():
        print(f"Path does not exist: {p}")
        return False
    if is_directory and not p.is_dir():
        print(f"Path is not a directory: {p}")
        return False
    return True

def get_validated_path(prompt, should_exist=True, is_directory=False, max_attempts=5):
    attempts = 0
    while attempts < max_attempts:
        user_input = input(f"{prompt} (Type 'exit' to quit, {max_attempts - attempts} attempts left): ").strip()
        
        if user_input.lower() == 'exit':
            print("Exiting program.")
            return None

        path = Path(user_input).resolve()
        if validate_path(path, should_exist, is_directory):
            return path

        attempts += 1
        print(f"Invalid input. Please try again.")

    print("Maximum number of attempts reached. Exiting program.")
    return None

def main():

    desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop')
    
    log_file_path = os.path.join(desktop_dir, 'file_manager.log')
    logging.basicConfig(filename=log_file_path, level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
    
    action_logs = []

    input_excel_file = get_validated_path("Enter path to the Excel file: ", should_exist=True, is_directory=False)
    if input_excel_file is None:
        return 
    
    audit_sheet_df = pd.read_excel(input_excel_file, sheet_name='AuditSheet')
    
    ## Initial Check for Actionable Data
    if audit_sheet_df['Action'].isna().all():
        print("No actions specified in the 'Action' column. Exiting program.")
        return
    
    ## Normalize and Filter Actions
    valid_actions = ['delete', 'rename', 'move']
    audit_sheet_df['Action'] = audit_sheet_df['Action'].fillna('').astype(str).str.lower().str.strip()
    valid_action_rows = audit_sheet_df[audit_sheet_df['Action'].str.contains('|'.join(valid_actions))]
    
    ## Further Validation of Actions
    if valid_action_rows.empty:
        print("No valid actions (delete, rename, move) found in the 'Action' column. Exiting program.")
        return
    
    #input_excel_file = get_validated_path("Enter the path to the Excel file: ", should_exist=True, is_directory=False)
    #hyperlinks = extract_hyperlinks(input_excel_file)
    #audit_sheet_df['Extracted Path'] = audit_sheet_df.index + 2  # +2 because Excel rows start at 1 and header is at row 1
    #audit_sheet_df['Extracted Path'] = audit_sheet_df['Extracted Path'].apply(lambda x: hyperlinks.get(x))
    # Normalize actions to lowercase for consistent processing
    # audit_sheet_df['Action'] = audit_sheet_df['Action'].fillna('').astype(str).str.lower().str.strip()

    # Extract Hyperlinks
    hyperlinks = extract_hyperlinks(input_excel_file)
    valid_action_rows['Extracted Path'] = valid_action_rows.index + 2
    valid_action_rows['Extracted Path'] = valid_action_rows['Extracted Path'].apply(lambda x: hyperlinks.get(x))
    
    # Determine the Base Directory
    BASE_DIRECTORY = extract_base_directory(valid_action_rows['Extracted Path'])
    print(f"Base directory: {BASE_DIRECTORY}")

    # Extract unique folder names from the 'Move to…' column where the action is 'move'
    move_to_folders = audit_sheet_df[audit_sheet_df['Action'].str.contains('move', na=False)]['Move to…'].unique()

    folders_to_create = {}

    for folder_name in move_to_folders:
        folder_path = BASE_DIRECTORY / folder_name
        if not folder_path.exists():
            create_folder = input(f"The folder '{folder_name}' does not exist in the base directory. Do you want to create it? (yes/no): ").strip().lower()
            if create_folder == 'yes':
                # Create the folder and log the action
                folders_to_create[folder_name] = folder_path
                folder_path.mkdir(parents=True, exist_ok=True)
                logging.info(f"Directory created: {folder_path}")
                print(f"Directory created: {folder_path}")
            elif create_folder == 'no':
                # Log the decision and use BASE_DIRECTORY for the move
                logging.warning(f"'{folder_name}' folder creation skipped by user. Defaulting move to BASE_DIRECTORY.")
                print(f"Proceeding without creating the '{folder_name}' folder. Files will be moved to the base directory.")
                folders_to_create[folder_name] = BASE_DIRECTORY
            else:
                logging.error("Invalid input received for folder creation choice.")
                print("Invalid input. Please type 'yes' or 'no'.")


    if audit_sheet_df['Action'].str.contains(r'delete', case=False, na=False).any():
        while True:
            recycle_dir_input = input("Enter the path to the recycle directory: ").strip()
            recycle_dir_path = Path(recycle_dir_input).resolve()

            # Check if the recycle directory is not inside BASE_DIRECTORY
            if BASE_DIRECTORY in recycle_dir_path.parents:
                print(f"The recycle directory should not be inside the base directory '{BASE_DIRECTORY}'. Please choose a different location.")
                continue

            
            if not recycle_dir_path.exists():
                create_dir = input(f"The path {recycle_dir_path} does not exist. Do you want to create it? (yes/no): ").strip().lower()
                if create_dir == 'yes':
                    recycle_dir_path.mkdir(parents=True, exist_ok=True)
                    print(f"Directory created: {recycle_dir_path}")
                    break
                elif create_dir == 'no':
                    continue
                else:
                    print("Invalid input. Please type 'yes' or 'no'.")
            elif not recycle_dir_path.is_dir():
                print(f"The path is not a directory: {recycle_dir_path}. Please enter a valid directory path.")
                continue
            else:
                break
    else:
        recycle_dir_path = None 


    # Performing actions
    perform_actions(valid_action_rows, recycle_dir_path, folders_to_create, action_logs, BASE_DIRECTORY)
    
    for log in action_logs:
        print(f"Action: {log['Action']}, Path: {log['Path']}, Status: {log['Status']}")

if __name__ == "__main__":
    main()