
# Option 1 audit and exclusion works. Takes a while on the network drive.
# Option 1 excluding two folders also worked with my one of my local folders
# It successfully creates two worksheets in one workbook: RawData and AuditSheet
# RawData has File Name	File Path;	File Type;	File Size (MB);	Date Created;	Last Modified;	Owner
# AuditSheet has Name	Item Type;	Owner;	Date Created;	Last Modified;	Path;	Consult staff?;	Proposed retention;	Justification;	Move to…;	Notes


import os
import logging
import pandas as pd
import re
import datetime
import platform
from pathlib import Path

desktop_path = os.path.join(os.path.expanduser("~"), 'Desktop')

# Configure logging to capture errors
logging.basicConfig(filename='error.log', level=logging.ERROR)

def audit_directory_process(desktop_path):
    directory_path = input("Enter the directory path to audit: ").strip()
    print(f"Path to audit: '{os.path.abspath(directory_path)}'")
    if os.path.isdir(directory_path):
        exclude_folders = get_exclusion_list(directory_path)

        valid_file_name = False
        while not valid_file_name:
            output_file_name = input("Enter the desired output file name (without extension): ").strip()
            if output_file_name and re.match("^[a-zA-Z0-9_-]*$", output_file_name):
                output_file = os.path.join(desktop_path, f'{output_file_name}.xlsx')
                valid_file_name = True
            else:
                print("Invalid file name. Please use only letters, numbers, hyphens, and underscores.")
        
        file_data = list_files(directory_path, exclude_folders)
        df_files = pd.DataFrame(file_data)
        hierarchical_data = generate_hierarchical_structure(directory_path, file_data)
        df_hierarchy = pd.DataFrame(hierarchical_data)

        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df_files.to_excel(writer, index=False, sheet_name='RawData')
                df_hierarchy.to_excel(writer, index=False, sheet_name='AuditSheet')
            print(f"File list saved to {output_file}")
        except Exception as e:
            logging.error(f"Failed to save the file {output_file}. Error: {e}")
            print(f"An error occurred while saving the file. Please check the error.log for more details.")
    else:
        print("Invalid directory path. Please enter a valid path.")


def get_file_owner(file_path):
    if platform.system() == 'Windows': 
        try:
            import win32security
            security_descriptor = win32security.GetFileSecurity(file_path, win32security.OWNER_SECURITY_INFORMATION)
            owner_sid = security_descriptor.GetSecurityDescriptorOwner()
            name, _, _ = win32security.LookupAccountSid(None, owner_sid)
            return name
        except FileNotFoundError as e:
            logging.error(f"File not found when trying to get owner: {file_path}, error: {e}")
            return "Unknown"
    return ""

def list_files(dir_path, exclude_folders):
    file_data = [] 
    dir_path = Path(dir_path).resolve() 
    exclude_folders = [Path(folder).resolve() for folder in exclude_folders]  # Convert exclude folders to resolved Path objects

    for root, dirs, files in os.walk(dir_path, topdown=True):
        root_path = Path(root).resolve()
        
        # Exclude specified folders
        dirs[:] = [d for d in dirs if root_path.joinpath(d).resolve() not in exclude_folders]
        
        for file in files:
            file_path = root_path / file
            try:
                file_stat = os.stat(file_path)
                file_size_mb = file_stat.st_size / (1024 * 1024)
                date_created = datetime.datetime.fromtimestamp(file_stat.st_ctime).strftime('%Y-%m-%d')
                last_modified = datetime.datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d')

                file_data.append({
                    'File Name': file,
                    'File Type': file_path.suffix,
                    'File Path': str(file_path), 
                    'File Size (MB)': round(file_size_mb, 2),
                    'Date Created': date_created,
                    'Last Modified': last_modified,
                    'Owner': get_file_owner(str(file_path))
                })
            except FileNotFoundError as e:
                logging.error(f"File not found when accessing metadata: {file_path}, error: {e}")
                continue  # Skip to the next file

    return file_data

def generate_hierarchical_structure(dir_path, file_data):
    hierarchical_data = []
    seen_dirs = set()

    dir_path = Path(dir_path)

    sorted_files = sorted(file_data, key=lambda x: Path(x['File Path']).parts)

    for file_info in sorted_files:
        file_path = Path(file_info['File Path'])
        parts = file_path.relative_to(dir_path).parts

        cumulative_path = dir_path  # Initialize cumulative_path at the start of the loop
        for i, part in enumerate(parts[:-1]):  # Iterate over parts (excluding the last part which is the file)
            cumulative_path = cumulative_path / part
            if str(cumulative_path) not in seen_dirs:
                folder_size_mb = get_folder_size(cumulative_path) / (1024 * 1024)  # Convert to MB
                hierarchical_data.append({
                    'Name': f'=HYPERLINK("{cumulative_path}", "{ " " * 4 * i + part}")',
                    'Item Type': 'Folder',
                    'Owner': file_info.get('Owner', 'Unknown'),
                    'Date Created': file_info.get('Date Created', ''),
                    'Last Modified': file_info.get('Last Modified', ''),
                    'Size (MB)': round(folder_size_mb, 2),
                    'Path': f'=HYPERLINK("{cumulative_path}", "Open Folder")',
                    'Action': '',
                    'Rename as…': '',
                    'Move to…': '', 
                    'Consult staff?': '',
                    'Proposed retention': '',
                    'Justification': '',
                    'Notes': ''
                })
                seen_dirs.add(str(cumulative_path))

        # Add file to the hierarchical data
        file_size_mb = file_info['File Size (MB)']  # This assumes file sizes are already calculated in MB
        hierarchical_data.append({
            'Name': f'=HYPERLINK("{file_path}", "{ " " * 4 * (len(parts) - 1) + parts[-1]}")',
            'Item Type': 'File',
            'Owner': file_info['Owner'],
            'Date Created': file_info['Date Created'],
            'Last Modified': file_info['Last Modified'],
            'Size (MB)': file_size_mb,
            'Path': f'=HYPERLINK("{file_path}", "Open File")',
            'Action': '',
            'Rename as…': '',
            'Move to…': '',  
            'Consult staff?': '',
            'Proposed retention': '',
            'Justification': '',
            'Notes': ''
        })

    return hierarchical_data

def get_folder_size(folder_path):
    total_size = 0
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            try:
                if not os.path.islink(file_path):  # Skip if it's a symbolic link
                    total_size += os.path.getsize(file_path)
            except FileNotFoundError as e:
                logging.error(f"File not found: {file_path}, error: {e}")
    return total_size


def get_exclusion_list(dir_path, retry_limit=3):
    dir_path = Path(dir_path).resolve()
    # Get actual folder names and convert them to lowercase for a case-insensitive comparison
    actual_folders = {f.name.lower(): f for f in dir_path.glob('*') if f.is_dir()}
    retry_count = 0

    while True:
        exclude_input = input("Enter the folder names to exclude (comma-separated), 'none' to exclude nothing, or 'quit' to exit: ").strip()

        if exclude_input.lower() in ['quit', 'exit']:
            exit()

        if exclude_input.lower() == 'none':
            print("No directories will be excluded.")
            return []

        folder_names = [name.strip() for name in exclude_input.split(',') if name.strip()]
        valid_paths = []
        invalid_paths = []

        for folder_name in folder_names:
            folder_name_lower = folder_name.lower()
            if folder_name_lower in actual_folders:
                valid_paths.append(str(actual_folders[folder_name_lower].resolve()))
            else:
                invalid_paths.append(folder_name)

        if not invalid_paths:
            print("\nFolders selected for exclusion:")
            for path in valid_paths:
                print(path)

            user_confirmation = input("Confirm exclusions (yes/no), or 'quit' to exit: ").strip().lower()
            if user_confirmation in ['quit', 'exit']:
                exit()

            if user_confirmation == 'yes':
                return valid_paths

        print(f"Invalid or non-existent directories: {', '.join(invalid_paths)}. Please re-enter all folder names correctly.")
        retry_count += 1
        if retry_count >= retry_limit:
            print(f"Exceeded maximum retries ({retry_limit}). Exiting program.")
            exit()

def find_common_base_directory(file_paths):
    paths = [Path(p) for p in file_paths]
    common_base = paths[0].parent
    for path in paths:
        while not str(path).startswith(str(common_base)):
            common_base = common_base.parent
            if common_base == Path(common_base.root):
                # We've reached the root of the file system
                return common_base
    return common_base

def process_uploaded_file(desktop_path):
    file_path_input = input("...").strip()
    file_path = Path(file_path_input)

    if not file_path.exists():
        print(f"File not found: {file_path}")
        return

    df_raw = None

    if file_path.suffix.lower() == '.xlsx':
        xls = pd.ExcelFile(file_path)
        # Attempt to find the correct sheet based on the 'File Path' column
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            if 'File Path' in df.columns:
                df_raw = df
                break
        if df_raw is None:
            print("No suitable sheet found in the Excel file.")
            return
    else:
        df_raw = pd.read_csv(file_path, dtype=str)
        if 'File Path' not in df_raw.columns:
            print("The CSV file does not contain a 'File Path' column.")
            return

    # Ensure that df_raw is not None and contains 'File Path' column
    if df_raw is not None and 'File Path' in df_raw.columns:
        # If 'Owner' column is not in df_raw, calculate the owner information
        if 'Owner' not in df_raw.columns:
            df_raw['Owner'] = df_raw['File Path'].apply(lambda x: get_file_owner(x))
            
        file_paths = df_raw['File Path'].tolist()
        common_base = find_common_base_directory(file_paths)

        print(f"The determined base directory for hierarchy is: {common_base}")

        # Generate hierarchical data using the common base as the dir_path
        hierarchical_data = generate_hierarchical_structure(common_base, df_raw.to_dict(orient='records'))
        df_hierarchy = pd.DataFrame(hierarchical_data)

        valid_file_name = False
        while not valid_file_name:
            output_file_name = input("Enter the desired output file name (without extension): ").strip()
            if output_file_name and re.match("^[a-zA-Z0-9_-]*$", output_file_name):
                output_file = os.path.join(desktop_path, f'{output_file_name}_audit_file.xlsx')
                valid_file_name = True
            else:
                print("Invalid file name. Please use only letters, numbers, hyphens, and underscores.")

        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df_raw.to_excel(writer, index=False, sheet_name='RawData')
                df_hierarchy.to_excel(writer, index=False, sheet_name='AuditSheet')
            print(f"Hierarchical structure and raw data saved to {output_file}")
        except Exception as e:
            print("An error occurred while saving the file. Please check the error.log for more details.")
            logging.error(f"Failed to save the file {output_file}. Error: {e}")
    else:
        print("The uploaded file does not contain the required data to generate a hierarchy.")

def main():
    print("Select an option:")
    print("1: Audit a directory")
    print("2: Process an uploaded file for hierarchical structure")
    print("3: Exit")

    user_choice = input("Enter your choice (1/2/3): ").strip()

    if user_choice == '1':
        audit_directory_process(desktop_path)
    elif user_choice == '2':
        process_uploaded_file(desktop_path)
    elif user_choice == '3':
        exit()
    else:
        print("Invalid choice, please select 1, 2, or 3.")

if __name__ == '__main__':
    main()
