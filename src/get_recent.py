import os

def get_excel_files_in_users(directory):
    excel_files = []
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(('.xls', '.xlsx', '.xlsm')):
                file_path = os.path.join(root, file)
                excel_files.append(file_path)
    
    return excel_files

# Example usage
users_directory = "/Users/aradbm/Desktop"
excel_files_in_users = get_excel_files_in_users(users_directory)

if excel_files_in_users:
    print("Excel files found in the Users directory:")
    for file_path in excel_files_in_users:
        print(file_path)
else:
    print("No Excel files found in the Users directory.")

    