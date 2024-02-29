import os
import shutil

def gather_excel_files(source_dir, destination_dir):
    # Create the destination directory if it doesn't exist
    os.makedirs(destination_dir, exist_ok=True)
    
    # Iterate through each process folder
    for i in range(1, 21):
        process_folder = os.path.join(source_dir, f'process_{i}')
        if os.path.exists(process_folder):
            # Iterate through files in the process folder
            for file_name in os.listdir(process_folder):
                if file_name.endswith('.xlsx'):
                    source_file_path = os.path.join(process_folder, file_name)
                    destination_file_path = os.path.join(destination_dir, file_name)
                    # Move the file to the destination directory
                    shutil.move(source_file_path, destination_file_path)

if __name__ == "__main__":
    source_directory = r'C:\\Users\\syrym\\Downloads\\022820241222PM-selenium-web-automatization\\excel-reports'
    destination_directory = r'C:\\Users\\syrym\\Downloads\\022820241222PM-selenium-web-automatization\\all_files'
    gather_excel_files(source_directory, destination_directory)
