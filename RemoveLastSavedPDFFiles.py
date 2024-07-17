import os

def RemoveLastSavedPDFFiles(separate_sheets_output_directory_path, services_picklist_output_directory_path, forecast_output_directory_path):
    """
    Remove all PDF files in the specified directory paths.

    :param paths: A list of directory paths where PDF files will be removed.
    """

    paths = [separate_sheets_output_directory_path, services_picklist_output_directory_path, forecast_output_directory_path]
    for path in paths:
        # Check if the path exists and is a directory
        if os.path.isdir(path):
            # List all files in the directory
            for filename in os.listdir(path):
                # Construct the full file path
                file_path = os.path.join(path, filename)
                # Check if the file is a PDF and remove it
                if filename.endswith('.pdf'):
                    os.remove(file_path)
                    print(f"Removed: {file_path}")
        else:
            print(f"Directory does not exist: {path}")