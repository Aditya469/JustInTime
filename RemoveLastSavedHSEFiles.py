from configparser import ConfigParser
import os

def RemoveLastSavedHSEFiles(hse_directory_path):
    """
    Remove all PDF files in the specified directory paths.

    :param paths: A list of directory paths where PDF files will be removed.
    """

    paths = [hse_directory_path]
    for path in paths:
        # Check if the path exists and is a directory
        if os.path.isdir(path):
            # List all files in the directory
            for filename in os.listdir(path):
                # Construct the full file path
                file_path = os.path.join(path, filename)
                # Check if the file is a PDF and remove it
                if filename.endswith('.hse'):
                    os.remove(file_path)
                    print(f"Removed: {file_path}")
        else:
            print(f"Directory does not exist: {path}")


#Example usage
# if __name__ == "__main__":
#     config = ConfigParser()
#     config.read('config.ini')  # Make sure 'config.ini' is in your application's directory
#     hse_directory_path = config['PATHS']['hse_directory_path'] 
#     RemoveLastSavedHSEFiles(hse_directory_path)