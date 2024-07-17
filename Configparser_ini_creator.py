from configparser import ConfigParser
import os

def updateConfigForPathLocations():
    # Get the directory of the current script
    script_directory = os.path.dirname(os.path.abspath(__file__))

    # Create the configparser object
    config = ConfigParser()

    # Dynamically construct paths based on the script directory
    config['PATHS'] = {
        'hse_directory_path': os.path.join(script_directory, 'HSE Files/'),
        'logo_path': os.path.join(script_directory, 'Logo/Ecam_logo.jpg'),
        'icon_path': os.path.join(script_directory, 'Logo/DataMile_logo.jpeg'),
        'output_directory_path': script_directory,
        'separate_sheets_output_directory_path': os.path.join(script_directory, 'Report/Picklist/Separate_sheets'),
        'excel_file_path': os.path.join(script_directory, 'SalesOrder_{current_date}.xlsx'),
        'full_picklist_pdf_output_directory_path': os.path.join(script_directory, 'Report/Picklist/FullPicklist'),
        'forecast_output_directory_path': os.path.join(script_directory, 'Report/SalesForecast'),
        'separate_pdf_output_directory_path': os.path.join(script_directory, 'Report/Picklist/Separate_sheets'),
        'required_prices_path': os.path.join(script_directory, 'RequiredPrices/RequiredPrices.xlsx'),
        'services_picklist_output_directory_path': os.path.join(script_directory, 'Report/Picklist/Services')
    }

    # Write the configuration file to 'config.ini'
    with open('config.ini', 'w') as configfile:
        config.write(configfile)

# Example usage
if __name__ == "__main__":
    updateConfigForPathLocations()