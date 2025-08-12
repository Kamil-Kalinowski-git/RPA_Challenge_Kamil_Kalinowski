# File: RPA_Challenge_Kamil_Kalinowski

# Libraries
import rpa
import pandas
import os
import shutil
import glob

def initialize_rpa():
    """Initializes the RPA environment and navigates to the challenge website."""
    try:
        rpa.init(visual_automation=True)
        rpa.url('https://rpachallenge.com/')
        return True, 'RPA environment initialized and configured for downloads.'
    except Exception as e:
        return False, f'Failed to initialize RPA: {e}'

def download_file_and_move(downloads_dir):
    """Downloads Excel file and moves it to the target directory."""
    try:
        user_downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')

        rpa.click('Download Excel')
        rpa.wait(5.0)
        
        excel_files = glob.glob(os.path.join(user_downloads_dir, '*.xlsx'))
        
        if not excel_files:
            return False, 'No Excel files found in the default downloads directory.'
            
        latest_file = max(excel_files, key=os.path.getctime)
        
        shutil.move(latest_file, downloads_dir)
        
        return True, f'File moved to {downloads_dir} successfully.'
    except Exception as e:
        return False, f'Failed to download or move the file: {e}'

def clear_download_directory(downloads_dir):
    """Clears the downloads directory after user confirmation."""
    if not os.listdir(downloads_dir):
        return True, f'The folder {downloads_dir} is already empty.'
    
    for attempt in range(3):
        response = input(f'The folder {downloads_dir} is not empty. Do you want to clear it and download a new file? (Y/n): ')
        if response.upper() == 'Y':
            try:
                shutil.rmtree(downloads_dir)
                os.makedirs(downloads_dir)
                return True, f'Folder {downloads_dir} was cleared.'
            except Exception as e:
                return False, f'Failed to clear {downloads_dir}. Reason: {e}'
        elif response.lower() == 'n':
            return False, 'Download cancelled by user. Exiting script.'
        else:
            print(f'Invalid input. Please enter Y for yes or n for no. You have {2 - attempt} attempt(s) left.')
            
    return False, 'Too many invalid attempts. Exiting program.'

def import_data(downloads_dir):
    """Reads data from the downloaded Excel file using pandas.""" 
    try:
        list_of_files = os.listdir(downloads_dir)
        if not list_of_files:
            return False, f'No files found in the downloads directory {downloads_dir}.'
        
        excel_file_name = list_of_files[0]
        excel_file_path = os.path.join(downloads_dir, excel_file_name)
        
        df = pandas.read_excel(excel_file_path)
        return True, df
    except FileNotFoundError:
        return False, f'File not found: {excel_file_path}'
    except Exception as e:
        return False, f'An error occurred while importing data: {e}'

def start_challenge():
    """Clicks the 'Start' button to begin the challenge."""
    try:
        rpa.click('Start')
        return True, 'RPA Challenge started successfully.'
    except Exception as e:
        return False, f'Failed to click the "Start" button: {e}'

def fill_form_with_data(df):
    """Fills out the form for each row of data and submits it using a loop."""
    try:
        for index, row in df.iterrows():
            rpa.type('//input[@ng-reflect-name="labelFirstName"]', row['First Name'])
            rpa.type('//input[@ng-reflect-name="labelLastName"]', row['Last Name '])
            rpa.type('//input[@ng-reflect-name="labelCompanyName"]', row['Company Name'])
            rpa.type('//input[@ng-reflect-name="labelRole"]', row['Role in Company'])
            rpa.type('//input[@ng-reflect-name="labelAddress"]', row['Address'])
            rpa.type('//input[@ng-reflect-name="labelEmail"]', row['Email'])
            rpa.type('//input[@ng-reflect-name="labelPhone"]', str(row['Phone Number']))
            rpa.click('//input[@value="Submit"]')
        return True, 'Form filled and submitted for all data rows.'
    except Exception as e:
        return False, f'An error occurred while filling the form: {e}'

def main():
    """Main function to orchestrate the RPA process."""
    print('Starting RPA Challenge automation...')
    downloads_dir = os.path.join(os.getcwd(), 'data', 'input')
    os.makedirs(downloads_dir, exist_ok=True)
    rpa_initialized = False
    
    try:
        # Step 1: Clear download directory
        success, message = clear_download_directory(downloads_dir)
        if not success:
            print(f'Error: {message}')
            return
        print(message)
        
        # Step 2: Initialize RPA
        success, message = initialize_rpa()
        if not success:
            print(f'Error: {message}')
            return
        rpa_initialized = True
        print(message)
        
        # Step 3: Download file
        success, message = download_file_and_move(downloads_dir)
        if not success:
            print(f'Error: {message}')
            return
        print(message)

        # Step 4: Import data
        success, data = import_data(downloads_dir)
        if not success:
            print(f'Error: {data}')
            return
        excel_data = data
        print('Data loaded successfully into a pandas DataFrame.')

        # Step 5: Start challenge
        success, message = start_challenge()
        if not success:
            print(f'Error: {message}')
            return
        print(message)
        
        # Step 6: Fill the form
        success, message = fill_form_with_data(excel_data)
        if not success:
            print(f'Error: {message}')
            return
        print(message)
        
        print('\nRPA Challenge completed successfully!')

    except Exception as e:
        print(f'\nAn unexpected error occurred during the process: {e}')
    finally:
        if rpa_initialized:
            rpa.close()
            print('\nBrowser closed.')

if __name__ == '__main__':
    main()