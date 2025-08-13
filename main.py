# File: RPA_Challenge_Kamil_Kalinowski

# Libraries
import rpa
import pandas
import os
import shutil
import glob
import time
import datetime


def initialize_rpa():
    """Initializes the RPA environment and navigates to the challenge website."""
    try:
        rpa.init(visual_automation=True)
        rpa.url('https://rpachallenge.com/')
        return True, 'RPA environment initialized and configured for downloads.'
    except Exception as e:
        return False, f'Failed to initialize RPA: {e}'

def download_and_move_file(downloads_dir):
    """Downloads an Excel file to the project's root folder and moves it to the target directory."""
    try:
        rpa.click('Download Excel')
        rpa.wait(5.0)

        project_root = os.getcwd()
        excel_files = glob.glob(os.path.join(project_root, '*.xlsx'))
        # excel_files = os.path.join(os.path.expanduser('~'), 'Downloads')
        
        if not excel_files:
            return False, f'No Excel file found in the project root directory: {project_root}'
        
        latest_file = max(excel_files, key=os.path.getmtime)
        shutil.move(latest_file, downloads_dir)
        
        return True, f'File {os.path.basename(latest_file)} was successfully moved to {downloads_dir}.'
    except Exception as e:
        return False, f'Failed to download or move the file: {e}'

def clear_download_directory(downloads_dir):
    """Clears the downloads directory to ensure it's empty before downloading a new file."""
    try:
        if os.path.exists(downloads_dir):
            shutil.rmtree(downloads_dir)
        os.makedirs(downloads_dir)
        return True, f'The folder {downloads_dir} was successfully created.'
    except Exception as e:
        return False, f'Failed to clear or create the folder {downloads_dir}. Reason: {e}'

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

def save_results(output_dir):
    """Retrieves the final score and saves it to a text file."""
    try:
        time_now = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        final_score = rpa.read('//div[@class="message2"]')
        results_file_path = os.path.join(output_dir, f'{time_now}_result.txt')
        
        with open(results_file_path, 'w') as f:
            f.write(final_score)
        
        return True, f'The result has been saved to the file: {results_file_path}'
    except Exception as e:
        return False, f'Failed to retrieve or save the result: {e}'
    
def main():
    """Main function to orchestrate the RPA process."""
    print('Starting RPA Challenge automation...')
    downloads_dir = os.path.join(os.getcwd(), 'data', 'input')
    output_dir = os.path.join(os.getcwd(), 'data', 'output')
    os.makedirs(output_dir, exist_ok=True)
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
        success, message = download_and_move_file(downloads_dir)
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

        # Step 7: Save the results
        success, message = save_results(output_dir)
        if not success:
            print(f'Error: {message}')
            return
        print(message)

    except Exception as e:
        print(f'\nAn unexpected error occurred during the process: {e}')
    finally:
        if rpa_initialized:
            rpa.close()
            print('\nBrowser closed.')

if __name__ == '__main__':
    main()