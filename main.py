# File: RPA_Challenge_Kamil_Kalinowski

# Libraries
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import shutil



def handle_downloads_directory(downloads_dir):

###################################################################################################################
# Checks if the downloads directory is not empty.
###################################################################################################################

    if os.listdir(downloads_dir):
        for attempt in range(3):
            response = input(f'The folder {downloads_dir} is not empty. Do you want to clear it and download a new file? (y/n): ')
            if response.lower() == 'y':
                try:
                    shutil.rmtree(downloads_dir)
                    os.makedirs(downloads_dir)
                except Exception as e:
                    print(f'Failed to clear {downloads_dir}. Reason: {e}')
                    return False
                print(f'Folder {downloads_dir} was cleared.')
                return True
            elif response.lower() == 'n':
                print('Download cancelled by user. Exiting script.')
                return False
            else:
                print(f'Invalid input. Please enter y for yes or n for no. You have {2 - attempt} attempt(s) left.')

        print('\nToo many invalid attempts. Exiting program.')
        return False

    return True

def quit_and_return(driver, message = 'Error!'):
    print(message)
    driver.quit()
    exit()

def download_excel_file(driver):

###################################################################################################################
# Downloading the Excel file
###################################################################################################################
    
    try:
        download_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='challenge.xlsx']"))
        )
        download_button.click()
        print('Excel file downloaded.')
        time.sleep(2) 

    except Exception as e:
        return False, f'Cannot download the file: {e}'
    
    return True, None

def import_excel_file(downloads_dir):

###################################################################################################################
# Import the Excel file
###################################################################################################################

    list_of_files = os.listdir(downloads_dir)
    if not list_of_files:
        return False, 'No files found in the downloads directory.', None

    excel_file_name = list_of_files[0]
    excel_file_path = os.path.join(downloads_dir, excel_file_name)

    try:
        df = pd.read_excel(excel_file_path)
    except IOError as e:
        return False, f'Cannot read the Excel file: {e}', None
    except Exception as e:
        return False, f'Unknown error while reading the Excel file: {e}', None
    
    return True, None, df

def click_Start(driver):

###################################################################################################################
# click Start
###################################################################################################################

    try:
        time.sleep(1)
        start_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn-large.uiColorButton'))
        )
        start_button.click()
        print('RPA Challenge started.')
    except Exception as e:
        return False, f'Cannot click the START button: {e}'
    
    return True, None






def main():

    print('Starting RPA Challenge automation...')

    # Downloads path
    downloads_dir = os.path.join(os.getcwd(), 'data', 'input')
    os.makedirs(downloads_dir, exist_ok=True)

    # Check
    if not handle_downloads_directory(downloads_dir):
        return
    
###################################################################################################################
# Browser configuration
###################################################################################################################
  
    chrome_options = webdriver.ChromeOptions()

    # Change downloads directory
    prefs = {'download.default_directory': downloads_dir}
    chrome_options.add_experimental_option('prefs', prefs)
    
###################################################################################################################
# Launching the browser
###################################################################################################################

    # Launch
    try:
        driver = webdriver.Chrome(options=chrome_options)
    except Exception as e:
        print(f'Cannot start Chrome: {e}')
        return

    # Max
    driver.maximize_window()

    # URL
    driver.get('https://rpachallenge.com/')

###################################################################################################################

    success, error = download_excel_file(driver)
    if not success:
        quit_and_return(driver, error)

    success, error = click_Start(driver)
    if not success:
        quit_and_return(driver, error)

    success, error, df = import_excel_file(downloads_dir)
    if not success:
        quit_and_return(driver, error)






    driver.quit()

###################################################################################################################

if __name__ == '__main__':
    main()