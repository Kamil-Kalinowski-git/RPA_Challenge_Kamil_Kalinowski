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

###################################################################################################################
# Checks if the downloads directory is not empty.
###################################################################################################################

def handle_downloads_directory(downloads_dir):

    if os.listdir(downloads_dir):
        for attempt in range(3):
            response = input(f'The folder {downloads_dir} is not empty. Do you want to clear it and download a new file? (y/n): ')
            if response.lower() == 'y':
                for file_to_be_deleted in os.listdir(downloads_dir):
                    file_path = os.path.join(downloads_dir, file_to_be_deleted)
                    try:
                        if os.path.isfile(file_path) or os.path.islink(file_path):
                            os.unlink(file_path)
                        elif os.path.isdir(file_path):
                            shutil.rmtree(file_path)
                    except Exception as e:
                        print(f'Failed to delete {file_path}. Reason: {e}')
                print(f'Folder {downloads_dir} was cleared.')
                return True
            elif response.lower() == 'n':
                print('Download cancelled by user. Exiting script.')
                return False
            else:
                print(f'Invalid input. Please enter y for yes or n for no. You have {2 - attempt} attempt(s) left.')

        print("\nToo many invalid attempts. Exiting program.")
        return False

    return True

###################################################################################################################
# Main
###################################################################################################################

def main():
    print('Starting RPA Challenge automation...')

###################################################################################################################
# Browser configuration
###################################################################################################################
  
    chrome_options = webdriver.ChromeOptions()

    # Downloads path
    downloads_dir = os.path.join(os.getcwd(), 'data', 'input')
    os.makedirs(downloads_dir, exist_ok=True)

    # Downloads directory
    prefs = {'download.default_directory': downloads_dir}
    chrome_options.add_experimental_option('prefs', prefs)

    # Check
    if not handle_downloads_directory(downloads_dir):
        return
    
###################################################################################################################
# Launching the browser
###################################################################################################################

    # Launch
    driver = webdriver.Chrome(options=chrome_options)

    # Max
    driver.maximize_window()

    # URL
    driver.get('https://rpachallenge.com/')

###################################################################################################################
# Downloading the Excel file
###################################################################################################################
    
    try:
        download_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='challenge.xlsx']"))
        )
        download_button.click()
        print('Excel file downloaded.')
        time.sleep(5) 

    except Exception as e:
        print(f'Cannot download the file: {e}')
        driver.quit()
        return

    print('Finished.')
    driver.quit()

###################################################################################################################
###################################################################################################################
###################################################################################################################

if __name__ == '__main__':
    main()