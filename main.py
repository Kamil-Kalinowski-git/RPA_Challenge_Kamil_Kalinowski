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