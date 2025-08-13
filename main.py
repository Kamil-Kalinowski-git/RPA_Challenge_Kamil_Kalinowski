# File: RPA_Challenge_Kamil_Kalinowski

import os
import shutil
import time
import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


def initialize_selenium(downloads_dir):
    """Initializes Selenium WebDriver and navigates to the challenge website."""
    try:
        chrome_options = Options()
        prefs = {
            "download.default_directory": downloads_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
        chrome_options.add_experimental_option("prefs", prefs)
        # chrome_options.add_argument("--headless=new")  # Uncomment if you want headless mode

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=chrome_options)
        driver.maximize_window()
        driver.get("https://rpachallenge.com/")

        return True, driver, "Selenium initialized and configured."
    except Exception as e:
        return False, None, f"Failed to initialize Selenium: {e}"

def download_file(driver, downloads_dir):
    """Downloads an Excel file to the project's root folder."""
    try:
        download_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Download Excel')]"))
        )
        download_button.click()
        time.sleep(3)
        return True, f"File was successfully downloaded {downloads_dir}."
    except Exception as e:
        return False, f"Failed to download or move the file: {e}"


def clear_download_directory(downloads_dir):
    """Clears the downloads directory to ensure it's empty before downloading a new file."""
    try:
        if os.path.exists(downloads_dir):
            shutil.rmtree(downloads_dir)
        os.makedirs(downloads_dir, exist_ok=True)
        return True, f"The folder {downloads_dir} was successfully created."
    except Exception as e:
        return False, f"Failed to clear the folder {downloads_dir}: {e}"


def import_data(downloads_dir):
    """Reads data from the downloaded Excel file using pandas."""
    try:
        excel_file_path = os.path.join(downloads_dir, "challenge.xlsx")
        df = pd.read_excel(excel_file_path)
        return True, df
    except FileNotFoundError:
        return False, f"File not found in: {downloads_dir}"
    except Exception as e:
        return False, f"An error occurred while importing data: {e}"


def start_challenge(driver):
    """Clicks the Start button to begin the challenge."""
    try:
        start_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Start')]"))
        )
        start_btn.click()

        time.sleep(3)

        return True, "RPA Challenge started successfully."
    except Exception as e:
        return False, f"Failed to click the Start button: {e}"

def fill_form_with_data(driver, df):
    """Fills out the form for each row of data and submits it using a loop."""
    try:
        for index, record in df.iterrows():
            driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelFirstName']").send_keys(record[0])
            driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelLastName']").send_keys(record[1])
            driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelCompanyName']").send_keys(record[2])
            driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelRole']").send_keys(record[3])
            driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelAddress']").send_keys(record[4])
            driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelEmail']").send_keys(record[5])
            driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelPhone']").send_keys(record[6])
            driver.find_element(By.XPATH, "//input[@value='Submit']").click()      
        return True, "Form filled."
    except Exception as e:
        return False, f"An error occurred while filling the form: {e}"


def save_results(driver, output_dir):
    """Retrieves the final score and saves it to a text file."""
    try:
        time_now = datetime.datetime.now().strftime('%Y%m%d%H%M%S')

        final_score_el = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.message2'))
        )
        final_score = final_score_el.text.strip()

        results_file_path = os.path.join(output_dir, f'{time_now}_result.txt')
        with open(results_file_path, 'w', encoding='utf-8') as f:
            f.write(final_score)

        return True, f"The result has been saved to the file: {results_file_path}"
    except Exception as e:
        return False, f"Failed to retrieve or save the result: {e}"


def main():
    """Main function to orchestrate the RPA process."""
    print("Starting RPA Challenge automation...")
    downloads_dir = os.path.join(os.getcwd(), "data", "input")
    output_dir = os.path.join(os.getcwd(), "data", "output")
    os.makedirs(downloads_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    

    driver = None

    try:
        # Step 1: Clear download directory
        success, message = clear_download_directory(downloads_dir)
        if not success:
            print(f"Error: {message}")
            return
        print(message)

        # Step 2: Initialize Selenium
        success, driver, message = initialize_selenium(downloads_dir)
        if not success:
            print(f"Error: {message}")
            return
        print(message)

        # Step 3: Download file
        success, message = download_file(driver, downloads_dir)
        if not success:
            print(f"Error: {message}")
            return
        print(message)

        # Step 4: Import data
        success, data = import_data(downloads_dir)
        if not success:
            print(f"Error: {data}")
            return
        excel_data = data
        print("Data loaded successfully.")

        # Step 5: Start challenge
        success, message = start_challenge(driver)
        if not success:
            print(f"Error: {message}")
            return
        print(message)

        # Step 6: Fill the form
        success, message = fill_form_with_data(driver, excel_data)
        if not success:
            print(f"Error: {message}")
            return
        print(message)

        # Step 7: Save the results
        success, message = save_results(driver, output_dir)
        if not success:
            print(f"Error: {message}")
            return
        print(message)

    except Exception as e:
        print(f"\nAn unexpected error occurred during the process: {e}")
    finally:
        if driver is not None:
            driver.quit()
            print("\nBrowser closed.")


if __name__ == "__main__":
    main()