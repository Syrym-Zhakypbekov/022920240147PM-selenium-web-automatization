import asyncio
import glob
import json
import logging
import os
import time
from concurrent.futures import ProcessPoolExecutor
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# Additional file handler for error logging
file_handler = logging.FileHandler('log.txt')
file_handler.setLevel(logging.ERROR)
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger('').addHandler(file_handler)

def init_webdriver(download_dir):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("user-agent=Mozilla/5.0")
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--no-sandbox")
    prefs = {"download.default_directory": download_dir}
    chrome_options.add_experimental_option("prefs", prefs)
    service = Service(executable_path="C:\\Users\\syrym\\Downloads\\022820241222PM-selenium-web-automatization\\chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def download_and_rename(credentials, vin_data, process_id):
    logging.basicConfig(level=logging.INFO, format=f'%(asctime)s - Process {process_id} - %(levelname)s - %(message)s')
    download_dir = f"C:\\Users\\syrym\\Downloads\\022820241222PM-selenium-web-automatization\\excel-reports\\process_{process_id}"
    os.makedirs(download_dir, exist_ok=True)
    driver = init_webdriver(download_dir)
    driver.set_page_load_timeout(30)
    try:
        driver.get(credentials['link'])
        driver.find_element(By.ID, "login").send_keys(credentials['login'])
        driver.find_element(By.ID, "pass").send_keys(credentials['password'] + Keys.RETURN)
        for selected_vin_data in vin_data:
            try:
                vin_number = selected_vin_data["vin"]
                logging.info(f"Processing VIN: {vin_number}")
                driver.get(credentials['link'])
                time.sleep(2)
                search_input = driver.find_element(By.ID, "search")
                search_input.send_keys(vin_number + Keys.RETURN)
                logging.info(f"Searching for VIN: {vin_number}.")
                time.sleep(2)
                target_element = driver.find_element(By.XPATH, f"//td[contains(text(), '{vin_number}')]")
                ActionChains(driver).move_to_element(target_element).click(target_element).perform()
                logging.info("Clicking on the table element with the VIN number.")
                time.sleep(2)
                download_button = driver.find_element(By.ID, "getEXCEL")
                download_button.click()
                logging.info("Clicking on the 'Download Excel' button.")
                time.sleep(5)
                downloaded_files = glob.glob(f"{download_dir}\\order_*.xlsx")
                if downloaded_files:
                    new_file_name = f"{download_dir}\\{selected_vin_data['id']}_{selected_vin_data['type']}_{selected_vin_data['vin']}_{selected_vin_data['order']}.xlsx"
                    os.rename(downloaded_files[0], new_file_name)
                    logging.info(f"File renamed to {new_file_name}.")
                else:
                    logging.error("No downloaded file found to rename.")
                driver.get("https://eurotest.kz/order/")
                logging.info("Returned to the main page.")
            except Exception as vin_error:
                logging.error(f"Error processing VIN {vin_number}: {vin_error}")
                with open('log.txt', 'a') as log_file:
                    log_file.write(f"Error processing VIN {vin_number}: {vin_error}\n")
    except Exception as e:
        logging.error(f"An error occurred in Process {process_id}: {e}")
    finally:
        driver.quit()

async def open_login_search_click_download_and_rename_async():
    with open("C:\\Users\\syrym\\Downloads\\022820241222PM-selenium-web-automatization\\credentials.json", "r", encoding='utf-8') as file:
        credentials = json.load(file)
    with open("C:\\Users\\syrym\\Downloads\\022820241222PM-selenium-web-automatization\\data.json", "r", encoding='utf-8') as file:
        vin_data = json.load(file)
    chunks = [vin_data[i::20] for i in range(20)]
    loop = asyncio.get_running_loop()
    with ProcessPoolExecutor(max_workers=20) as executor:
        tasks = [loop.run_in_executor(executor, download_and_rename, credentials, chunk, i + 1) for i, chunk in enumerate(chunks)]
        await asyncio.gather(*tasks)

if __name__ == "__main__":
    asyncio.run(open_login_search_click_download_and_rename_async())
