from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import getpass
import os
import pandas as pd

INSTANCE_URL = "https://url goes here.atlassian.net/" #REPLACE WITH YOUR ATLASSIAN CLOUD URL
OBJECT_SCHEMA_URL = INSTANCE_URL + "jira/servicedesk/assets/object-schema/schema id goes here?typeId=" #Replace with schema url
folder_name = "jira_asset"
csv_directory = os.getcwd()+'/jira_asset'
dataframes = []
asset_ids = ['101','102...'] #Replace with asset ids
    
if __name__ == "__main__":
    email_id = input("Email ID : ")
    pw = getpass.getpass("Password : ")

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    if email_id != '' and pw != '':
        try:
            options = webdriver.ChromeOptions()
            #options.add_argument("--headless=new") #Remove the comment for headless
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            prefs = {"download.default_directory" : os.getcwd()+r"\jira_asset"}
            options.add_experimental_option("prefs",prefs)
            driver = webdriver.Chrome(options=options)
        except Exception as e:
            print(e)
            exit()

        print("Logging in...")
        driver.get(INSTANCE_URL)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID,'username')))
        driver.find_element(By.ID,'username').send_keys(email_id)
        driver.find_element(By.ID,'login-submit').click()
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID,'password')))
        driver.find_element(By.ID,'password').send_keys(pw)
        driver.find_element(By.ID,'login-submit').click()
        time.sleep(2)
        print("Exporting assets...")

        for index, asset_id in enumerate (asset_ids):
            driver.get(OBJECT_SCHEMA_URL+asset_id)
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,'//button[@data-testid="servicedesk-insight-bulk-actions.ui.button"]')))
            driver.find_element(By.XPATH,'//button[@data-testid="servicedesk-insight-bulk-actions.ui.button"]').click()
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,'//button[@data-testid="servicedesk-insight-bulk-actions.ui.export"]')))
            driver.find_element(By.XPATH,'//button[@data-testid="servicedesk-insight-bulk-actions.ui.export"]').click()
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,'//div[@data-testid="servicedesk-insight-export-objects-modal.ui.component--header"]')))
            time.sleep(2)
            driver.find_elements(By.XPATH,'//button[@class="css-1yeatxf"]')[1].click()
            time.sleep(5)
            files = os.listdir(csv_directory)
            files = [os.path.join(csv_directory, file) for file in files if os.path.isfile(os.path.join(csv_directory, file))]
            files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            if files:
                most_recent_file = files[0]
                new_file_name = driver.find_elements(By.XPATH,'//span[@class="i4kpzn-3 bmrxDe"]')[index].text + ".csv"
                os.rename(os.path.join(csv_directory, most_recent_file), os.path.join(csv_directory, new_file_name))
            time.sleep(2)

        print('Creating Excel workbook...')

        csv_files = [file for file in os.listdir(csv_directory) if file.endswith('.csv')]

        excel_writer = pd.ExcelWriter(os.getcwd()+'/jira_asset/assets_workbook.xlsx', engine='xlsxwriter')

        for csv_file in csv_files:
            csv_path = os.path.join(csv_directory, csv_file)
            df = pd.read_csv(csv_path)
            sheet_name = os.path.splitext(csv_file)[0]
            df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
            os.remove(csv_path)

        excel_writer.close()

        print("Compiled to assets_workbook Excel workbook.")
        driver.quit()
        exit()
    else:
        exit()