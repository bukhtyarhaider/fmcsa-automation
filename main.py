# This script retrieves a single record, averaging 7.746 seconds per fetch

import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# Path to the Excel file
file_path = '/Users/bukhtyarhaider/Projects/Automations/fmcsa/output/output.xlsx'

# Log file path
log_file_path = '/Users/bukhtyarhaider/Projects/Automations/fmcsa/logs/log.txt'

# Function to log activities
def log_activity(activity, start_time, end_time, extra_data=""):
    duration = (end_time - start_time).total_seconds()
    with open(log_file_path, 'a') as log_file:
        log_file.write(f"{activity} | Start: {start_time} | End: {end_time} | Duration: {duration} seconds | Data: {extra_data}\n")

# Read the Excel file and get the last mcNumber
try:
    df = pd.read_excel(file_path)
    mcNumber = df['MC NO.'].iloc[-1] + 1  # Increment the last mcNumber
except FileNotFoundError:
    # Initialize DataFrame with all necessary columns
    columns = [
        'MC NO.', 'USDOT Number', 'Entity Type', 'USDOT Status', 'Operating Authority Status',
        'Legal Name', 'Phone Number', 'Email Address', 'Power Units', 'Drivers', 'Carrier Operation', 'Status'
    ]
    df = pd.DataFrame(columns=columns)
    mcNumber = 1550000  # Default starting value if no file exists

# Specify the path to the ChromeDriver
service = Service('/Users/bukhtyarhaider/Projects/Automations/fmcsa/chromedriver-mac-arm64/chromedriver')
driver = webdriver.Chrome(service=service)

try:
    while True:  # Infinite loop
        activity_start_time = datetime.now()
        
        # Open the webpage at the start of each loop to reset the state
        driver.get("https://safer.fmcsa.dot.gov/CompanySnapshot.aspx")

        # Wait for the radio button to be clickable and click it
        radio_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="2"]'))
        )
        radio_button.click()

        # Find the input field by XPath and enter the value
        input_field = driver.find_element(By.XPATH, '//*[@id="4"]')
        input_field.clear()  # Clear previous value
        input_field.send_keys(str(mcNumber))

        # Wait for the search button to be clickable and click it
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/form/p/table/tbody/tr[4]/td/input'))
        )
        search_button.click()

        time.sleep(3)  # Wait for page to load

        # Check for the 'Record Inactive' or 'Record Not Found' page
        try:
            inactive_text = driver.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/p/font/b/i').text.strip()
            status = "Record Inactive" if "Record Inactive" in inactive_text else "Record Not Found"
        except Exception:
            status = "Record Active"
        try:
            operatingAuthorityStatusText = driver.find_element(By.XPATH, '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[8]/td/font/b[1]').text if status == "Record Active" else ''
        except Exception:
            operatingAuthorityStatusText = driver.find_element(By.XPATH, '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[8]/td').text if status == "Record Active" else ''
        
        operatingAuthorityStatus = operatingAuthorityStatusText.replace("For Licensing and Insurance details click here.", "").strip()
       
        if status == "Record Active":
            usdot_number = driver.find_element(By.XPATH, '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[5]/td[1]').text.strip()
            entity_type = driver.find_element(By.XPATH, '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[3]/td').text.strip()
            usdot_status = driver.find_element(By.XPATH, '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[4]/td[1]').text.strip()
            legal_name = driver.find_element(By.XPATH, '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[11]/td').text.strip()
            phone_number = driver.find_element(By.XPATH, '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[14]/td').text.strip()
            power_units = driver.find_element(By.XPATH, '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[17]/td[1]').text.strip()
            drivers = driver.find_element(By.XPATH, '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[17]/td[2]/font/b').text.strip()
            target_table_xpath = '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[21]/td/table'
            tables = driver.find_elements(By.XPATH, target_table_xpath)
            carrier_operation = driver.find_element(By.XPATH, '/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[21]/td/table/tbody/tr[2]/td[1]/table/tbody/tr[2]/td[1]').text.strip() if status == "Record Active" else '',

            if str(carrier_operation) == "('X',)":
                carrier_operation = "Interstate"
            else:
                carrier_operation = ''
            
            # Navigate to ai.fmcsa.dot.gov
            driver.get(f"https://ai.fmcsa.dot.gov/SMS/Carrier/{usdot_number}/Overview.aspx?FirstView=True")

            try:
                # Wait for the Carrier Registration button to be clickable and click it
                CarrierRegistrationBtn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="CarrierRegistration"]/a[1]'))
                    ).click()
                time.sleep(5)
                emailAddress = driver.find_element(By.XPATH, '//*[@id="regBox"]/ul[1]/li[7]/span').text.strip()
                error = ""

            except Exception:
                closeNoticeBtn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="simplemodal-container"]/a'))
                    ).click()
                emailAddress = ''
                error = f"{usdot_number} requires logged into the SMS"
            
            
         
        else:
            usdot_number = ''
            entity_type = ''
            usdot_status = ''
            legal_name = ''
            phone_number = ''
            power_units = ''
            drivers = ''
            carrier_operation = ''
            emailAddress = ''
            error = ''

        company_info = {
            'MC NO.': mcNumber,
            'USDOT Number': usdot_number,
            'Entity Type': entity_type,
            'USDOT Status': usdot_status,
            'Operating Authority Status': operatingAuthorityStatus,
            'Legal Name': legal_name,
            'Phone Number': phone_number,
            'Email Address': emailAddress,
            'Power Units': power_units,
            'Drivers': drivers,
            'Carrier Operation': carrier_operation,
            'Status': status,
            'Error': error,
        }

        # Append the new data to the DataFrame and save to Excel
        new_data = pd.DataFrame([company_info])
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel(file_path, index=False)
        
        if status != "Record Active":
            print(f"{status} for MC/MX Number = {mcNumber}")
        else:
            print(f"Record for MC/MX Number = {mcNumber} is added in the sheet")

        activity_end_time = datetime.now()
        log_activity(f"Processed MC Number {mcNumber}", activity_start_time, activity_end_time, extra_data=status)

        mcNumber += 1  # Increment mcNumber for the next loop

except Exception as e:
    error_time = datetime.now()
    log_activity("An error occurred", activity_start_time, error_time, extra_data=str(e))
    print("An error occurred:", e)
finally:
    driver.quit()  # Properly close the browser
