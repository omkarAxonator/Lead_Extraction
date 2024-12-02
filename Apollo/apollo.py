# Import necessary modules
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from datetime import date
import csv
import os
import subprocess
import json

# Function to get ordinal suffix for a day
def get_ordinal_suffix(day):
    if 4 <= day <= 20 or 24 <= day <= 30:
        suffix = "th"
    else:
        suffix = ["st", "nd", "rd"][day % 10 - 1]
    return suffix

def get_date():
    # Get today's date
    today = date.today()
    day = today.day
    month = today.strftime("%B").lower()
    year = today.year
    ordinal_suffix = get_ordinal_suffix(day)
    formatted_date = f"{day}{ordinal_suffix}_{month}_{year}"
    return formatted_date

def get_file_path():
    # Ensure the target folder exists
    folder_name = "Scraped_Lead_Data"
    os.makedirs(folder_name, exist_ok=True)  # Create the folder if it doesn't exist
    date = get_date()

    # Write the DataFrame to an Excel file
    csv_file = f"{date}_extracted_data.csv"
    file_path = os.path.join(folder_name, csv_file)
    return file_path

def read_file(file_path):
    with open(file_path, 'r') as file:
        return file.read()
    
def write_file(file_path,file_content):
    with open(file_path, 'w') as file :
        file.write(file_content)

config_file = 'Config.json'
config = read_file(config_file)
config_json = json.loads(config)
Start_Page = config_json['Start_Page']
end_page = config_json['end_page']
chrome_exe_path = config_json['chrome_exe_path']
chrome_debug_path = config_json['chrome_debug_path']
debugging_port = config_json['debugging_port']

# Launch Chrome in debugging mode (PowerShell-compatible command)
run_chrome_in_debugging = [
    "powershell.exe",
    "-Command",
    rf"& '{chrome_exe_path}' --remote-debugging-port={debugging_port} --user-data-dir='{chrome_debug_path}'"
]

print("Launching Chrome in debugging mode...")
subprocess.Popen(run_chrome_in_debugging)

# Wait for Chrome to launch
time.sleep(5)

# Set up Selenium and the web driver
service = Service(ChromeDriverManager().install())

# Set the debugging URL and existing user data directory
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

# Connect to the running Chrome instance
driver = webdriver.Chrome(service=service, options=chrome_options)

# Define login details and URLs
filter_url = 'https://app.apollo.io/#/people?page=1&personLocations[]=United%20States&sortAscending=false&sortByField=%5Bnone%5D&organizationIndustryTagIds[]=5567cdde73696439812c0000&organizationIndustryTagIds[]=5567ce2773696454308f0000&personTitles[]=facility%20manager&personTitles[]=facilities%20director'
file_path = get_file_path()


# Step 1: Open the login page
driver.get(filter_url)
time.sleep(3)

# Step 4: Wait for the rows container to load
print(f"Started - {filter_url}")
time.sleep(60)
try:
    rows_container = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div.zp_tFLCQ[role="rowgroup"]'))
    )
    print("Rows container loaded.")
except TimeoutException:
    print("Rows container did not load in time.")
    driver.quit()
    exit()


# Step 5: Extract data from rows
for page in range(Start_Page, end_page+1):  # Adjust range as needed
    print(f"\nStarted Page : {page}")
    try:
        # Re-locate the rows container after each page change
        rows_container = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.zp_tFLCQ[role="rowgroup"]'))
        )
        rows = rows_container.find_elements(By.CSS_SELECTOR, 'div.zp_hWv1I[role="row"]')

        for row in rows:
            data = []

            try:
                # Locate the user name
                name_div = row.find_element(By.CSS_SELECTOR, 'div.zp_TPCm2.zp_PTp8r')
                name_anchor = name_div.find_element(By.CSS_SELECTOR, 'a.zp_p2Xqs.zp_v565m')
                user_name = name_anchor.text

                # Locate spans and extract title, company name, and location
                spans = row.find_elements(By.CSS_SELECTOR, 'span.zp_xvo3G')
                if len(spans) >= 3:
                    title = spans[0].text
                    company_name = spans[1].text
                    location = spans[2].text
                else:
                    title = company_name = location = "N/A"  # Handle missing spans gracefully

                # Split location into city and state
                if location != "N/A" and ", " in location:
                    city, state = location.split(", ", 1)  # Split on the first comma and space
                else:
                    city, state = "N/A", "N/A"  # Default values if the location format is unexpected

                # Extract LinkedIn link
                anchor_tags = row.find_elements(By.CSS_SELECTOR, 'a.zp_p2Xqs.zp_qe0Li.zp_S5tZC')
                linkedin_link = anchor_tags[1].get_attribute('href') if len(anchor_tags) > 1 else "N/A"

                
                employee_size_tag = row.find_elements(By.CSS_SELECTOR, 'span.zp_mE7no')
                employee_size = employee_size_tag[0].text
                # Append the extracted information to the data list
                data.append({
                    "Name": user_name,
                    "Title": title,
                    "Company Name": company_name,
                    "City": city,
                    "State": state,
                    "linkedin_link": linkedin_link,
                    "Employee_Size" : employee_size
                })
                

                with open(file_path, mode='a', newline='', encoding='utf-8') as file:
                    file_exists = os.path.exists(file_path)  # Check if the file exists
                    writer = csv.DictWriter(file, fieldnames=["Name", "Title", "Company Name", "City", "State", "linkedin_link","Employee_Size"])
                    if not file_exists:
                        writer.writeheader()  # Write the header only if the file does not exist
                    writer.writerows(data)

            except NoSuchElementException:
                print("Some elements were not found in a row. Skipping...")

    except TimeoutException:
        print(f"Rows container did not load on page {page}. Skipping...")
    except NoSuchElementException:
        print(f"No rows found on page {page}.")
    
    print(f"Completed {page} !!!!")
    config_json['Start_Page'] = page + 1
    # Convert the updated dictionary to JSON string
    updated_config = json.dumps(config_json, indent=4)
    write_file(config_file,updated_config)
    if Start_Page == end_page :
        break
    time.sleep(60)

print(f"Data extraction complete. Records have been appended to {file_path}")