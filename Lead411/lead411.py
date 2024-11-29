import time
import pandas as pd
import random
from datetime import date
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

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

file_path = get_file_path()
file_exists = os.path.exists(file_path)  # Check if the file exists

# Set up Selenium and the web driver
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')  # Maximize the browser window
# options.add_argument('--headless')  # Run headlessly
driver = webdriver.Chrome(service=service, options=options)

# Define login details and URLs
login_url = 'http://web.lead411.com/login'


# Step 1: Log in to Lead411
driver.get(login_url)

data = []

time.sleep(200)
print("Data extraction well strat in 1 min")
time.sleep(60)
print("Data extraction started")
for i in range(17, 100):
    rows = driver.find_elements(By.XPATH, f"//tr[@index='{i}']")
    

    # Iterate over the rows to extract data
    for row in rows:
        name_element = row.find_element(By.XPATH, "./td[2]")
        name = row.find_element(By.XPATH, "./td[2]").text.split('\n')
        linkedin_link = None
        try:
            linkedin_element = name_element.find_element(By.CSS_SELECTOR, "a[href*='linkedin.com']")
            linkedin_link = linkedin_element.get_attribute('href') if linkedin_element else None
        except NoSuchElementException:
            linkedin_link = None
        company = row.find_element(By.XPATH, "./td[3]").text.split('\n')[0]
        
        time.sleep(random.randint(10, 20))
        row.find_element(By.XPATH,"./td[3]/div/div/div[1]/span").click()
        
        company_location_elements=driver.find_elements(By.XPATH, f"//ul[contains(.,'{company}')]/li[3]/div[2]")
        if company_location_elements:
            # Extracting the text and joining the address parts
            company_location_lines = [element.text for element in company_location_elements]
            address_lines = company_location_lines[0].split('\n')
            company_location = ', '.join(address_lines)
        else:
            company_location = ''
        
        email = row.find_element(By.XPATH, "./td[4]").text
        phone_elements = row.find_element(By.XPATH, "./td[5]").text.split('\n')
        # Filter out unwanted parts from the phone numbers
        phones = [phone for phone in phone_elements if phone.isdigit() or phone.startswith('+')]

        # Append structured data to the list
        data.append([name[0], name[1],linkedin_link, company, email, phones[0] if len(phones) >=1  else None, phones[1] if len(phones) >= 2 else None, name ,company_location,phones if len(phones) >= 1 else None])
        target_element = row.find_element(By.XPATH, "./td[3]/div/div/div[1]/span")
        time.sleep(random.randint(1, 5))
        driver.execute_script("arguments[0].scrollIntoView(true);", target_element)

    # Create a DataFrame from the data
    df = pd.DataFrame(data, columns=['name', 'designation','Linkedin', 'company', 'email', 'phone 1', 'phone 2','Raw Name','company_location','Raw Phone'])

    # Write the DataFrame to an Excel file
    df.to_excel(file_path, index=False)

    # Print a confirmation message
    print(f"{i}Data has been written to {file_path}")


