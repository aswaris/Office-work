import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Path to the template Excel file
excel_file_path = r'C:\Users\Aslam\Documents\sentoWA.xlsx'

# Initialize the WebDriver
driver = webdriver.Chrome(r'C:\Users\Aslam\Desktop\webdriver\chromedriver.exe')

# Load WhatsApp Web
driver.get('https://web.whatsapp.com')
input("Press Enter after scanning QR code...")

# Read the mobile numbers from the Excel file
df = pd.read_excel(excel_file_path)
mobile_numbers = df['Mobile Number'].astype(str).tolist()
file_paths = df['File Path'].tolist()

# Iterate over the mobile numbers and file paths simultaneously
for mobile_number, file_path in zip(mobile_numbers, file_paths):
    # Check if the mobile number contains the country code prefix
    if not mobile_number.startswith('+'):
        mobile_number = '+' + mobile_number

    # Open a new chat with the mobile number
    search_box = driver.find_element(By.CSS_SELECTOR, '.selectable-text.copyable-text.iq0m558w')
    search_box.send_keys(mobile_number)
    time.sleep(2)
    search_box.send_keys(Keys.ENTER)
    time.sleep(2)

     # Check if the file path is valid
    if os.path.isfile(file_path):
        # Attach the PDF file
        attachment_icon = driver.find_element(By.XPATH, '//div[@title="Attach"]')
        attachment_icon.click()

        document_icon = driver.find_element(By.XPATH, '//input[@accept="*"]')
        document_icon.send_keys(file_path)
        time.sleep(2)

        send_button = driver.find_element(By.CSS_SELECTOR, "span[data-testid='send']")
        send_button.click()

        time.sleep(2)
    else:
        print(f"File not found: {file_path}")

    # Clear the search box for the next mobile number
    search_box = driver.find_element(By.CSS_SELECTOR, '.selectable-text.copyable-text.iq0m558w')
    search_box.clear()
    time.sleep(2)

# Close the browser
driver.quit()