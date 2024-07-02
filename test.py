import pandas as pd
from selenium import webdriver
import time

# Load data from Excel sheet
data = pd.read_excel('file_path_and_mobile.xlsx')

# Initialize Chrome WebDriver
driver = webdriver.Chrome()

# Open WhatsApp Web
driver.get('https://web.whatsapp.com/')
time.sleep(15)  # Allowing time to scan the QR code

# Iterate through each row in the Excel sheet
for index, row in data.iterrows():
    file_path = row['file_path']
    mobile_number = str(row['mobile_number'])
    
    # Compose the message with the PDF file
    message = f"Sending you the PDF file {file_path}"
    
    # Locate the chat input field
    chat_box = driver.find_element_by_xpath('//div[@contenteditable="true"][@data-tab="6"]')
    
    # Enter the mobile number
    chat_box.send_keys(mobile_number)
    time.sleep(2)
    chat_box.send_keys('\ue007')  # Press Enter to open chat
    
    # Attach the PDF file
    attachment = driver.find_element_by_xpath('//input[@accept="*"]')
    attachment.send_keys(file_path)
    time.sleep(2)
    
    # Send the message
    chat_box = driver.find_element_by_xpath('//div[@contenteditable="true"][@data-tab="6"]')
    chat_box.send_keys(message)
    time.sleep(2)
    chat_box.send_keys('\ue007')  # Press Enter to send message
    time.sleep(2)

# Close the WebDriver
driver.quit()

