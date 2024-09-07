from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import time
from datetime import datetime
import pandas as pd
#import pyperclip
import os
import logging


logging.basicConfig(
    filename="alethea.log",
    encoding="utf-8",
    filemode="a",
    format="{asctime} - {levelname} - {message}",
    style="{",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger=logging.getLogger() 
logger.setLevel(logging.INFO)
YES_ENTER = 1
NO_ENTER = 0
YES_UNBLUR = 1
NO_UNBLUR = 0

# Global variables
driver = None
short_wait_time = 1
standard_wait_time = 2
long_wait_time = 5

def save_excel(excel_data, excel_file):
    try:
        excel_data.to_excel(excel_file, index=False)
    except Exception as e:
        print("Error: ", e)
        logging.error("Error: " + e)

def wait(seconds):
    time.sleep(seconds)
    
def form_type_data(find_method, identifier, data, press_enter_after=0, unblur_after=0):
    global driver
    global standard_wait_time
    field = driver.find_element(find_method, identifier)
    field.send_keys(data)
    time.sleep(short_wait_time)
    if press_enter_after:
        field.send_keys(Keys.RETURN)  # Press Enter
        time.sleep(short_wait_time)
    if unblur_after:
        driver.execute_script("arguments[0].blur();", field)


def form_paste_data(find_method, identifier, data, press_enter_after=1, unblur_after=0, special_script=''):
    global driver
    global short_wait_time
    field = driver.find_element(find_method, identifier)
    if special_script != '':
        driver.execute_script(special_script, field)
    field.click()  # Click the field to focus
    temp = pyperclip.paste() # save current clipboard data
    pyperclip.copy(data)
    field.send_keys(Keys.CONTROL, 'v')  # Paste from clipboard
    pyperclip.copy(temp) # restore saved data to clipboard
    time.sleep(short_wait_time)
    if press_enter_after:
        field.send_keys(Keys.RETURN)  # Press Enter
        time.sleep(short_wait_time)
    if unblur_after:
        driver.execute_script("arguments[0].blur();", field)

def form_submit(find_method, identifier):
    global driver
    field = driver.find_element(find_method, identifier)
    field.submit()

def printplus(msg):
    print(msg)
    logging.info(msg)
    
# main function - loop through excel file, if a message is found which is both due (24 hours aka 86400 seconds after last time) and "ready" status, will end after submitting 1 message
def main(settings_arr):
    global driver
    global standard_wait_time
    username = settings_arr[0]
    password = settings_arr[1]
    excel_file = settings_arr[2]
    
    logging.info("Starting script with user: " + username);
    #print(os.getcwd());
    # Read data from Excel file using pandas
    try:
        excel_data = pd.read_excel(excel_file, dtype = object, engine='openpyxl')
    except Exception as e:
        print("Error reading Excel file:", e)
        logging.error("Error reading Excel file: " + e)
        return

    # check for any messages due to be sent
    # Iterate through the rows of the Excel file
    #print("Current timestamp:" + str(time.time()));
    for index, row in excel_data.iterrows():
        go = 0
        #print()
        #print(row)
        last_time = datetime.strptime(row['last time'], '%I:%M %p on %b %d, %Y')
        #print("last message time: " + row['last time'])
        #print(row['econsultid'])
        #print(row['message'])
            
        if row['status'] == 'ready':
            msg = row['name'] + " ready"
            if time.time() > last_time.timestamp()+86400:
                msg += ", due"
                go = 1
                
            else:
                 msg += ", not due"
        else:  
            msg = row['name'] + " not ready"
        printplus(msg)
        
        if go == 1:
            # Create a new instance of Microsoft Edge WebDriver
            driver = webdriver.Edge()
            driver.set_window_size(1400, 1000)

            # Open the econsult messaging page, which should load the login screen if not logged in
            driver.get('https://aletheamd.com/dashboard/secure-messaging?econsultId='+row['econsultid'])
            
            if driver.current_url == "https://aletheamd.com/login":
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']")))
                form_type_data(By.CSS_SELECTOR, "input[type='email']", username)
                # Find the username and password input fields and fill them
                form_type_data(By.CSS_SELECTOR, "input[type='password']", password, 1)
                
                WebDriverWait(driver, 10).until(EC.url_to_be("https://aletheamd.com/dashboard/secure-messaging?econsultId="+row['econsultid']))

                # Type out message
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "textarea[placeholder='Type a message']")))
                form_type_data(By.CSS_SELECTOR, "textarea[placeholder='Type a message']", row['message']) 
                form_type_data(By.CSS_SELECTOR, "textarea[placeholder='Type a message']", [Keys.CONTROL, Keys.RETURN]) # ENTER will go to new line, CTRL+ENTER to send
                
                printplus(row['name'] + " reply sent")
                excel_data.at[index, 'status'] = 'sent';
                save_excel(excel_data, excel_file);
                driver.quit()
                exit(0)
    printplus("All lines parsed, no actionable items found. Exiting.")
    

def get_settings():
    try:
        with open("settings.txt", "r") as file:
            lines = file.readlines()
            suser = lines[0].split(":")[1].strip()
            spass = lines[1].split(":")[1].strip()
    except Exception as e:
        printplus("Error: "+e)
    arr = [suser, spass, 'alethea.xlsx'];
    return arr

# Run the main function
settings = get_settings()
main(settings)

