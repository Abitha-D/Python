from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time

import gspread
from oauth2client.service_account import ServiceAccountCredentials

#import spreadsheetcode

import datetime

def Open_Website():
    # Open Chrome
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get('http://13.200.17.36/ecesisapp_git/tfs')
    return driver

def Login(driver):
    username = 'R'
    password = 'test123@'

    try:
        username_field = driver.find_element(By.NAME, 'User')  # Google search bar's name is 'q'
        password_field = driver.find_element(By.NAME, 'Password')

        username_field.clear()  # Clear the fields
        password_field.clear()

        username_field.send_keys(username)
        password_field.send_keys(password)

        login_button = driver.find_element(By.CLASS_NAME, 'btn ')
        login_button.send_keys(Keys.RETURN)

        # Wait for the page to load
        time.sleep(5)

        # Wait for the URL to change after login
        WebDriverWait(driver, 10).until(EC.url_to_be('http://13.200.17.36/ecesisapp_git/tfs/Home/applicationhome'))

    except Exception as e:
        print("An error occurred:", e)

def select_function_module(driver):
    # Find the second li element
    function_module = driver.find_elements(By.CSS_SELECTOR, 'ul#navigation-bar > li')[2]
    function_module.click()

def select_create_order_module(driver):
    # Wait for the submenu to appear
    submenu_xpath = "//ul[@module_id='103_module']"  # Corrected XPath for the submenu
    submenu = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, submenu_xpath)))

    # Click on the submodule with ID "10029_sub_module_id" (Create Order)
    create_order_submodule_element = submenu.find_element(By.ID, '10029_sub_module_id')
    create_order_submodule_element.click()

def client_name(driver, clientname):
    try:
        # Wait for the dropdown element to be clickable
        iframe_id = "framecontent"  # Replace with the actual ID or name of the iframe
        driver.switch_to.frame(iframe_id)
        dropdown_id = "ddlSubClient"
        dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, dropdown_id)))

        # Execute JavaScript click on the dropdown
        driver.execute_script("arguments[0].click();", dropdown)

        # Wait for the dropdown options to be visible
        dropdown_option_xpath = "//option[contains(text(), '{}')]".format(clientname)
        dropdown_option = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, dropdown_option_xpath)))

        # Click on the desired option
        dropdown_option.click()

    except Exception as e:
        print("An error occurred:", e)

def subject_address(driver, subjectaddress):
    # subject_address_input = driver.find_element(By.ID, 'txtsubaddress')
    subject_address_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, 'txtsubaddress')))
    subject_address_input.clear()  # Clear the fields
    subject_address_input.send_keys(subjectaddress)
    subject_address_input.send_keys(Keys.RETURN)

def order_no(driver, orderno):
    # order_no_input = driver.find_element(By.ID, 'txtorderportalno')
    order_no_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, 'txtorderportalno')))
    order_no_input.clear()  # Clear the fields
    order_no_input.send_keys(orderno)
    order_no_input.send_keys(Keys.RETURN)
def order_type(driver, ordertype):
    try:
        # Wait for the dropdown element to be clickable
        dropdown_id = "ddlorderType"
        dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, dropdown_id)))

        # Execute JavaScript click on the dropdown
        driver.execute_script("arguments[0].click();", dropdown)

        # Wait for the dropdown options to be visible
        dropdown_option_xpath = f"//option[text()='{ordertype}']"
        dropdown_option = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, dropdown_option_xpath)))

        # Click on the desired option
        dropdown_option.click()

    except Exception as e:
        print("An error occurred:", e)

def order_due_split(driver, duedatetime):
    # Split the duedatetime string by space to separate date and time
    date_str, time_str, time_zone = duedatetime.split()

    # Split the date string by hyphen to get month, day, and year
    month, day, year = date_str.split('-')

    # Reformat the date string as MM/DD/YYYY
    date = f"{month}/{day}/{year}" #Date: 06/28/2018

    # Combine the time string with AM/PM
    times = f"{time_str} {time_zone}" #Time: 03:25 AM

    return date, times

def order_due_date(driver, duedate):
    try:
        # due_date_input = driver.find_element(By.ID, 'txtdob')
        due_date_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'txtdob')))
        driver.execute_script('document.getElementById("txtdob").removeAttribute("readonly")')
        due_date_input.clear()
        due_date_input.send_keys(duedate)
        due_date_input.send_keys(Keys.RETURN)

    except Exception as e:
        print("An error occurred:", e)

def order_due_time(driver, duetime):
    try:
        # due_time_input = driver.find_element(By.ID, 'txtdobtime')
        due_time_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'txtdobtime')))
        due_time_input.clear()  # Clear the fields
        due_time_input.send_keys(duetime)
        due_time_input.send_keys(Keys.RETURN)
    except Exception as e:
        print("An error occurred:", e)

def time_zone(driver, timezone):
    try:
        EST = driver.find_element(By.XPATH, '//*[@value="EST"]')
        PST = driver.find_element(By.XPATH, '//*[@value="PST"]')
        MST = driver.find_element(By.XPATH, '//*[@value="PST"]')

        if timezone == 'PST':
            PST.click()
        elif timezone == 'MST':
            MST.click()
        else:
            EST.click()
    except Exception as e:
        print("An error occurred:", e)

def photo_grapher(driver, photographer):
    try:
        # Wait for the dropdown element to be clickable
        dropdown_id = "ddlpht"
        dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, dropdown_id)))

        # Execute JavaScript click on the dropdown
        driver.execute_script("arguments[0].click();", dropdown)

        if photographer == '' or photographer == 'Ecesis Photographer':
            # Wait for the dropdown options to be visible
            dropdown_option_xpath = f"//option[text()='Ecesis Photographer']"
            dropdown_option = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, dropdown_option_xpath)))

            # Click on the desired option
            dropdown_option.click()
        else:
            # Wait for the dropdown options to be visible
            dropdown_option_xpath = f"//option[text()='{photographer}']"
            dropdown_option = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, dropdown_option_xpath)))

            # Click on the desired option
            dropdown_option.click()

    except Exception as e:
        print("An error occurred:", e)

def photographer_notes(driver, photographernotes):
    try:
        # photo_notes_input = driver.find_element(By.ID, 'txtnotepg')
        photo_notes_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'txtnotepg')))
        photo_notes_input.clear()  # Clear the fields
        photo_notes_input.send_keys(photographernotes)
        photo_notes_input.send_keys(Keys.RETURN)
    except Exception as e:
        print("An error occurred:", e)

def portal_fun(driver, portal):
    try:
        # Wait for the dropdown element to be clickable
        dropdown_id = "ddlportal"
        dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, dropdown_id)))

        # Execute JavaScript click on the dropdown
        driver.execute_script("arguments[0].click();", dropdown)

        # Wait for the dropdown options to be visible
        dropdown_option_xpath = f"//option[text()='{portal}']"
        dropdown_option = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, dropdown_option_xpath)))

        # Click on the desired option
        dropdown_option.click()

    except Exception as e:
        print("An error occurred:", e)

def photographer_fee(driver, photographerfee):
    try:
        # photo_fee_input = driver.find_element(By.ID, 'txtphtfee')
        photo_fee_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'txtphtfee')))
        photo_fee_input.clear()  # Clear the fields
        photo_fee_input.send_keys(photographerfee)
        photo_fee_input.send_keys(Keys.RETURN)
    except Exception as e:
        print("An error occurred:", e)

def order_value(driver, ordervalue):
    try:
        # order_value_input = driver.find_element(By.ID, 'txtordervalue')
        order_value_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'txtordervalue')))
        order_value_input.clear()  # Clear the fields
        order_value_input.send_keys(ordervalue)
        order_value_input.send_keys(Keys.RETURN)
    except Exception as e:
        print("An error occurred:", e)

def remarks_fun(driver, remarks):
    try:
        # remarks_input = driver.find_element(By.ID, 'txtremarks')
        remarks_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'txtremarks')))
        remarks_input.clear()  # Clear the fields
        remarks_input.send_keys(remarks)
        remarks_input.send_keys(Keys.RETURN)
    except Exception as e:
        print("An error occurred:", e)

def saveorder(driver):
    try:
        save_button = driver.find_element(By.XPATH, '//button[@onclick="SaveData()"]')
        save_button.click()
    except Exception as e:
        print("An error occurred:", e)

def modal_okay(driver):
    try:
        driver.switch_to.default_content()
        sweet_alert = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "sweet-alert"))
        )
        # Locate all the options
        yes_button = sweet_alert.find_element(By.CLASS_NAME, "confirm")
        yes_button.click()

    except Exception as e:
        print("An error occurred:", e)

def fetch_orders_with_empty_done(spreadsheet):
    order_sheet = spreadsheet.worksheet('Orders')
    
    try:
        orders_data = order_sheet.get_all_values()
        orders_headers = orders_data[0]
        orders_data_rows = orders_data[1:]
        orders_data_rows.reverse()  # If you want the latest orders first
        orders_with_empty_done = []
        
        for row in orders_data_rows:
            if not row[-3]:  # Assuming "Done" column is the last column
                orders_with_empty_done.append({orders_headers[i]: row[i] for i in range(len(orders_headers))})
        
        return orders_with_empty_done
    except Exception as e:
        print("An error occurred:", e)

def fetch_order_values(spreadsheet):
    order_sheet = spreadsheet.worksheet('Orders')

    try:
        orders_data = order_sheet.get_all_values()
        orders_headers = orders_data[0]
        orders_data_rows = orders_data[1:]

        for index, row in enumerate(orders_data_rows, start=2):  # Start from 2 to match sheet index
            if not row[-3]:  # Assuming "Done" column is the last column
                order = {orders_headers[i]: row[i] for i in range(len(orders_headers))}
                return index, order  # Return row index and order values
        
        # If no such row is found, return None
        return None, None
    except Exception as e:
        print("An error occurred:", e)

def move_row(source_sheet, destination_sheet, copy_delete_index):
    row_values = source_sheet.row_values(copy_delete_index)     # Get values of the row to be moved
    destination_sheet.append_row(row_values)    # Append the row to the destination worksheet
    source_sheet.delete_rows(copy_delete_index)    # Delete the row from the source worksheet

def check_for_new_rows(spreadsheet):
    order_sheet = spreadsheet.worksheet('Orders')
    orders_data = order_sheet.get_all_values()

    for row in orders_data:
        if not row[-3]:  # Assuming "Done" column is the last column
            return True
    return False

def check_required_fields(ordervalues):
    required_fields = ['Client', 'Address', 'OrderType', 'Portal', 'PhotoFee', 'OrderValue']
    for field in required_fields:
        if not ordervalues.get(field):
            return True  # Return True if any required field is empty
    return False

def check_all_fields_empty(ordervalues):
    required_fields = ['Client', 'Address', 'OrderType', 'Portal', 'PhotoFee', 'OrderValue']
    
    # Check if all required fields are empty
    all_empty = all(ordervalues.get(field) == '' for field in required_fields)
    
    return all_empty



def main():
    driver = Open_Website()
    Login(driver)
    select_function_module(driver)
    select_create_order_module(driver)
    

    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheetapi-419712-5eef6519a06b.json', scope)
    client = gspread.authorize(credentials)
    # Open the spreadsheet
    spreadsheet = client.open('Create-Order')
    order_sheet = spreadsheet.worksheet('Orders')

    while True:
        if check_for_new_rows(spreadsheet):
            neworders = fetch_orders_with_empty_done(spreadsheet)  # fetch all the orders that are new
            n = len(neworders)
            for i in range(1, n+1): 
                row_index, ordervalues = fetch_order_values(spreadsheet) #fetch first row that is new

                if check_all_fields_empty(ordervalues):
                    order_sheet.delete_rows(row_index)
                    continue

                if check_required_fields(ordervalues):
                    print("Waiting for required fields to be filled...")
                    time.sleep(20)
                    continue
                
                start_time = time.time()
                # clientname = "Bang _Andrew Persaud"
                clientname = ordervalues.get("Client")
                client_name(driver, clientname)

                # subjectaddress = "Dummy 6327 Peach Ave, Corona, CA, 92880"
                subjectaddress = ordervalues.get("Address")
                subject_address(driver, subjectaddress)

                # orderno = "653421"
                orderno = ordervalues.get("OrderNo")
                order_no(driver, orderno)

                # ordertype = "EXT"
                ordertype = ordervalues.get("OrderType")
                order_type(driver, ordertype)

                # duedatetime = "04-28-2024 03:25 AM"
                duedatetime = ordervalues.get("OrderDue")
                if duedatetime != "":
                    duedate, duetime = order_due_split(driver, duedatetime)
                    order_due_date(driver, duedate)
                    order_due_time(driver, duetime)

                # timezone = ""
                timezone = ordervalues.get("TimeZone")
                time_zone(driver, timezone)

                # photographer = ""
                photographer = ordervalues.get("Photographer")
                photo_grapher(driver, photographer)

                # photographernotes = ""
                photographernotes = ordervalues.get("Notes")
                if photographernotes != "":
                    photographer_notes(driver, photographernotes)

                #portal = "Single Source" #convert to single letters and compare
                portal = ordervalues.get("Portal")
                portal_fun(driver, portal)

                # photographerfee = "10"
                photographerfee = ordervalues.get("PhotoFee")
                photographer_fee(driver, photographerfee)

                # ordervalue = "10"
                ordervalue = ordervalues.get("OrderValue")
                order_value(driver, ordervalue)

                    # remarks = ""
                remarks = ordervalues.get("Remarks")
                if remarks != "":
                    remarks_fun(driver, remarks)

                saveorder(driver)
                end_time = time.time()
                modal_okay(driver)

                select_create_order_module(driver)

                order_sheet.update_cell(row_index, len(ordervalues) - 2, "created")

                start_datetime = datetime.datetime.fromtimestamp(start_time)
                end_datetime = datetime.datetime.fromtimestamp(end_time)

                # Convert datetime objects to Indian Standard Time (IST)
                ist = datetime.timezone(datetime.timedelta(hours=5, minutes=30))
                start_ist = start_datetime.astimezone(ist)
                end_ist = end_datetime.astimezone(ist)

                # Convert datetime object to a string format
                start_ist_str = start_ist.strftime("%Y-%m-%d %H:%M:%S")
                end_ist_str = end_ist.strftime("%Y-%m-%d %H:%M:%S")

                order_sheet.update_cell(row_index, len(ordervalues) - 1, start_ist_str)
                order_sheet.update_cell(row_index, len(ordervalues), end_ist_str)

        else:
           # If no new rows are found, sleep for 10 seconds before checking again
           time.sleep(10)

if __name__ == "__main__":
   main()
   