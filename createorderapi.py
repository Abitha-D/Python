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

from datetime import datetime, timedelta
import pytz


import requests
import json

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

def fetch_required_values(GetRequiredValuesAPI, SelectrequestBody):
    try:
        response = requests.post(GetRequiredValuesAPI, json=SelectrequestBody)

        if response.ok == True:
            responseData = response.json()
            apiStatus = responseData['apiStatus']

            if apiStatus == 0:
                result = json.loads(responseData['result'])
                client_id = result[0]['ClientId']
                portal_id = result[0]['PortalId']
                photographer_id = result[0]['PhotographerId']
                ordertype_id = result[0]['OrderTypeId']
                zipcode_no = result[0]['Zipcode']
                
                return client_id, portal_id, photographer_id, ordertype_id, zipcode_no
            else:
                print("API Status Error:", responseData['APIStatusMessage'])
                client_id, portal_id, photographer_id, ordertype_id, zipcode_no = 0
                return client_id, portal_id, photographer_id, ordertype_id, zipcode_no
        else:
            print("Status Code: ", response.status_code)
    except Exception as e:
        print("An error occurred:", e)

def fetch_createorder_api(CreateOrderAPI, PostRequestedBody):
    try:
        response = requests.post(CreateOrderAPI, json=PostRequestedBody)

        if response.ok == True:
            responseData = response.json()
            apiStatus = responseData['apiStatus']

            if apiStatus == 0:
                result = json.loads(responseData['result'])
                APIstatusmsg = result['APIStatusMessage']
            
                return APIstatusmsg
            else:
                result = json.loads(responseData['result'])
                APIstatusmsg = result['APIStatusMessage']
                
                return APIstatusmsg
        else:
            # Assuming 'response' contains the JSON response
            response_data = json.loads(response.text)

            # Extract the 'result' field and parse it as JSON
            result_data = json.loads(response_data['result'])

            # Extract the 'APIStatusMessage' field from the parsed 'result' data
            api_status_message = result_data['APIStatusMessage']

            return api_status_message
    except Exception as e:
        print("An error occurred:", e)

def main():
    
    baseUrl = "http://13.200.17.36/Order_Count_Insight/api/OrderCountInsights" 
    # tokenget =  "eyJhbGciOiJodHRwOi8vd3d3LnczLm9yZy8yMDAxLzA0L3htbGRzaWctbW9yZSNobWFjLXNoYTI1NiIsInR5cCI6IkpXVCJ9.eyJqdGkiOiIxNjlkYzAxNy0wYWExLTQyZmItYjYxZS0wNzAyMzg0NjFiZWUiLCJFbXBsb3llZU5hbWUiOiJSYW1LdW1hciIsImV4cCI6MTcxMzQ1MDI4MiwiaXNzIjoiaHR0cDovLzEzLjIwMC4xNy4zNi8iLCJhdWQiOiJodHRwOi8vMTMuMjAwLjE3LjM2LyJ9.uF7aMoZPS_Cbp5sxv1xsnBtfrpFPCYqfXPyhaoYbRaQ"
    GetRequiredValuesAPI = f"{baseUrl}/GetRequiredValues"

    CreateOrderAPI = f"{baseUrl}/CreateOrder"

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

                clientname = ordervalues.get("Client")
                clientname = clientname.rstrip(' ')

                subjectaddress = ordervalues.get("Address")
                subjectaddress = subjectaddress.rstrip(' ')

                orderno = ordervalues.get("OrderNo")
                orderno = orderno.rstrip(' ')

                ordertype = ordervalues.get("OrderType")
                ordertype = ordertype.rstrip(' ')

                duedatetime = ordervalues.get("OrderDue")
                duedatetime = duedatetime.rstrip(' ')
                if duedatetime != "":
                    # Check if duedatetime is in the format "06/28/2018 03:25 AM"
                    if "-" in duedatetime:
                        # Convert the format to "06/28/2018 03:25 AM"
                        duedatetime = datetime.strptime(duedatetime, "%m-%d-%Y %I:%M %p")

                else:
                    # Get the current date and add 3 days
                    current_datetime = datetime.now()
                    new_due_date = current_datetime + timedelta(days=3)
                    # Format the new due date as desired
                    duedatetime = new_due_date.strftime("%m/%d/%Y %I:%M %p") #
                    # Parse the formatted due date back into a datetime object
                    formatted_due_date = datetime.strptime(duedatetime, "%m/%d/%Y %I:%M %p")

                timezone = ordervalues.get("TimeZone")
                timezone = timezone.rstrip(' ')
                if timezone == "":
                    timezone = "EST"

                # Define time differences for each timezone
                time_diff = {
                    "EST": timedelta(hours=9, minutes=30),
                    "PST": timedelta(hours=12, minutes=30),
                    "MST": timedelta(hours=11, minutes=30)
                }

                formatted_ist_time = formatted_due_date + time_diff.get(timezone)
                ist_time = formatted_ist_time.strftime("%m/%d/%Y %H:%M:%S")
               

                photographer = ordervalues.get("Photographer")
                photographer =  photographer.rstrip(' ')
                if photographer == "":
                    photographer = "Ecesis Photographer"

                # photographernotes = ""
                photographernotes = ordervalues.get("Notes")

                #portal = "Single Source" #convert to single letters and compare
                portal = ordervalues.get("Portal")
                portal = portal.rstrip(' ')

                # photographerfee = "10"
                photographerfee = ordervalues.get("PhotoFee")
                photographerfee = photographerfee.rstrip(' ')

                # ordervalue = "10"
                ordervalue = ordervalues.get("OrderValue")
                ordervalue = ordervalue.rstrip(' ')

                    # remarks = ""
                remarks = ordervalues.get("Remarks")

                # Format order data for API
                SelectrequestBody = {
                    'clientName': clientname,
                    'subjectAddress': subjectaddress,
                    'orderType': ordertype,
                    'photographerName': photographer,
                    'portalName': portal,
                    # Add other fields as needed
                }

                # start_time = time.time()
                start_time = time.time()
                client_id, portal_id, photographer_id, ordertype_id, zipcode_no = fetch_required_values(GetRequiredValuesAPI, SelectrequestBody)

                isfrombts = 0
                employeeid = 328
                isbulk =0

                PostRequestedBody = {
                    "clientId": client_id,
                    "procBulkData": f"{isfrombts}~{employeeid}~{orderno}~{duedatetime}~{subjectaddress}~{ordertype_id}~{portal_id}~{photographer_id}~{client_id}~{ordervalue}~{zipcode_no}~{ist_time}~{isbulk}~{photographernotes}~"
                }

                if client_id == 0:
                    message = "Values not Found"
                    
                elif client_id < 0:
                    message = "Client Not Found"
                    
                elif portal_id < 0:
                    message = "Portal Not Found"
                    
                elif photographer_id < 0:
                    message = "Photographer Not Found"
                    
                elif ordertype_id < 0:
                    message = "OrderType Not Found"
                    
                elif zipcode_no == "-5":
                    message = "Zipcode Not Found"
                    
                else:
                    message = fetch_createorder_api(CreateOrderAPI, PostRequestedBody)

                end_time = time.time()
                order_sheet.update_cell(row_index, len(ordervalues) - 2, message)

                start_datetime = datetime.fromtimestamp(start_time)
                end_datetime = datetime.fromtimestamp(end_time)

                # Convert datetime object to a string format
                start_ist_str = start_datetime.strftime("%Y-%m-%d %H:%M:%S")
                end_ist_str = end_datetime.strftime("%Y-%m-%d %H:%M:%S")


                order_sheet.update_cell(row_index, len(ordervalues) - 1, start_ist_str)
                order_sheet.update_cell(row_index, len(ordervalues), end_ist_str)

        else:
           # If no new rows are found, sleep for 10 seconds before checking again
           time.sleep(10)

if __name__ == "__main__":
   main()
   