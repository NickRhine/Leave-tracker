import requests
import os
from io import BytesIO
import pandas as pd

#Get github stashed secrets containing the necessary credentials
TENANT_ID = os.getenv("AZURE_TENANT_ID")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
SHAREPOINT_FILE_ID = os.getenv("SHAREPOINT_FILE_ID")
SHAREPOINT_SITE_ID = os.getenv("SHAREPOINT_SITE_ID")

#Local testing not workflow
# TENANT_ID = "14e5adcf-4446-4cee-bac9-e293492fa769"
# CLIENT_ID = "d1476c68-95dc-4ded-8725-9a0e31e75df5"
# CLIENT_SECRET = "mBL8Q~LNKQtRvLkjjHRFX4-kPXhnQooiMFdFrdwC"
# SHAREPOINT_FILE_ID = "012LJMUY6BHXDWVGWPI5DIT3YPOFVODUTI"
# SHAREPOINT_SITE_ID = "netorg7968809.sharepoint.com,d6ef5094-875f-47d7-93c4-43ae171a04ff,883a8121-0374-49f4-9476-2d3b9a1cb38a"

#Excel sheet url grabs all data
# url_b = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{SHAREPOINT_FILE_ID}/workbook/worksheets('Data')/range(address='B2:B10')"
# url_c = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{SHAREPOINT_FILE_ID}/workbook/worksheets('Data')/range(address='C2:10')"
url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{SHAREPOINT_FILE_ID}/workbook/worksheets('Data')/usedrange"


#Get the access token from azure to edit files
def get_access_token():
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }

    response = requests.post(token_url, headers=headers, data=data)
    response.raise_for_status()  # Raise an error for bad responses
    return response.json()["access_token"]

def get_excel_data(url,access_token):
    headers = {
        "Authorization": f"Bearer {access_token}"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()

    return response.json()["values"]


def update_cell_value(new_value, access_token):
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{SHAREPOINT_FILE_ID}/workbook/worksheets('Data')/range(address='D19')"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    data = {
        "values": [[new_value]]  # Update the value in cell D8
    }

    response = requests.patch(url, headers=headers, json=data)
    response.raise_for_status()
    print("Cell D8 updated successfully!")

#extract values from column B and C return them as lists
def get_columnB_columnC(data, b_values, c_values):

    for row in data:
        if len(row) > 1: #Check for value in column B
            b_values.append(row[1])
        if len(row) > 2: #Check for value in column C
            c_values.append(row[2])


    b_values = b_values[1:] #Remove header value
    c_values = c_values[1:] #Remove header value

    #Remove all empty string values from the list
    for i in range(len(b_values)):
        if b_values[i] == "":
            b_values = b_values[:i]
            break
    
    for i in range(len(c_values)):
        if c_values[i] == "":
            c_values = c_values[:i]
            break

    return b_values, c_values

def update_excel_column(access_token, modified_values, range_address):
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{SHAREPOINT_FILE_ID}/workbook/worksheets('Data')/range(address='{range_address}')"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    data = {
        "values": [[value] for value in modified_values]  # Format as a 2D array
    }
    
    response = requests.patch(url, headers=headers, json=data)
    response.raise_for_status()
    print("Excel file updated successfully!")

def updated_values(b_values, c_values):
    for i in range(len(b_values)):
        b_values[i] = b_values[i] + c_values[i]
    return b_values

# Main function to run the script
if __name__ == '__main__':
    token = get_access_token()
    excel_data = get_excel_data(url, token)
    b_values, c_values = [], []
    b_values, c_values = get_columnB_columnC(excel_data, b_values, c_values)
    
    updated_b_values = []
    updated_b_values = updated_values(b_values, c_values)

    address_range = f"B2:B{len(updated_b_values) + 1}"
    update_excel_column(token, updated_b_values, address_range)
    print("Done")