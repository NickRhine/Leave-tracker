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

#Download the excel file from sharepoint to edit it
def download_excel_file(access_token):
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drives/items/{SHAREPOINT_FILE_ID}/workbook/worksheets('Data')/range(address='D19')" #usedRange
    headers = {"Authorization": f"Bearer {access_token}"}

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    return pd.read_excel(BytesIO(response.content))

# Modify Excel File by adding "Added leave" column to the "Total available leave" column
def modify_excel(df):
    new_column_name = f"Month_{pd.Timestamp.now().strftime('%Y-%m')}"
    df[new_column_name] = 0.25  # Adding 0.25 to the new column
    return df

# Upload the modified Excel file back to SharePoint
def upload_excel_file(df, access_token):
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drives/items/{SHAREPOINT_FILE_ID}/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }

    response = requests.put(url, headers=headers, data=output)
    response.raise_for_status()
    print("File updated successfully!")


def get_cell_value(access_token):
    # URL to access cell D8 in the 'Data' worksheet
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive/items/{SHAREPOINT_FILE_ID}/workbook/worksheets('Data')/range(address='D19')"
    headers = {"Authorization": f"Bearer {access_token}"}
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    # Get the current value of cell D8
    current_value = response.json()["values"][0][0]  # The cell's value is stored in this format
    return current_value

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

if __name__ == '__main__':
    token = get_access_token()
    current_value = get_cell_value(token)
    new_value = current_value + 1 # test 
    update_cell_value(new_value, token)
    print("Done")