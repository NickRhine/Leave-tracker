import requests
import os
import numpy as np
from io import BytesIO
import pandas as pd

TENANT_ID = os.getenv("AZURE_TENANT_ID")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
SHAREPOINT_FILE_ID = os.getenv("SHAREPOINT_FILE_ID")
SHAREPOINT_SITE_ID = os.getenv("SHAREPOINT_SITE_ID")

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

def download_excel_file(access_token):
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drives/{DRIVE_ID}/root:/{FILE_PATH}:/content"
    headers = {"Authorization": f"Bearer {access_token}"}

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    return pd.read_excel(BytesIO(response.content))

# Modify Excel File 
def modify_excel(df):
    new_column_name = f"Month_{pd.Timestamp.now().strftime('%Y-%m')}"
    df[new_column_name] = 0.25  # Adding 0.25 to the new column
    return df

# Upload the modified Excel file back to SharePoint
def upload_excel_file(df, access_token):
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drives/{DRIVE_ID}/root:/{FILE_PATH}:/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }

    response = requests.put(url, headers=headers, data=output)
    response.raise_for_status()
    print("File updated successfully!")

if __name__ == '__main__':
    token = get_access_token()
    print(token)