import requests

def get_outlook_emails(access_token, user_email, top=50):
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages?$top={top}"

    if last_run:
        last_run_str = last_run.isoformat()
        url += f"&$filter=receivedDateTime ge {last_run_str}"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)

    # Check for errors
    if response.status_code != 200:
        print("Error response:", response.content)  # Print error response content
    response.raise_for_status()
    return response.json() # Data is in JSON


def get_excel_data(access_token, file_id, sheet_name):
    """
    Reads data from a specific Excel sheet in OneDrive.
    """
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets/{sheet_name}/usedRange(valuesOnly=true)"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()
