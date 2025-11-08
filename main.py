import requests
from msal import PublicClientApplication

CLIENT_ID = "92196e36-4333-451f-873e-f6da4df63081"
TENANT_ID = "70de1992-07c6-480f-a318-a1afcba03983"
AUTHORITY = f"https://login.microsoftonline.com/consumers"
SCOPES = ["Files.ReadWrite.All"]

#Creates the MSAL app
app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
headers = None
#Acquires token
result = None
accounts = app.get_accounts()
try:
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    else:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            print("Can't start device flow. Check your client ID and tenant ID.")
            flow={}
        else:
            print(flow["message"])
            result = app.acquire_token_by_device_flow(flow)
except Exception as e:
    print("An error occurred during authentication:", str(e))
    result = None
if not result or "access_token" not in result:
    print("Authentication failed. Exiting.")
    result = None
    #headers= None
   # exit(0)
else:
    access_token = result['access_token']
    headers = {'Authorization': f'Bearer {access_token}'}

#Get most recent Word document
if headers:
    recent_files = requests.get("https://graph.microsoft.com/v1.0/me/drive/recent", headers=headers).json()
    if "value" in recent_files and recent_files["value"]:
        doc = recent_files["value"][0]
        WORD_FILE_PATH = doc["name"]
        print(f"Working on most recent document: {WORD_FILE_PATH}")
    else:
        raise Exception("No recent Word files found.")
        doc = None
    if headers:
    #Download file
        recent_files = requests.get("https://graph.microsoft.com/v1.0/me/drive/recent", headers=headers).json()
        if "value" in recent_files and recent_files["value"]:
            doc = recent_files["value"][0]
            WORD_FILE_PATH = doc["name"]
        metadata_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{doc['id']}"
        metadata = requests.get(metadata_url, headers=headers).json()
        download_url = metadata["@microsoft.graph.downloadUrl"]

        local_filename = WORD_FILE_PATH
        response = requests.get(download_url)
        with open(local_filename, "wb") as f:
            f.write(response.content)

        print(f"Downloaded {local_filename} successfully.\n")

        # Upload modified file
        upload_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{doc['id']}/content"
        with open(local_filename, "rb") as f:
            upload_response = requests.put(upload_url, headers=headers, data=f)

        if upload_response.status_code in (200, 201):
            print(f"Uploaded {local_filename} successfully.")
        else:
            print("Upload failed:", upload_response.text)
    else:
     print("No document to process.")
else:
    print("No headers available for requests.")
