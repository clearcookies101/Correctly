import requests
from msal import PublicClientApplication
import language_tool_python
from docx import Document
import re
import docx2txt

CLIENT_ID = "92196e36-4333-451f-873e-f6da4df63081"
TENANT_ID = "70de1992-07c6-480f-a318-a1afcba03983"
AUTHORITY = f"https://login.microsoftonline.com/consumers"
SCOPES = ["Files.ReadWrite.All"]

#Creates the MSAL app
app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
headers = None

#For detection
tool = language_tool_python.LanguageTool('en-US')

#Acquires token
result = None
try:
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    else:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            print("Can't start device flow. Check your client ID and tenant ID.")
        else:
            print(flow["message"])
            result = app.acquire_token_by_device_flow(flow)
except Exception as e:
    print("An error occurred during authentication:", str(e))
    result = None

if not result or "access_token" not in result:
    print("Authentication failed. Exiting.")
    result = None
else:
    accessToken = result['access_token']
    headers = {'Authorization': f'Bearer {accessToken}'}

if headers:
    recentFiles = requests.get("https://graph.microsoft.com/v1.0/me/drive/recent", headers=headers).json()
    #Get most recent Word document
    if "value" in recentFiles and recentFiles["value"]:
        if "value" in recentFiles and recentFiles["value"]:
            docItem = recentFiles["value"][0]
            WORD_FILE_PATH = docItem["name"]
            print(f"Working on most recent document: {WORD_FILE_PATH}")

            #Download file
            metadataUrl = f"https://graph.microsoft.com/v1.0/me/drive/items/{docItem['id']}"
            metadata = requests.get(metadataUrl, headers=headers).json()
            downloadUrl = metadata["@microsoft.graph.downloadUrl"]

            filename = WORD_FILE_PATH
            response = requests.get(downloadUrl)
            with open(filename, "wb") as f:
                f.write(response.content)

            print(f"Downloaded {filename} successfully.\n")

            #Read document
            doc = Document(filename)
            #text = ""
            allText = []

            #Extract text from document and test
            text = docx2txt.process(filename)
            print(text)
                
            #Actual modification of the file the grammar check
            incorrectWords = tool.check(text)
            print(f"{len(incorrectWords)} are mispelled.")
            for mistake in incorrectWords:
                print(f"Incorrect word: {mistake.context[mistake.offset:mistake.offset + mistake.errorLength]}")
                print(f"Suggestions: {mistake.replacements}")
                text = text.replace(mistake.context[mistake.offset:mistake.offset + mistake.errorLength], 
                                    mistake.replacements[0] if mistake.replacements else "")
            #Upload modified file back to OneDrive
            uploadUrl = f"https://graph.microsoft.com/v1.0/me/drive/items/{docItem['id']}/content"
            with open(filename, "rb") as f:
                uploadResponse = requests.put(uploadUrl, headers=headers, data=f)

            if uploadResponse.status_code in (200, 201):
                print(f"Uploaded {filename} successfully.")
            else:
                print("Upload failed:", uploadResponse.text)
        else:
            print("No recent Word files found.")
    else:
        print("No headers available for requests.")
