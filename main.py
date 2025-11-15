import requests
from msal import PublicClientApplication
import language_tool_python
from docx import Document
import subprocess
import docx2txt
import os

CLIENT_ID = "92196e36-4333-451f-873e-f6da4df63081"
TENANT_ID = "70de1992-07c6-480f-a318-a1afcba03983"
AUTHORITY = f"https://login.microsoftonline.com/consumers"
SCOPES = ["Files.ReadWrite.All"]

#For detection
tool = language_tool_python.LanguageTool('en-US')


#Creates MSAL app
app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
headers = None

def authenticate():
    """Authenticate user with Microsoft OAuth device flow."""
    global headers

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    else:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            return "Device flow failed."

        print(flow["message"])  # Will also show in terminal
        result = app.acquire_token_by_device_flow(flow)

    if not result or "access_token" not in result:
        return "Authentication failed."

    headers = {'Authorization': f'Bearer {result["access_token"]}'}
    return "Authentication successful."

def get_latest_word_file():
    """Gets the most recent .docx file from OneDrive."""
    global headers

    url = "https://graph.microsoft.com/v1.0/me/drive/root/children?$orderby=lastModifiedDateTime desc"
    response = requests.get(url, headers=headers)

    if not response.ok:
        return None, "Failed to retrieve recent files."

    files = response.json().get("value", [])
    word_files = [f for f in files if f["name"].lower().endswith(".docx")]

    if not word_files:
        return None, "No .docx files found."

    return word_files[0], "Found latest Word document."

def download_file(docItem):
    """Downloads the Word document."""
    metadataUrl = f"https://graph.microsoft.com/v1.0/me/drive/items/{docItem['id']}"
    metadata = requests.get(metadataUrl, headers=headers).json()
    downloadUrl = metadata["@microsoft.graph.downloadUrl"]

    response = requests.get(downloadUrl)

    filename = docItem["name"]
    with open(filename, "wb") as f:
        f.write(response.content)

    return filename, "Download complete."

def correct_document(filename):
    """Correct grammar and return corrected file name."""
    doc = Document(filename)
    text = docx2txt.process(filename).strip()
    words = tool.check(text)

    for mistake in words:
        if mistake.ruleIssueType != 'misspelling':
            continue
        if not mistake.replacements:
            continue

        incorrect = mistake.context[mistake.offset:mistake.offset + mistake.errorLength]
        suggestion = mistake.replacements[0]

        if incorrect and suggestion:
            for paragraph in doc.paragraphs:
                if incorrect in paragraph.text:
                    for run in paragraph.runs:
                        run.text = run.text.replace(incorrect, suggestion)

    new_filename = f"corrected_{filename}"
    doc.save(new_filename)
    return new_filename, "Correction complete."

def upload_corrected(docItem, corrected_filename):
    """Uploads corrected file back to OneDrive."""
    uploadUrl = f"https://graph.microsoft.com/v1.0/me/drive/items/{docItem['id']}/content"

    with open(corrected_filename, "rb") as f:
        response = requests.put(uploadUrl, headers=headers, data=f)

    if response.status_code in [200, 201]:
        return "Upload successful."
    else:
        return "Upload failed."

def levenshteinCorrection(str1, str2):
    strlen1 = len(str1)
    strlen2 = len(str2)

    if strlen1 < strlen2: #Ensure str1 is the longer string for the algorithm
        temp = str1
        str1 = str2
        str2 = temp

    if len(str2) == 0:#If the string is empty then the distance is the length of the other string
        return len(str1)

    prevRow = range(len(str2) + 1) #The previous row is the distance between str1 and str2
    for i in range(strlen1): #Goes through each character in str1 using the sequence of characters
        char1 = str1[i] #To get the current character in str1
        currRow = [i + 1]
        for j in range(strlen2):
            char2 = str2[j] #To get the current character in str2
            inserts = prevRow[j + 1] + 1
            deletes = currRow[j] + 1
            if char1 == char2:
                subs = currRow[j]
            else:
                subs = currRow[j] + 1
            currRow.append(min(inserts, deletes, subs))#Ensures the smallest distance is taken
        prevRow = currRow #The current row becomes the previous row for the next iteration

    return currRow[-1] #Returns the last element which is the Levenshtein distance
