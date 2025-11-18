from msal import PublicClientApplication
import language_tool_python
from docx import Document
import requests
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
    #Authenticate user with Microsoft OAuth device flow.
    global headers

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    else:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            return "Device flow failed."

        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    if not result or "access_token" not in result:
        return "Authentication failed."

    headers = {'Authorization': f'Bearer {result["access_token"]}'}
    return "Authentication successful."

def get_latest_word_file():
    #Gets the most recent .docx file from OneDrive.
    global headers

    #Get most recent Word doc
    url = "https://graph.microsoft.com/v1.0/me/drive/root/children?$orderby=lastModifiedDateTime desc"
    response = requests.get(url, headers=headers)

    if not response.ok:
        print("Retrieving recent files failed. Exiting.")
        return None, "Failed to retrieve recent files."

    files = response.json().get("value", [])
    word_files = [f for f in files if f["name"].lower().endswith(".docx")]

    if not word_files:
        print("No Word files found in recent documents.")
        return None, "No .docx files found."

    #Pick the most recent Word doc
    docItem = word_files[0]
    print(f"Working on most recent document: {docItem['name']}")
    return docItem, "Found latest Word document."

def download_file(docItem):
    # Download the Word document
    metadataUrl = f"https://graph.microsoft.com/v1.0/me/drive/items/{docItem['id']}"
    metadataResponse = requests.get(metadataUrl, headers=headers)
    requests.patch(metadataUrl, headers=headers)
    if not metadataResponse.ok:
        print("Error fetching metadata. Exiting.")
        return None, "Download failed."

    metadata = metadataResponse.json()
    downloadUrl = metadata["@microsoft.graph.downloadUrl"]

    fileResponse = requests.get(downloadUrl)
    filename = docItem["name"]
    if not fileResponse.ok:
        print(f"Download failed.")
        return None, "Download failed."

    with open(filename, "wb") as f:
        f.write(fileResponse.content)
    print(f"File download success.\n")
    return filename, "Download complete."

def correct_document(filename):
    #Actual modification of the file the grammar check
    doc = Document(filename)
    #Extract text from document and test
    text = docx2txt.process(filename).strip()
    print(text)

    #For detection
    words = tool.check(text)
    typos = [w for w in words if w.ruleIssueType == 'misspelling']
    print(f"{len(typos)} are mispelled.")

    for mistake in words:
        #Reseting the variables
        incorrectWord = None
        suggestion = None
        dist = None
        if mistake.ruleIssueType != 'misspelling':
            continue  # Only misspellings
        if not mistake.replacements:
            continue  # When no suggestions are available

        incorrectWord = mistake.context[mistake.offset:mistake.offset + mistake.errorLength].strip()
        suggestion = mistake.replacements[0].strip() if mistake.replacements else "None"

        if not incorrectWord or not suggestion or suggestion == "None":
            continue  # Skip if either is empty
        try:
            dist = levenshteinCorrection(incorrectWord, suggestion)
        except Exception:
            print(f"Levenshtein calculation failed. Skipping word {incorrectWord}.")
            continue
        #For testing
        if dist is not None:
            print(f"Incorrect word: {incorrectWord}, Suggestion: {suggestion}, Distance: {dist}")
        if dist <=3: #Only highlight and replace if its likely to be a typo due to its small distance
            for paragraph in doc.paragraphs:
                messageText = ''.join(run.text for run in paragraph.runs)
                if incorrectWord in paragraph.text:
                    messageText = messageText.replace(incorrectWord, suggestion)
                    for run in paragraph.runs:
                        run.text = ''
                        if incorrectWord in run.text:
                            run.font.highlight_color = 6
                            run.text = run.text.replace(incorrectWord, suggestion)

    #Save modified doc
    altFilename = f"Corrected {filename}"
    doc.save(altFilename)
    return altFilename, "Correction complete."

def upload_corrected(docItem, corrected_filename):
    #Upload modified file back to OneDrive
    uploadUrl = f"https://graph.microsoft.com/v1.0/me/drive/items/{docItem['id']}/content"

    with open(corrected_filename, "rb") as f:
        uploadResponse = requests.put(uploadUrl, headers=headers, data=f)

    if uploadResponse.status_code in [200, 201]:
        print(f"Uploaded file successfully.")
        return "Upload successful."
    else:
        print("Upload failed.")
        return "Upload failed."

def open_file_windows(filename):
    #Convert path to Windows format for WSL
    windowsPath = subprocess.check_output(["wslpath", "-w", os.path.abspath(filename)]).decode().strip()
    #Open the modified document
    subprocess.run(["explorer.exe", windowsPath])
    print(f"Saved the new document as {filename}.")
