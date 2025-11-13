import requests
from msal import PublicClientApplication
import language_tool_python
from docx import Document
import re
import docx2txt

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

CLIENT_ID = "92196e36-4333-451f-873e-f6da4df63081"
TENANT_ID = "70de1992-07c6-480f-a318-a1afcba03983"
AUTHORITY = f"https://login.microsoftonline.com/consumers"
SCOPES = ["Files.ReadWrite.All"]

#Creates MSAL app
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
            print("Device flow failed. Exiting.")
        else:
            print(flow["message"])
            result = app.acquire_token_by_device_flow(flow)
except Exception:
    print("An error occurred during authentication. Exiting.")
    result = None

if not result or "access_token" not in result:
    print("Authentication failed. Exiting.")
    result = None
else:
    accessToken = result['access_token']
    headers = {'Authorization': f'Bearer {accessToken}'}

    #Get most recent Word doc
    response = requests.get(
    "https://graph.microsoft.com/v1.0/me/drive/recent",
    headers=headers
)

if not response.ok:
    print("Error retrieving recent files:", response.status_code, response.text)
    exit()

recentFiles = response.json()

#Filter for .docx
wordFiles = [f for f in recentFiles.get("value", []) if f["name"].lower().endswith(".docx")]

if not wordFiles:
    print("No Word files found in recent documents.")
    exit()

#Pick the most recent Word doc
docItem = wordFiles[0]
print(f"Working on most recent document: {docItem['name']}")
# Download the file
metadataUrl = f"https://graph.microsoft.com/v1.0/me/drive/items/{docItem['id']}"
metadataResponse = requests.get(metadataUrl, headers=headers)

if not metadataResponse.ok:
    print("Error fetching metadata. Exiting.")
    exit()

metadata = metadataResponse.json()
downloadUrl = metadata["@microsoft.graph.downloadUrl"]

fileResponse = requests.get(downloadUrl)
filename = docItem["name"]
if not fileResponse.ok:
    print(f"Download failed.")
    exit()
with open(filename, "wb") as f:
    f.write(fileResponse.content)
print(f"File download success.\n")
#Read document
doc = Document(filename)
#Extract text from document and test
text = docx2txt.process(filename)
print(text)

#Actual modification of the file the grammar check
words = tool.check(text)
print(f"{len(words)} are mispelled.")
for mistake in words:
    if mistake.replacements:
        incorrectWord=mistake.context[mistake.offset:mistake.offset + mistake.errorLength]
        suggestion=mistake.replacements[0]
        dist=levenshteinCorrection(incorrectWord, suggestion)
        #For testing
        print(f"Incorrect word: {incorrectWord}, Suggestion: {suggestion}, Distance: {dist}")
        if dist <= 3: #Only highlight and replace if its likely to be a typo due to its small distance
            for paragraph in doc.paragraphs:
                if incorrectWord in paragraph.text:
                    for run in paragraph.runs:
                        if incorrectWord in run.text:
                            run.font.highlight_color = 6
                            doc.add_paragraph(f"Suggestion for'{incorrectWord}': '{suggestion}'")
                                 #   run.text = run.text.replace(incorrectWord, suggestion
#Save modified doc
altFilename = f"corrected {filename}"
doc.save(altFilename)
#Upload modified file back to OneDrive
uploadUrl = f"https://graph.microsoft.com/v1.0/me/drive/items/{docItem['id']}/content"
with open(filename, "rb") as f:
    uploadResponse = requests.put(uploadUrl, headers=headers, data=f)

if uploadResponse.status_code in [200, 201]:
        print(f"Uploaded file successfully.")
else:
    print("Upload failed.")
