# Correctly – AI-Powered OneDrive Grammar Correction Tool

Correctly is a Python application that automatically:

- Authenticates with Microsoft OneDrive  
- Downloads your most recent Word (`.docx`) document  
- Uses LanguageTool to correct grammar and spelling  
- Saves a corrected version locally  
- Uploads the corrected version back to OneDrive  
- Provides a simple Tkinter GUI to run everything with one click  

---

# Installation & Setup Guide

You can run this project on **Ubuntu (WSL)** or **Windows**.  
Instructions for both are included below.

---

    # Option 1 — Windows Setup (Recommended)

        ### 1. Install Python 3.10+ for Windows
        Download from:
        https://www.python.org/downloads/

        Make sure to check:


        ---

        ### 2. Install Java (required for LanguageTool)

        Open PowerShell:

        ```powershell
        winget install EclipseAdoptium.Temurin.17.JRE

        java -version

        python -m venv .venv
        .\.venv\Scripts\activate

        pip install requests msal language-tool-python python-docx docx2txt lxml

        Run the UI : python ui.py

    # Option 2 - Ubuntu/WSL Setup
        1. Install WSL on Windows on Powershell as Administrator
            wsl --install (Restart PC after)

        2. Update Ubuntu
            sudo apt update && sudo apt upgrade -y

        3. Install Python and pip
            sudo apt install python3 python3-pip python3-venv -y

        4. Install Java 
            sudo apt install default-jre -y
            (verify installation : java -version)
        
        5. Create environment
            python3 -m venv env
            source env/bin/activate
    
        6. Install Python packages 
            pip install requests msal language-tool-python python-docx docx2txt lxml

        7. Run program 
            python3 ui.py





# How the app works 
    After running the program...

    Step #1 : Authentication 
        - Click the Authentication button and a message in the terminal will appear 
            "To sign in, use a web browser to open the page https://www.microsoft.com/link and enter the code XXXXXXXX"
        - Enter the Code, allow permissions, receive the "Authentication Successful."
    Step #2 : Correct Process 
        - The app finds your most recent .docx one Onedrive, downloads it, corrects it, and stores a saved version locally as "corrected_<filename>.docx", uploads back to Onedrive, and displays where the file was saved, provides buttons to open the file or folder 