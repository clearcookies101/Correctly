import tkinter as tk
from tkinter import scrolledtext
from main import authenticate, get_latest_word_file, download_file, correct_document, upload_corrected, open_file_windows
import os

corrected_file_path = None

def log(message):
    output.insert(tk.END, message + "\n")
    output.see(tk.END)

def do_authenticate():
    result = authenticate()
    log(result)

def do_start():
    docItem, message = get_latest_word_file()
    log(message)
    if not docItem:
        return

    original = docItem['name']
    log(f"Your original Onedrive file is : {original}")

    #Downloaded file 
    filename, message = download_file(docItem)
    log(message)
    log(f"Downloaded file is saved on: {filename}")

    # Correct the document
    corrected_filename, message = correct_document(filename)
    log(message)

    global corrected_file_path
    corrected_file_path = os.path.abspath(corrected_filename)

    log(f"Corrected file is saved on: {corrected_filename}")
    log(f"Corrected filepath is: {corrected_file_path}")

    message = upload_corrected(docItem, corrected_filename)
    log(message)

    log("Process completed.")

def open_corrected_file():
    if corrected_file_path and os.path.exists(corrected_file_path):
        open_file_windows(corrected_file_path)
        log("Opened corrected file.")
    else:
        log("Corrected file not found. Please run the correction process first.")
root = tk.Tk()
root.title("OneDrive Word Document via Correctly!")

tk.Label(root, text="Correctly - Grammar Checker", font=("Arial", 16 )).pack()

btn_authenticate = tk.Button(root, text="Authenticate with Microsoft", command=do_authenticate)
btn_authenticate.pack(pady=5)

btn_start = tk.Button(root, text="Start Correction Process", command=do_start)
btn_start.pack(pady=5)

btn_open_corrected_file = tk.Button(root, text="Open Corrected File", command=open_corrected_file)
btn_open_corrected_file.pack(pady=5)

output = scrolledtext.ScrolledText(root, width=60, height=20)
output.pack(pady=10)

root.mainloop()
