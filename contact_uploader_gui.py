import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import threading
import re
spinner_frames = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
spinner_running = False
spinner_index = 0

SCOPES = ['https://www.googleapis.com/auth/contacts']

def get_google_service():
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    service = build('people', 'v1', credentials=creds)
    return service

def upload_contacts(file_path, area_name): 
    df = pd.read_excel(file_path)
    service = get_google_service()
    total = len(df)
    progress_bar['maximum'] = total
    for index, row in df.iterrows():
        raw_number = str(row['Phone No'])
        cleaned_number = re.sub(r'\D', '', raw_number)
        if len(cleaned_number) == 10:
            phone = cleaned_number
        else:
            continue    
        name = f"{area_name} - {phone}"
        contact_body = { "names": [{"givenName": name}], "phoneNumbers": [{"value": phone}]}
        try:
            service.people().createContact(body=contact_body).execute()
        except Exception as e:
            print(f"Failed to upload contact {name}: {e}")
        progress_bar['value'] = index + 1
        status_label.config(text = f"Uploading {index + 1} of {total} Contacts")
        root.update_idletasks()
    status_label.config(text = "Upload Complete!")
    progress_bar['value'] = 0

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def trigger_upload():
    upload_thread = threading.Thread(target = run_upload)
    upload_thread.start()

def run_upload():
    file_path = file_entry.get()
    area_name = area_entry.get()
    if not file_path or not area_name:
        messagebox.showerror("ERROR, Plaese select an Excel file and enter an area name!")
        return
    try:
        upload_contacts(file_path, area_name)
        messagebox.showinfo("Success", "Contacts uploaded successfully!")
    except Exception as e:
        messagebox.showerror("Upload failed!", str(e)) 

def trigger_deletion():
    deletion_thread = threading.Thread(target=delete_contact_by_area, args=(area_entry.get(),))
    deletion_thread.start()

def animate_spinner():
    global spinner_running, spinner_index
    if spinner_running:
        frame = spinner_frames[spinner_index % len(spinner_frames)]
        status_label.config(text = f"Deleting contacts {frame}")
        spinner_index += 1
        root.after(100, animate_spinner)

def delete_contact_by_area(area_name):
    global spinner_running, spinner_index
    spinner_running = True
    spinner_index = 0
    animate_spinner()
    try:
        service = get_google_service()
        connections = service.people().connections().list(resourceName='people/me', pageSize=1000, personFields='names').execute()
        deleted_count = 0
        for person in connections.get('connections', []):
            names = person.get('names', [])
            if names:
                name = names[0].get("givenName", "")
                if name.startswith(area_name + " -"):
                    try:
                        service.people().deleteContact(resourceName=person["resourceName"]).execute()
                        deleted_count += 1
                    except Exception as e:
                        print(f"Failed to delete contact {name}: {e}")
        spinner_running = False
        status_label.config(text = "Deletion Complete!")
        messagebox.showinfo("Success", f"Deleted {deleted_count} contacts!")
    except Exception as e:
        spinner_running = False
        status_label.config(text = "Deletion Failed!")
        messagebox.showerror("Error", f"Failed to delete contacts: {e}")
    

root = tk.Tk()
root.title('Contact Uploader')
tk.Label(root, text="Excel file:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
file_entry = tk.Entry(root, width = 40)
file_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Area Name:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
area_entry = tk.Entry(root, width = 40)
area_entry.grid(row=1, column=1, padx=5, pady=5)
tk.Label(root, text=" ").grid(row=1, column=2)

tk.Button(root, text="Upload Contacts", command=trigger_upload).grid(row=2, column=1, pady=10)   
tk.Button(root, text="Delete Contacts", command=trigger_deletion).grid(row=3, column=1, pady=5)

status_label = tk.Label(root, text = "", width = 30, anchor='center')
status_label.grid(row=4, column=1, padx = 5, pady = 5)
progress_bar = ttk.Progressbar(root, orient = 'horizontal', length = 300, mode = 'determinate')
progress_bar.grid(row = 5, column = 1, pady = 10)
progress_bar['value'] = 0

root.mainloop()