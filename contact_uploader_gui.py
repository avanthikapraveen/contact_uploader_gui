import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/contacts']

def get_google_service():
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    service = build('people', 'v1', credentials=creds)
    return service

def upload_contacts(file_path, area_name):
    df = pd.read_excel(file_path)
    service = get_google_service()
    for index, row in df.iterrows():
        phone = str(row['Phone No']).strip()
        name = f"{area_name} - {phone}"
        contact_body = { "names": [{"givenName": name}], "phoneNumbers": [{"value": phone}]}
        service.people().createContact(body = contact_body).execute()

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def trigger_upload():
    file_path = file_entry.get()
    area_name = area_entry.get()
    if not file_path or not area_name:
        messagebox.showerror("ERROR, Plaese select an Excel file and enter an area name")
        return
    try:
        upload_contacts(file_path, area_name)
        messagebox.showinfo("Success", "Contacts uploaded successfully!")
    except Exception as e:
        messagebox.showerror("Upload failed!", str(e)) 

# def delete_contact_by_area(area_name):
#     service = get_google_service()
#     connections = service.people().connections().list(resourceName='people/me', pageSize=1000, personFields='names').execute()
#     deleted_count = 0
#     for person in connections.get('connections',[]):
#         names = person.get('names', [])


root = tk.Tk()
root.title('WhatsApp Contact Uploader')
tk.Label(root, text="Excel file:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
file_entry = tk.Entry(root, width = 40)
file_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Area Name:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
area_entry = tk.Entry(root, width = 40)
area_entry.grid(row=1, column=1, padx=5, pady=5)
tk.Label(root, text=" ").grid(row=1, column=2)

tk.Button(root, text="Upload Contacts", command=trigger_upload).grid(row=2, column=1, pady=10)   

root.mainloop()