# Contact Uploader Tool

This is a desktop tool that helps upload a large number of phone numbers from an Excel file to Google Contacts. It can also delete previously uploaded contacts based on area name.

## What It Does

- Reads phone numbers from an Excel (.xlsx) file.
- Uploads them to your Google Contacts.
- Saves each contact in the format: AreaName - PhoneNumber.
- Can delete previously uploaded contacts by matching area name.

## Example Use Case

You want to create a WhatsApp group with many people, but you don’t want to manually save all their numbers. This tool helps:

1. Upload all the numbers to Google Contacts with a label (area name).
2. WhatsApp will sync the contacts.
3. You can easily add them to a WhatsApp group.
4. After use, you can delete them with one click.

## Excel File Format

- File type: `.xlsx` (Excel)
- The column containing phone numbers should be named exactly: `Phone No`
- The phone number values can be in any format. The tool will clean and format them automatically.

## How to Run

1. Clone or download this repository.
2. Open the folder in your system.
3. Create a file named `credentials.json` (from your Google Cloud Console).
4. Place the `credentials.json` file in the same folder as the Python script.
5. Run the script: `python main.py`
6. A window will open where you can:
   - Select your Excel file.
   - Enter the area name.
   - Upload or delete contacts.

## Google Setup (One-time)

To generate the `credentials.json`:

1. Go to https://console.cloud.google.com/
2. Create a new project.
3. Enable the "People API" for the project.
4. Go to "Credentials", and click "Create OAuth client ID".
5. Choose "Desktop App" and download the file.
6. Rename it to `credentials.json` and place it in the project folder.
