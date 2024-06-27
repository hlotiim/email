# Email Processing Bot
This Python script connects to a Gmail account using IMAP, fetches all emails, extracts specific information, and saves it to an Excel file. The script processes emails from top email providers, extracts phone numbers, and provides detailed progress updates during execution.

## Features
- Connects to a Gmail account using IMAP.
- Fetches emails from the "All Mail" label.
- Extracts sender's email, full name, subject, and phone numbers from the email body.
- Saves extracted data to an Excel file.
- Provides detailed progress updates, including the number of emails processed, elapsed time, and estimated remaining time.

## Detailed Steps
Connecting to Gmail using IMAP
Enable IMAP in your Gmail account:

Go to your Gmail account.
Click on the gear icon in the upper right corner and select "See all settings".
Go to the "Forwarding and POP/IMAP" tab.
In the "IMAP access" section, select "Enable IMAP".
Click "Save Changes".
Get the App Password for Gmail:

Go to your Google Account.
Click on "Security" in the left-hand menu.
Under "Signing in to Google", select "App passwords".
You might need to sign in again.
At the bottom, click "Select app" and choose "Other (Custom name)".
Enter a name (e.g., "Python Email Bot") and click "Generate".
Copy the app password. You will use this instead of your regular password.

## Requirements
- Python 3.x
- `imaplib` (part of Python standard library)
- `email` (part of Python standard library)
- `re` (part of Python standard library)
- `pandas`
- `openpyxl`

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/hlotiim/email.git
   cd email

## Install the required Python packages:
```pip install pandas openpyxl

## Configuration
Update the following lines in main.py with your Gmail credentials:
# Email account credentials
username = 'EMAIL_HERE'
password = 'PASSWORD_HERE'
