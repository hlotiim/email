# Email Processing Bot

This Python script connects to a Gmail account using IMAP, fetches all emails, extracts specific information, and saves it to an Excel file. The script processes emails from top email providers, extracts phone numbers, and provides detailed progress updates during execution.

## Table of Contents
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Running the Script](#running-the-script)
- [Detailed Steps](#detailed-steps)
  - [Connecting to Gmail using IMAP](#connecting-to-gmail-using-imap)
  - [Getting the App Password for Gmail](#getting-the-app-password-for-gmail)
- [Contributing](#contributing)
- [License](#license)

## Features
- Connects to a Gmail account using IMAP.
- Fetches emails from the "All Mail" label.
- Extracts sender's email, full name, subject, and phone numbers from the email body.
- Saves extracted data to an Excel file.
- Provides detailed progress updates, including the number of emails processed, elapsed time, and estimated remaining time.

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
