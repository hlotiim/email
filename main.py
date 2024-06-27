import imaplib
import email
import re
import pandas as pd
import time
from email.header import decode_header
from email.utils import parseaddr
import os

# Email account credentials
username = 'EMAIL_HERE'
password = 'PASSWORD_HERE'

# IMAP server details
imap_server = 'imap.gmail.com'
imap_port = 993

# List of top email providers
top_providers = [
    '@gmail.com', '@outlook.com', '@yahoo.com', '@hotmail.com', '@aol.com', 
    '@live.com', '@msn.com', '@icloud.com', '@mail.com', '@zoho.com'
]

# Connect to the server
mail = imaplib.IMAP4_SSL(imap_server, imap_port)
try:
    mail.login(username, password)
    print("Login successful")
except mail.error as e:
    print(f"Login failed: {e}")
    exit(1)

# Select "All Mail" label
mail.select('"[Gmail]/All Mail"')
print("All Mail selected")

# Function to extract phone numbers from email body
def extract_phone_numbers(text):
    phone_pattern = re.compile(
        r'\+?\d[\d\s.-]{7,}\d'
    )
    matches = phone_pattern.findall(text)
    phone_numbers = [match.strip() for match in matches if len(match) >= 10]  # Filter out short numbers
    return phone_numbers

# Function to check if email is from a top provider
def is_from_top_provider(email_address):
    return any(email_address.endswith(domain) for domain in top_providers)

# Lists to store extracted data
senders = []
full_names = []
subjects = []
phone_numbers = []

# Function to decode email parts safely
def decode_payload(part):
    try:
        return part.get_payload(decode=True).decode('utf-8')
    except UnicodeDecodeError:
        try:
            return part.get_payload(decode=True).decode('latin1')
        except UnicodeDecodeError:
            return part.get_payload(decode=True).decode('utf-8', errors='ignore')

# Function to save data to Excel
def save_to_excel():
    df = pd.DataFrame({
        'Full Name': full_names,
        'Email Received From': senders,
        'Subject': subjects,
        'Phone Number': phone_numbers
    })
    temp_file = 'emails_temp.xlsx'
    final_file = 'emails.xlsx'
    df.to_excel(temp_file, index=False)
    if os.path.exists(final_file):
        os.remove(final_file)
    os.rename(temp_file, final_file)

# Function to fetch emails using UID
def fetch_emails_using_uid(mail, batch_size=1000):
    result, data = mail.uid('search', None, 'ALL')
    email_uids = data[0].split()
    total_emails = len(email_uids)
    print(f"Found {total_emails} emails")

    start_time = time.time()
    processed_count = 0

    for start in range(0, total_emails, batch_size):
        end = min(start + batch_size, total_emails)
        batch_uids = email_uids[start:end]
        batch_uids_str = ','.join([uid.decode('utf-8') for uid in batch_uids])  # Decode bytes to strings
        result, data = mail.uid('fetch', batch_uids_str, '(RFC822)')
        
        for idx, response_part in enumerate(data):
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                # Decode the email subject
                subject, encoding = decode_header(msg['Subject'])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding or 'utf-8')
                # Get the email sender's name and email address
                from_ = msg.get('From')
                if isinstance(from_, email.header.Header):
                    from_ = decode_header(from_)[0][0]
                if isinstance(from_, bytes):
                    from_ = from_.decode(encoding or 'utf-8')
                full_name, email_address = parseaddr(from_)
                # Check if the email is from a top provider
                if not is_from_top_provider(email_address):
                    continue
                # Extract the email body
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get('Content-Disposition'))
                        if content_type == 'text/plain' and 'attachment' not in content_disposition:
                            body = decode_payload(part)
                else:
                    body = decode_payload(msg)
                phone_nums = extract_phone_numbers(body)
                # Only append data if phone numbers are found
                if phone_nums:
                    senders.append(email_address)
                    full_names.append(full_name)
                    subjects.append(subject)
                    phone_numbers.append(', '.join(phone_nums))
                    processed_count += 1

                    if processed_count % 10 == 0:
                        elapsed_time = time.time() - start_time
                        avg_time_per_email = elapsed_time / processed_count
                        remaining_emails = total_emails - processed_count
                        remaining_time = remaining_emails * avg_time_per_email
                        remaining_minutes, remaining_seconds = divmod(remaining_time, 60)
                        print(f"Processed {processed_count}/{total_emails} emails")
                        print(f"Elapsed time: {elapsed_time:.2f} seconds")
                        print(f"Estimated remaining time: {int(remaining_minutes)} minutes {int(remaining_seconds)} seconds")

                        # Save to Excel
                        save_to_excel()
                        print(f"Saved to emails.xlsx   {processed_count}/{total_emails}")

# Fetch and process emails using UID
fetch_emails_using_uid(mail)

# Logout and close connection
mail.logout()
print("Logged out and connection closed")
