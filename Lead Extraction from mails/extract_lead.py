import imaplib
import email
from openpyxl import Workbook
from datetime import datetime

# Function to extract emails
def extract_emails(email_address, password, subject_keyword):
    # Connect to IMAP server
    mail = imaplib.IMAP4_SSL('pop.gmail.com')
    mail.login(email_address, password)
    mail.select('inbox')

    # Search for emails with specific subject
    result, data = mail.search(None, f'(SUBJECT "{subject_keyword}")')
    
    emails = []
    if result == 'OK':
        for num in data[0].split():
            result, data = mail.fetch(num, '(RFC822)')
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            date = msg["Date"]
            subject = msg["Subject"]
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))

                    if "attachment" not in content_disposition:
                        body += str(part.get_payload(decode=True))
            else:
                body = str(msg.get_payload(decode=True))

            emails.append((date, subject, body))

    mail.close()
    mail.logout()

    return emails

# Function to save emails to Excel
def save_to_excel(emails, filename):
    wb = Workbook()
    ws = wb.active
    ws.append(["Date", "Subject", "Body"])

    for email in emails:
        ws.append(email)

    wb.save(filename)

# Main function
def main():
    email_address = "vinay@axonator.com"
    password = "88$XDaG$34kf2%L2"
    subject_keyword = "New Signup on Axonator"
    output_file = "New_Signup_on_Axonator.xlsx"

    emails = extract_emails(email_address, password, subject_keyword)
    save_to_excel(emails, output_file)

if __name__ == "__main__":
    main()
