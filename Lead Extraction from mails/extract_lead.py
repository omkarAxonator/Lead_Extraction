import pandas as pd
import imaplib
import email
from openpyxl import Workbook
from config import mail_config
import re

# Function to extract emails
def extract_emails(email_address, password, filterquery):
    # Connect to IMAP server
    mail = imaplib.IMAP4_SSL('pop.gmail.com')
    mail.login(email_address, password)
    mail.select('inbox')

    # Search for emails with specific subject
    result, data = mail.search(None,filterquery)
    
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


def extract_raw_emails():
    rawEmails = extract_emails(mail_config.GMAIL_USER, mail_config.GMAIL_PASSWORD, mail_config.FILTERQUERY)
    raw_output_file = f"RAW_{mail_config.GMAIL_USER}_{mail_config.SUBJECT}_{mail_config.startDate}_{mail_config.endDate}.xlsx"
    save_to_excel(rawEmails, raw_output_file)
    return raw_output_file

def remove_internal_emails(filename):
    df = pd.read_excel(filename)
    output_filename = f'filter1_{filename}'

    # # Drop rows where any cell contains 'Aarohi Kulkarni'
    df = df[~df.apply(lambda row: row.astype(str).str.contains('Aarohi Kulkarni').any(), axis=1)]
    df = df[~df.apply(lambda row: row.astype(str).str.lower().str.contains('aarohi kulkarni').any(), axis=1)]
    df = df[~df.apply(lambda row: row.astype(str).str.lower().str.contains('@axonator.com').any(), axis=1)]
    df = df[~df.apply(lambda row: row.astype(str).str.lower().str.contains('white snow').any(), axis=1)]
    df = df[~df.apply(lambda row: row.astype(str).str.lower().str.contains('whitesnow').any(), axis=1)]
    df = df[~df.apply(lambda row: row.astype(str).str.contains('Whitesnow').any(), axis=1)]
    df = df[~df.apply(lambda row: row.astype(str).str.contains('Jayesh Kitukale').any(), axis=1)]
    df['Body'] = df['Body'].str.replace(r'\n', '<br/>')
    # Apply replacements to each row in the 'Body' column
    df['Body'] = df['Body'].str.replace(r'<span>', '<br/>')


    # Save the result back to an Excel file
    df.to_excel(output_filename, index=False)
    print("\nFile Saved:")
    return output_filename

def format_excel(filter1_filename):
    df = pd.read_excel(filter1_filename)

    # Check if 'Body' column exists
    if 'Body' not in df.columns:
        raise KeyError("The 'Body' column does not exist in the DataFrame.")

    # Function to extract the required fields from the Body text
    def extract_info(Body):
        Body = Body.replace('\n', '<br/>')  # Replace newlines with spaces
        Body = Body.replace('<span>', '<br/>')  # Replace newlines with spaces
        user_name = re.search(r'User Name:\s*(.*?)(<br/>|$)', Body)
        email = re.search(r'Email:\s*(.*?)(<br/>|$)', Body)
        organization = re.search(r'Organization:\s*(.*?)(<br/>|$)', Body)
        industry = re.search(r'Industry:\s*(.*?)(<br/>|$)', Body)
        lead_looking_for = re.search(r'Lead is looking for:\s*(.*?)(<br/>|$)', Body)

        return {
            'User Name': user_name.group(1).strip() if user_name else None,
            'Email': email.group(1).strip() if email else None,
            'Organization': organization.group(1).strip() if organization else None,
            'Industry': industry.group(1).strip() if industry else None,
            'Lead is looking for': lead_looking_for.group(1).strip() if lead_looking_for else None
        }

    # Apply replacements to each row in the 'Body' column
    df['Body'] = df['Body'].str.replace(r'\n', '<br/>')
    # Apply replacements to each row in the 'Body' column
    df['Body'] = df['Body'].str.replace(r'<span>', '<br/>')
    extracted_data = df['Body'].apply(lambda x: extract_info(x) if isinstance(x, str) else None)

    # Create a new DataFrame from the extracted data
    new_df = pd.DataFrame(extracted_data.tolist())

    # Save the new DataFrame to a new Excel file
    final_outputfile = f'final_{filter1_filename}'
    new_df.to_excel(final_outputfile, index=False)

    print(f"Extracted data has been saved to {final_outputfile}")
    return final_outputfile


# Main function
def main():
    rawFileName = extract_raw_emails()
    filter1fileName = remove_internal_emails(rawFileName)
    format_excel(filter1fileName)

if __name__ == "__main__":
    main()









