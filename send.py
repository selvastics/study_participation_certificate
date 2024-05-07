from exchangelib import Credentials, Account, DELEGATE, Message, FileAttachment, HTMLBody, Configuration
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
import os
import time
import logging
import urllib3

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')

# Disable SSL verification (use only if necessary and understand the risks)
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

# Exchange server credentials and configuration
username = '[PLACEHOLDER]'
password = '[PLACEHOLDER]'
credentials = Credentials(username, password)
config = Configuration(server='exchange.wwu.de', credentials=credentials)  # Exchange server address

# Set up the Exchange account
account = Account(primary_smtp_address=username, config=config, autodiscover=False, access_type=DELEGATE)

# Directory containing your PDF files
pdf_directory = 'CERTIFICATE_WRITER(PUB)/certificates'

# Function to send email
def send_email(filename, recipient_email, account):
    try:
        # Email subject and body (HTML formatted)
        subject = "Teilnahme an der [PLACEHOLDER] und geleistet Versuchspersonstunden"
        body = HTMLBody(f"""
        <strong>Liebe:r Teilnehmer:in,</strong><br><br>
        vielen Dank für Ihre Teilnahme an unserer Studie zu [PLACEHOLDER]. Anbei erhalten Sie Ihren Nachweis über die geleistet Versuchspersonstunde. Sollten Sie Fragen haben, stehen wir Ihnen gerne zur Verfügung.<br><br>
        Beste Grüße,<br>
        [PLACEHOLDER]
        """)

        # Create a new email message
        email = Message(
            account=account,
            folder=account.sent,
            subject=subject,
            body=body,
            to_recipients=[recipient_email]
        )

        # Attach the PDF file
        file_path = os.path.join(pdf_directory, filename)
        with open(file_path, 'rb') as f:
            content = f.read()
        attachment = FileAttachment(name=filename, content=content)
        email.attach(attachment)

        # Send the email
        email.send_and_save()
        logging.info(f"Sent email to {recipient_email} with {filename}")
    except Exception as e:
        logging.error(f"Error sending email to {recipient_email} with {filename}: {e}")

# Main function to iterate over files and send emails
def main():
    logging.info("Starting email sending process")
    all_files = os.listdir(pdf_directory)
    sorted_files = sorted(all_files)  # Sort files alphabetically
    for filename in sorted_files:
        if filename.endswith("_certificate.pdf"):
            recipient_email = filename.split('_certificate.pdf')[0]
            send_email(filename, recipient_email, account)
            time.sleep(2)  # Delay to avoid rate limiting
    logging.info("All emails sent.")


if __name__ == "__main__":
    main()
