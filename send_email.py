import os
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import base64
from google.auth.transport.requests import Request

# SCOPES for Gmail API
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Function to authenticate Gmail API
def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('gmail', 'v1', credentials=creds)
        return service
    except HttpError as error:
        print(f'An error occurred: {error}')
        return None

# Function to create the email message with attachment
def create_message_with_attachment(sender, to, subject, message_text, file_path):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    # Attach the message body
    message.attach(MIMEText(message_text, 'html'))  # Ensure it's 'html' type

    # Attach the brochure
    with open(file_path, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
        message.attach(part)

    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

# Function to send the email message
def send_message(service, user_id, message):
    try:
        message = service.users().messages().send(userId=user_id, body=message).execute()
        print(f"Message Id: {message['id']}")
        return message
    except HttpError as error:
        print(f'An error occurred: {error}')
        return None

# Function to send bulk emails
def send_bulk_emails():
    service = authenticate_gmail()

    # Load the list of sponsors from the CSV
    df = pd.read_csv('sponsors_list.csv')

    # Iterate through each sponsor in the CSV
    for index, row in df.iterrows():
        sponsor_name = row['Sponsor Name']
        email = row['Email']

        print(f"Sending email to: {sponsor_name} at {email}")

        # Prepare the email content
        subject = "Invitation to Collaborate: Sponsorship Opportunity for D3 Fest – The Techno-Management Fest of IIIT Bhubaneswar"
        message_text = f"""
        <html>
        <head></head>
        <body>
            <p>Dear {sponsor_name},</p>
            
            <p>We hope this email finds you well.</p>
            
            <p>We are reaching out on behalf of <strong>IIIT Bhubaneswar</strong> to invite you to collaborate with us for <strong>D3 Fest 2024</strong>, our premier Techno-Management Fest scheduled from <strong>8th to 10th November</strong>. This highly anticipated event showcases the spirit of innovation and leadership, uniting the brightest minds in technology and management from across the country.</p>
            
            <p><strong>D3 Fest</strong> is a unique convergence of cutting-edge technology and management expertise, featuring a wide array of exciting events such as:</p>
            
            <ul>
                <li>Craft-N-Code</li>
                <li>Build a Bot</li>
                <li>CyberSec Battle</li>
                <li>Pitch and Win</li>
                <li>Robo Race</li>
                <li>AutoBot</li>
                <li>DigiCast</li>
                <li>Beyond Boundaries</li>
                <li>E-Talk</li>
            </ul>

            <p>These events offer a fantastic opportunity for your brand to engage with highly talented and motivated students, creating meaningful connections with the next generation of innovators.</p>

            <p>We would greatly appreciate the chance to discuss how your organization can partner with us to make this remarkable fest even more impactful.</p>

            <p>Would it be possible to schedule a brief online meeting at a time convenient for you? Additionally, we would be grateful if you could share the contact details of the relevant person, and we’d be happy to connect with them over the phone.</p>

            <p>We believe this collaboration would be mutually beneficial, and we look forward to the opportunity to work together to make <strong>D3 Fest 2024</strong> an outstanding success.</p>

            <p>Thank you for your time and consideration.</p>

            <p>Warm regards,<br>
            <strong>D3 Fest Organizing Team</strong><br>
            IIIT Bhubaneswar<br>
            <strong>Phone:</strong> 9198851103<br>
            <strong>Email:</strong> dcube_techfest@iiit-bh.ac.in</p>
        </body>
        </html>
        """

        # Create the email with attachment
        message = create_message_with_attachment('me', email, subject, message_text, 'brochure.pdf')

        # Send the email
        send_message(service, 'me', message)

# Start the bulk email sending process
if __name__ == '__main__':
    send_bulk_emails()
