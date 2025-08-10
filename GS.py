import pandas as pd
import smtplib
import json
import uuid
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

# Load config
with open("GS.json") as f:
    config = json.load(f)

SMTP_SERVER = config["smtp_server"]
SMTP_PORT = config["smtp_port"]
SENDER_EMAIL = config["sender_email"]
SENDER_PASSWORD = config["sender_password"]
EMAIL_SUBJECT = config["subject"]

# Extract domain from sender email for Message-ID
email_domain_match = re.search(r"@(.+)", SENDER_EMAIL)
email_domain = email_domain_match.group(1) if email_domain_match else "example.com"

# Load contacts
contacts = pd.read_excel("emails_validatedd.xlsx")
print("Excel columns:", contacts.columns.tolist())

# Load email body template
with open("GS.txt", "r") as f:
    body_template = f.read()

# Signature block
signature = """
--
Regards,  
Sai Deekshith  
 
"""

# Start SMTP session
server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls()
server.login(SENDER_EMAIL, SENDER_PASSWORD)

# Send emails
for index, row in contacts.iterrows():
    try:
        recipient_email = row["EMAIL"]
        full_name = row.get("NAME", "")

        if pd.isna(full_name) or not str(full_name).strip():
            full_name = ""
            first_name = ""
        else:
            first_name = str(full_name).split()[0]

        # Format body
        personalized_body = body_template.format(
            first_name=first_name or "there", full_name=full_name
        ).strip()
        full_message = personalized_body + "\n\n" + signature

        # Compose email
        msg = MIMEMultipart()
        msg["From"] = formataddr(("Sai Deekshith", SENDER_EMAIL))
        msg["To"] = recipient_email
        msg["Subject"] = EMAIL_SUBJECT
        msg["Message-ID"] = f"<{uuid.uuid4()}@{email_domain}>"

        msg.attach(MIMEText(full_message, "plain"))

        # Send
        server.send_message(msg)
        print(f"✅ Email sent to {full_name} <{recipient_email}>")

    except Exception as e:
        print(f"❌ Failed to send email to {recipient_email} at row {index}: {e}")

server.quit()