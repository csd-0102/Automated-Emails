import pandas as pd
import smtplib
import json
import uuid
import re
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr

# Load config
with open("config.json") as f:
    config = json.load(f)

SMTP_SERVER = config["smtp_server"]
SMTP_PORT = config["smtp_port"]
SENDER_EMAIL = config["sender_email"]
SENDER_PASSWORD = config["sender_password"]
EMAIL_SUBJECT = config["subject"]

# Extract domain from sender email
email_domain_match = re.search(r"@(.+)", SENDER_EMAIL)
email_domain = email_domain_match.group(1) if email_domain_match else "example.com"

# Load Excel with sheet names
excel_path = "emails_validatedd.xlsx"
xls = pd.read_excel(excel_path, sheet_name=None)

# Get pending and sent contacts
pending_df = xls.get("Pending", pd.DataFrame())
sent_df = xls.get("Sent", pd.DataFrame())

# Process only the first 50
to_send = pending_df.head(10)
remaining = pending_df.iloc[10:]

# Load email body
with open("email_body.txt", "r") as f:
    body_template = f.read()

# Load resume
with open("resume.pdf", "rb") as f:
    resume_data = f.read()

# Signature
signature = """
--
Regards,  
Sai Deekshith  
Data Engineer | Data Analyst  
📧 csaideejobs@gmail.com  
📞 +1-518-517-9558  
"""

# Start SMTP session
server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls()
server.login(SENDER_EMAIL, SENDER_PASSWORD)

# Keep track of rows that were successfully sent
successfully_sent = []

for index, row in to_send.iterrows():
    try:
        recipient_email = row["EMAIL"]
        full_name = row.get("NAME", "")
        first_name = str(full_name).split()[0] if full_name and not pd.isna(full_name) else ""

        personalized_body = body_template.format(
            first_name=first_name or "there", full_name=full_name
        ).strip()
        full_message = personalized_body + "\n\n" + signature

        # Compose message
        msg = MIMEMultipart()
        msg["From"] = formataddr(("Sai Deekshith", SENDER_EMAIL))
        msg["To"] = recipient_email
        msg["Subject"] = EMAIL_SUBJECT
        msg["Message-ID"] = f"<{uuid.uuid4()}@{email_domain}>"
        msg.attach(MIMEText(full_message, "plain"))

        # Attach resume
        attachment = MIMEApplication(resume_data, _subtype="pdf")
        attachment.add_header("Content-Disposition", "attachment", filename="Sai_Deekshith_Chinthalwar.pdf")
        msg.attach(attachment)

        server.send_message(msg)
        print(f"✅ Email sent to {full_name} <{recipient_email}>")

        # Save row to sent list
        successfully_sent.append(row)

    except Exception as e:
        print(f"❌ Failed to send to {recipient_email} at row {index}: {e}")

server.quit()

# Append successful rows to 'Sent' sheet
if successfully_sent:
    sent_df = pd.concat([sent_df, pd.DataFrame(successfully_sent)], ignore_index=True)

# Save remaining rows in 'Pending', updated 'Sent'
with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
    remaining.to_excel(writer, index=False, sheet_name="Pending")
    sent_df.to_excel(writer, index=False, sheet_name="Sent")
