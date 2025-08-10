import smtplib

smtp_server = "smtp.gmail.com"
smtp_port = 587
email = "csaideejobs@gmail.com"
password = "ynmhsxkyiksmnhqm"  # no spaces

try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email, password)
    print("✅ Login successful!")
    server.quit()
except Exception as e:
    print(f"❌ Login failed: {e}")