import os
import base64
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import smtplib
import csv

# Load CSV data
csv_file = 'email.csv'
df = pd.read_csv(csv_file)

# Setup Jinja2 environment
env = Environment(loader=FileSystemLoader('templates'))
template = env.get_template('email_template.html')

# SMTP configuration for Outlook
smtp_server = 'smtp.office365.com'
smtp_port = 587
smtp_user = 'enterUsername'
smtp_password = 'enterPassword'

# Function to send email
def send_email(to_email, subject, html_content, images):
    msg = MIMEMultipart('related')
    msg['Subject'] = subject
    msg['From'] = smtp_user
    msg['To'] = to_email

    msg_alternative = MIMEMultipart('alternative')
    msg.attach(msg_alternative)

    msg_text = MIMEText(html_content, 'html')
    msg_alternative.attach(msg_text)

    for image_path, cid in images.items():
        with open(image_path, 'rb') as img:
            msg_image = MIMEImage(img.read())
            msg_image.add_header('Content-ID', f'<{cid}>')
            msg.attach(msg_image)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, to_email, msg.as_string())

# Open and read a CSV file
with open('email.csv', 'r') as csvfile:
  reader = csv.reader(csvfile)
  for line in reader:
    to_email = line[0]
    company_name = line[1]
    
    # Render email content
    html_content = template.render(company=company_name)
    
    # Define inline images (add your image paths and CIDs here)
    images = {
        './templates/facilitypicture.png': 'image1_cid'
    }
    
    subject = f'Hello {company_name}!'
    send_email(to_email, subject, html_content, images)

print('Emails sent successfully!')