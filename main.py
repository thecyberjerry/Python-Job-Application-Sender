import sys
from email.mime.text import MIMEText
import smtplib
import pandas as pd
import re
import os
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from PIL import Image, ImageFont, ImageDraw, ImageColor

load_dotenv()
subject = "Application for React Developer"
body = """Dear HR Manager,

I am writing to express my interest in the position of React developer as advertised recently. My qualifications, skills, and experience align closely with your requirements for this role.

Please find my CV and supporting documents attached for your review. I would be delighted to discuss how I can contribute to your team and look forward to hearing from you soon about this exciting opportunity.

Kind regards,

Himanshu"""
sender = os.getenv("SENDER_EMAIL")
password = os.getenv("SENDER_EMAIL_PASSWORD")
resume = "ResumePath or Resume Name here"
EmailsFile = "Emails.xlsx"

try:
    file_name = EmailsFile
    df = pd.read_excel(file_name)
    emails_list = []
    invalid_emails = []
    for index, row in df.iterrows():
        for col in df.columns:
            if isinstance(row[col], str):
                valid = re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', row[col])
                emails_list.append(row[col]) if valid else invalid_emails.append(row[col])
            else:
                invalid_emails.append(row[col])
                continue
    recipients = emails_list
    if not len(emails_list): raise Exception("Empty emails list")
    print("Invalid Email entry found:\n", invalid_emails) if len(invalid_emails) else print()
except Exception as e:
    print(e) if e else print("Some Error Occured in Emails")
    sys.exit()


def send_email(subject, body, sender, recipients, password):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender
    msg.attach(MIMEText(body, "plain"))
    with open(resume, 'rb') as resume:
        attach = MIMEApplication(resume.read(), _subtype="pdf")
    attach.add_header('Content-Disposition', 'attachment', filename=str(resume))
    msg.attach(attach)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
        smtp_server.login(sender, password)
        smtp_server.send_message(msg, sender, recipients)
    print("Message sent!")

try:
    send_email(subject, body, sender, recipients, password)
except:
    print("Some error occured while sending mails")
