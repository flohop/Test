import csv
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from random import randint
import smtplib

creds = os.path.dirname(os.path.abspath(__file__)) + "\credentials" + "\email.csv"


def send_email(email_send, article, article_name="Artikel"):

    # Enter an email you want to send mails to others

    with open(creds, 'r') as email_infos:
        e_reader = csv.reader(email_infos)
        for row in e_reader:
            email_user = row[0]
            email_password = row[1]

    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = article_name

    body = (
    'This email was sent to you by a bot, below is your requested .pdf file.\n'
    'if you didnt request a file, please ignore this email.'
    )
    msg.attach(MIMEText(body, 'plain'))

    filename = article

    part = MIMEBase('application', 'octet-stream')
    with open(filename, 'rb') as attachment_file:
        part.set_payload(attachment_file.read())

    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= "+filename)

    msg.attach(part)
    text = msg.as_string()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email_user, email_password)

    server.sendmail(email_user, email_send, text)
    server.quit()


class Verification():
    """ Send a verification code to the user and return the code via a function"""
    def __init__(self, email_send):
        self.email_send = email_send
        self.verif_code = randint(1000, 9999)

    def get_verif_code(self):
        """"Get the verification code as an int"""
        return self.verif_code

    def send_verify_mail(self):
        """Send an email with the verificatian code to the user"""

        with open(creds, 'r') as email_infos:
            e_reader = csv.reader(email_infos)
            for row in e_reader:
                email_user = row[0]
                email_password = row[1]

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['To'] = self.email_send
        msg['Subject'] = "Email verification"

        body = "Verification code: " + str(self.verif_code)
        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, email_password)

        server.sendmail(email_user, self.email_send, text)
        server.quit()


