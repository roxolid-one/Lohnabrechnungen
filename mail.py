import ntpath
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import constants


class Mail:
    MY_EMAIL = constants.MY_EMAIL
    PASSWORD = constants.PASSWORD

    def __init__(self, recipients, subject, body, attachment_path=None):
        self.sender = self.MY_EMAIL
        self.receiver = recipients
        self.subject = subject
        self.body = body
        self.attachment_path = attachment_path
        self.message = MIMEMultipart()
        self.attachment_name = ''

        self.compiled_message = self.compile_message()

    def compile_message(self):
        self.message['From'] = self.sender
        self.message['To'] = self.receiver
        self.message['Subject'] = self.subject
        self.message.attach(MIMEText(self.body, 'plain'))
        if self.attachment_path is not None:
            path, self.attachment_name = ntpath.split(self.attachment_path)
            with open(self.attachment_path, 'rb') as binary_attch:
                payload = MIMEBase('application', 'octate-stream', Name=self.attachment_name)
                payload.set_payload((binary_attch).read())
                encoders.encode_base64(payload)
                payload.add_header('Content-Decomposition', 'attachement', filename=self.attachment_name)
                self.message.attach(payload)
        return self.message.as_string()

    def send_mail(self):
        message = self.compiled_message

        with smtplib.SMTP("smtp.gmail.com") as connection:
            connection.starttls()
            connection.login(user=self.MY_EMAIL, password=self.PASSWORD)
            connection.sendmail(from_addr=self.sender, to_addrs=self.receiver, msg=message)
