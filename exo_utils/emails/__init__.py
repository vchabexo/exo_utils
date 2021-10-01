import smtplib
from email.mime.multipart import MIMEMultipart
from email.utils import COMMASPACE, formatdate
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from pathlib import Path


def send_mail(send_from, send_to, subject, message, files=[], server="localhost", port=587, username='', password=''):
    """Compose and send email with provided info and attachments.

    Args:
        send_from (str): from name
        send_to (list[str]): to name(s)
        subject (str): message title
        message (str): message body
        files (list[str]): list of file paths to be attached to email
        server (str): mail server host name
        port (int): port number
        username (str): server auth username
        password (str): server auth password
    """


    #construction of the message
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    #adding attachemements
    msg.attach(MIMEText(message))

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename="{}"'.format(Path(path).name))
        msg.attach(part)

    #connection to server
    server  = smtplib.SMTP(server, port)
    server.ehlo()
    server.starttls()
    server.ehlo()

    server.login(username, password)

    #sending the email
    server.sendmail(send_from, send_to, msg.as_string())
    server.quit()

    return