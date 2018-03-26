
import configparser
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate


def send_mail(send_from, send_to, subject, text, files=None,
              server="127.0.0.1"):
    assert isinstance(send_to, list)
    
    username = cf.get("EMAIL","user")
    password = cf.get("EMAIL","password")
    
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
            msg.attach(part)


    smtp = smtplib.SMTP(server)   
    smtp.ehlo()
    smtp.starttls()
    smtp.login(username,password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()
    

server = 'smtp.gmail.com:587'
send_from = 'bain.bai@everfast.com.au'
send_to = ['bain.bai@everfast.com.au']
subject = 'test'
text = 'test'
files = ["timetest.py","timetest2.py"]
send_mail(send_from,send_to,subject,text,files,server)