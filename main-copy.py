import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import datetime

def write_logs_001(logg):
    with open("log.log", "a") as f:
        f.write(f"{datetime.datetime.now()} -- {logg}\n")

def load_xls_002():
    try:
        fi = r"your path to file"
        df1 = pd.read_excel(fi, usecols='A, J, W')
        for i in df1.iterrows():
            if "NaN" not in str(i[1]) and "Unknown" not in str(i[1]):
                a = str(i[1])
                a = a[a.index("Domain"):a.index("Name")].replace(" ", "").replace("00:00:00", "").replace("Domain:", "").replace("Expired:", "").replace("\n", " ").replace("\r", "").replace("\t", "")
                dd_today = datetime.datetime.strptime(str(datetime.date.today()), "%Y-%m-%d")
                a = a.replace("\n", "")
                a = a.split(" ")
                dd = datetime.datetime.strptime(a[2], "%Y-%m-%d")
                delta = dd - dd_today
                if delta.days < 30:
                    sendemail_003(a[0], delta.days, a[1])

    except Exception as e:
        write_logs_001(f"Error #002 - Reason: {e}")

def sendemail_003(dom, dayz, contact):
    try:
        email_list = ["list@list", "of@of", "senders@senders"]
        for i in email_list:
            email = "domain_sender@yourdomain.com"  # the email where you sent the email
            password = "your password"
            subject = f"Domain {dom} soon to expire"
            message = f"Domain {dom} is soon to expire please act on this matter. \nSend email to {contact} \n\nDomain will expire in {dayz} Days."
            msg = MIMEMultipart()
            msg["From"] = email
            msg["To"] = i
            msg["Subject"] = subject
            msg.attach(MIMEText(message, 'plain'))
            server = smtplib.SMTP("SMTP-SERVER", 587)
            server.starttls()
            server.login(email, password)
            text = msg.as_string()
            server.sendmail(email, i, text)
            server.quit()
            write_logs_001(f"Email sent to {i} about {dom}")

    except Exception as e:
        write_logs_001(f"Error #003 - Reason: {e}")

load_xls_002()
