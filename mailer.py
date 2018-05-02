# Import smtplib for the actual sending function
import smtplib

import csv

# Import the email modules we'll need
from email.message import EmailMessage

import logging
import time
import argparse

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# create a file handler

handler = logging.FileHandler('Filtering.log',mode='w')
handler.setLevel(logging.INFO)

# create a logging format
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)

# add the handlers to the logger
logger.addHandler(handler)

#https://stackoverflow.com/questions/1557571/how-do-i-get-time-of-a-python-programs-execution
start_time = time.time()

parser = argparse.ArgumentParser(description='To send several emails.')
#parser.add_argument('--to', action='store_true')

# https://stackoverflow.com/questions/15753701/argparse-option-for-passing-a-list-as-option
# This is the correct way to handle accepting multiple arguments.
# '+' == 1 or more.
# '*' == 0 or more.
# '?' == 0 or 1.
# An int is an explicit number of arguments to accept.
parser.add_argument('--destin')
parser.add_argument('--template')
parser.add_argument('--password_email')

args = parser.parse_args()

def sendMail(iDestination, iTemplate):
    logger.info("Sending template to email : " + str(iDestination["email"]))

    # Create a text/plain message
    msg = EmailMessage()
    msg.set_content(iTemplate)

    # me == the sender's email address
    # you == the recipient's email address
    msg['Subject'] = 'etcequecamarche'
    msg['From'] = "charles.walker.37@outlook.com"
    msg['To'] = iDestination["email"]

    # Send the message via our own SMTP server.

    s = smtplib.SMTP('smtp-mail.outlook.com', 587, timeout=120)
    s.ehlo() # Hostname to send for this command defaults to the fully qualified domain name of the local host.
    s.starttls() #Puts connection to SMTP server in TLS mode
    s.ehlo()
    s.login('charles.walker.37@outlook.com', args.password_email)
    s.send_message(msg)
    logger.info("sent")
    s.quit()

if __name__== "__main__":
    with open(args.template, encoding="utf8") as aTemplate:
        aTemplate = aTemplate.read()
        logger.info("aTemplate: " + str(aTemplate))

        with open(args.destin, encoding="utf8") as aDest:
            logger.info("aDest: " + str(aDest))
            reader = csv.DictReader(aDest)
            for aOneEntry in reader:
                logger.info("aOneEntry: " + str(aOneEntry))
                sendMail(aOneEntry,aTemplate)