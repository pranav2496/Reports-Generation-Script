
from urllib.request import HTTPError
import python_http_client
import datetime
import requests
import hashlib
import binascii
import hmac
import time, json
import os,sys,getopt,re
import sendgrid
import base64
import logging
import numpy
from urllib.request import HTTPError
import python_http_client
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (To, Cc, Mail, Personalization, Email, Attachment,
FileContent, FileName, FileType, Disposition, ContentId, Content)

SENDGRID_API_KEY = '-----------Enter the SendGrid API key here -----------------'


logging.basicConfig(filename="data/logfile.log",
                            filemode='a',
                            format='%(asctime)s,%(msecs)d :: %(name)s :: %(levelname)s :: %(message)s',
                            level=logging.DEBUG)

logging.info("Process Started")

def SendEmail():
    try:


        #filename = "data\\*.xlsx"
        #list_of_files = glob.glob(f"*{filename}") # * means all if need specific format then *.csv
        #latest_file = max(list_of_files, key=os.path.getctime)

        emails = ["pranav@manchtech.com"]
        ccemails = []
        subject = "Coke Sap Report"
        attachment = document_download()

        p = Personalization()
        from_email= Email("no-reply@manchtech.com","Support Team At Manchtech")
        for to_email_id in emails:
            to_email = To(to_email_id)
            p.add_to(to_email)
        if Enquiry(ccemails).size: 
            for cc_email_id in ccemails:
                cc_email = Cc(cc_email_id)
                p.add_cc(cc_email)
        content = Content("text/plain", 'PFA report')

        message = Mail(from_email, to_email, subject, content)
        message.add_personalization(p)


        filename = os.path.basename(attachment)
        file_extension = os.path.splitext(filename)[1].split('.')[1]

        
        with open(attachment, 'rb') as f:
            data = f.read()
        f.close()

    except(IndexError, UnboundLocalError) as e:
        print("Error : " + str(e))
        logging.error(e)


    encoded = base64.b64encode(data).decode()
    attachment = Attachment()
    attachment.file_content = FileContent(encoded)
    attachment.file_type = FileType('application/'+file_extension)
    attachment.file_name = FileName(filename)
    attachment.disposition = Disposition('attachment')
    attachment.content_id = ContentId('Example Content ID')
    message.attachment = attachment


    try:
        sendgrid_client = sendgrid.SendGridAPIClient(api_key=SENDGRID_API_KEY)
        response = sendgrid_client.client.mail.send.post(request_body=message.get())
        logging.debug("Email is Sent with File: "+filename)
        print("Success..")
    except(HTTPError, python_http_client.exceptions.BadRequestsError) as e:
        print("Error : " + str(response.e))
        logging.error(e)

def Enquiry(lis1): 
    return(numpy.array(lis1)) 
      
def document_download():
    try:

        authorization = login_authorization()

        url = "https://uat.manchtech.com/app/admin/custom-report/company/92/report/generateConsolidatedReport"

        payload = {}
        files = {}

        headers = {
            'x-authorization': '--------------------------'
        }

        headers['x-authorization'] = authorization

        response = requests.request("GET", url, headers=headers, data = payload, files = files)
        now = datetime.datetime.now()

        if response.status_code == 200:
            filepath = "data/CokeSapReport-{}".format(now.strftime('%d%b%Y-%H%M%S'))+".xlsx"
            with open(filepath,'wb') as f:
                f.write(response.content)
            logging.debug("Report File "+ os.path.basename(filepath) +" is Downloaded")
            return f.name
        else:
            print(response.content)

    except(HTTPError, python_http_client.BadRequestsError) as e:
        print("Error : " + str(e))
        logging.error(e)



def login_authorization():

    try:
        username = "---enter username----"
        password = "---password ----"

        url = "https://uat.manchtech.com/app/user/login"

        payload = "{\n  \"email\": \"%s\",\n  \"password\": \"%s\"\n}" %(username,password)


        headers = {
        'Content-Type': 'application/json'
        }

        response = requests.request("POST", url, headers=headers, data = payload)

        data = json.loads(response.content.decode('utf-8'))
        return data['payLoad']['authToken']

    except(IndexError, UnboundLocalError) as error:
        print("Error : "+str(error))
        logging.error(error)



if __name__ == "__main__":

    from urllib3.exceptions import NewConnectionError
    from requests.exceptions import ConnectionError

    try:
        SendEmail()

    except(ConnectionError, HTTPError, NewConnectionError) as e:
        print
        logging.critical("Internet is Down..")

                                        