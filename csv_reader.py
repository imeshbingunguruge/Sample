import csv
import requests
import smtplib
import ssl
import logging
import traceback
import io
from io import BytesIO
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
import pandas as pd
from openpyxl import Workbook,load_workbook

#consider utilizing an online document and access the data , will leave the code commented as this is only an improvement suggestion

# def get_sharepoint(): 

#     # SharePoint site URL and credentials
#     site_url = "https://asdasd.sharepoint.com/sites/yoursite"
#     username = sharepoint_username
#     password = sharepoint_password

#     # Read data from Excel sheet using pandas
#     relative_file_url = 'relative_url/users.csv'

#     try:
#         # Authenticate with SharePoint using Office365-REST-Python-Client
#         ctx_auth = AuthenticationContext(site_url)
#         if ctx_auth.acquire_token_for_user(username, password):
#             ctx = ClientContext(site_url, ctx_auth)
#             web = ctx.web
#             ctx.load(web)
#             ctx.execute_query()

#             response = File.open_binary(ctx, relative_file_url)
#             ###print(response)

#             #save data to BytesIO stream
#             bytes_file_obj = io.BytesIO()
#             bytes_file_obj.write(response.content)
#             bytes_file_obj.seek(0) #set file object to start

#             #read file into pandas dataframe
#             df = pd.read_csv(bytes_file_obj, sheet_name = "sheetname", engine='openpyxl')

#             # Print the data
#             #print(df)

#             lists = df.values.tolist()

#             return lists

#         else:
#             print("Failed to authenticate")
                
#     except Exception:
#         print("An error occurred while accessing data in sharepoint")


#initiate the logging
logging.basicConfig(filename='error_log.txt',filemode='a+',format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

#email functionality , utilize internal smtp server to send mails , additional functions are commented which can be used depeneding on the use case
def mail_alert(to_add, msg, sub):
    port = 25
    smtp_server = "intrelay.lseg.com"
    sender_email = "example@example.com"
    to_address = ["additional_email@example.com"]
    to_address.append(to_add)
    receiver_emails = to_address
    subject = sub
    body_text = "Hi Team,\n\n{}\nThanks and Regards,\nYour Team Name".format(msg)
    
    
    msg = EmailMessage()
    # msg = MIMEMultipart("alternative")
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = to_address

    # body_text = (f"Hi Team,\n\n{msg}\n\nThanks and Regards,\nDatadog Platform Team")
    msg.set_content(body_text)
    # msg.add_alternative(body_text, subtype="html")
    # msg.attach(MIMEText(body_text,"plain"))
    # msg.attach(MIMEText(body_text,"html"))

    # print(msg)

    context = ssl._create_unverified_context()
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls(context=context)
        server.sendmail(sender_email, receiver_emails, msg.as_string())
        server.quit()


#check email is not null, this can also be extended to check the content as well using a regex if needed
def check_email(email_value):
     if email_value == '':
          return False
     else:
          return True



def create_users(file_path):

    #get the users that will not be created into this
    error_creating_users = []
    
    #utilize the utf encoding so that the words are properly formatted
    with open(file_path, mode ='r' ,encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            print(row)
            valid_email = check_email(row['email'])

            print(valid_email)

            try:

               response=requests.post("https://example.com/api/create_user";,json=row)

               if response.status_code != 201:
                  error_creating_user.append(row['name'])
                  print("Error creating user:" , row["email"])


            except Exception as e:
               print(traceback.format_exc())
               print(e)
               logging.exception(e)
               error_creating_user.append(row['name'])

     new_msg = "Hi," + " \n\n" + "This is to inform you that the following user were not created"  + " \n\n"

     for user_not_created in error_creating_users:
          new_msg = new_msg + str(user_not_created) + "\n"
                            
     new_msg = new_msg + " \n\n" + "please take necessary action if applicable" 

     mail_alert("additional_cc_email@example.com",new_msg, "[ACTION REQUIRED] : User(s) not created")
     

create_users("users.csv")