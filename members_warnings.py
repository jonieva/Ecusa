import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
import datetime

# import base64
# import httplib2
# from email.mime.text import MIMEText
#
# import oauth2client
# from oauth2client import client
# from oauth2client import tools
#
# from apiclient.discovery import build
# from oauth2client.client import flow_from_clientsecrets
from oauth2client.file import Storage

# TO INSTALL API:   pip install --upgrade google-api-python-client
#                   pip install --upgrade python-gflags
from apiclient import discovery

credentials_file = "Ecusa-credentials-secretary.json"

def get(row, letter):
    return row[ord(letter) - ord('A')].decode("utf-8")


def members_about_to_expire_warning(warn_about_members_expired=True):
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    gc = gspread.authorize(credentials)

    # Open a worksheet from spreadsheet with one shot
    wks = gc.open("Active Full Members").sheet1
    #wks = gc.open_by_key('1OhLymORZbrteqqM0F2tFgC4sCEqANAQEiLajegDmtWI')
    #wks = gc.open_by_url('https://docs.google.com/spreadsheets/d/1OhLymORZbrteqqM0F2tFgC4sCEqANAQEiLajegDmtWI')
    # Fetch a cell range
    i = 2
    row = wks.row_values(i)
    row_count = wks.row_count

    date = get(row, 'M')

    first_reminder_timespan_days = -4
    second_reminder_timespan_days = 0
    today = datetime.datetime.today().date()
    members_about_to_expire = ["About to"]
    members_expired = [""]

    # while date and i < row_count:
    #     d = datetime.datetime.strptime(date, "%m/%d/%Y %H:%M:%S").date()
    #     if (today-d).days == 365-first_reminder_timespan_days:
    #         members_about_to_expire.append(u"Name: {};\nDate: {};\nEmail: {}\nRow number: {}\n\n".
    #                                        format(get(row, 'P'), get(row, 'M'), get(row, 'T'), i))
    #     elif (today-d).days >= 365-second_reminder_timespan_days:
    #         members_expired.append(u"Name: {};\nDate: {};\nEmail: {}\nRow number: {}\n\n".
    #                                        format(get(row, 'P'), get(row, 'M'), get(row, 'T'), i))
    #     #message_text += u"Name: {};\nDate: {};\nEmail: {}\nRow: {}".format(get(row, 'P'), get(row, 'M'), get(row, 'T'), i)
    #     i += 1
    #     row = wks.row_values(i)
    #     date = get(row, 'M')

    if not members_about_to_expire and not members_expired:
        return
    if not members_about_to_expire and not warn_about_members_expired:
        return

    message_text = "This is an automatic notification for ECUSA secretary.\n\n"
    if members_about_to_expire:
        message_text += "Members about to expire:\n\n"
        message_text += "\n".join(members_about_to_expire)
        message_text += "\n\n************************"
    if members_expired:
        message_text += "\nMembers EXPIRED:\n\n"
        message_text += "\n".join(members_expired)

    send_email("jorgeonieva@gmail.com", "ECUSA users notification", message_text)


def send_email(toaddrs, subject, body):
    fromaddr = 'secretary-bos@ecusa.es'
    with open("mail_pwd.txt", 'r') as f:
        password=f.readline()

    server = smtplib.SMTP('smtp.gmail.com:587')
    server.ehlo()
    server.starttls()
    server.login(fromaddr, password)

    msg = "\r\n".join([
        "From: {}",
        "To: {}",
        "Subject: {}",
        "",
        "{}"
    ]).format(fromaddr, toaddrs, subject, body)

    server.sendmail(fromaddr, toaddrs, msg)
    server.quit()



if __name__ == "__main__":
    members_about_to_expire_warning()

# def create_message():
#     message_text= "hola"
#     message = MIMEText(message_text)
#     message['to'] = "jorgeonieva@gmail.com"
#     message['from'] = "secretary-bos@ecusa.es"
#     message['subject'] = "Users about to expire"
#     m = {'raw': base64.urlsafe_b64encode(message.as_string())}
#     return m

# cell_list = wks.range('M2:T400')
# len(cell_list)
# for c in (c.value for c in cell_list if c.value.strip() != ''):
#     print c
#
# c = cell_list[0]

# http = credentials.authorize(httplib2.Http())
# service = discovery.build('gmail', 'v1', http=http)
# email = (service.users().messages().send(userId="jorgeonieva@gmail.com", body=m).execute())
#
#
# 1085409003224-ppb5797vn71eih59uo1bl6v9i2ihihbr.apps.googleusercontent.com
# DPZkqS_1q_MKXxfrDaCMrhMu
#
# # SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
# CLIENT_SECRET_FILE = '/Data/jonieva/Dropbox/Ecusa/Ecusa-credentials-secretary-OAuth.json'
# APPLICATION_NAME = 'ECUSA'
# # Check https://developers.google.com/gmail/api/auth/scopes for all available scopes
# SCOPE = 'https://www.googleapis.com/auth/gmail.compose'



# # Try to retrieve credentials from storage or run the flow to generate them
# credentials = STORAGE.get()
# if credentials is None or credentials.invalid:
#   credentials = tools.run_flow(flow, STORAGE)
#
# # Authorize the httplib2.Http object with our credentials
# http = credentials.authorize(http)
#
# # Build the Gmail service from discovery
# gmail_service = build('gmail', 'v1', http=http)



# def get_credentials():
#     """Gets valid user credentials from storage.
#
#     If nothing has been stored, or if the stored credentials are invalid,
#     the OAuth2 flow is completed to obtain the new credentials.
#
#     Returns:
#         Credentials, the obtained credential.
#     """
#     try:
#         import argparse
#         flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
#     except ImportError:
#         flags = None
#
#     home_dir = os.path.expanduser('~')
#     credential_dir = os.path.join(home_dir, 'tmp', '.credentials')
#     if not os.path.exists(credential_dir):
#         os.makedirs(credential_dir)
#     credential_path = os.path.join(credential_dir, 'ecusa-mail.json')
#
#     store = oauth2client.file.Storage(credential_path)
#     credentials = store.get()
#     if not credentials or credentials.invalid:
#         flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPE)
#         flow.user_agent = APPLICATION_NAME
#         if flags:
#             credentials = tools.run_flow(flow, store, flags)
#         else: # Needed only for compatibility with Python 2.6
#             credentials = tools.run(flow, store)
#         print('Storing credentials to ' + credential_path)
#     return credentials


