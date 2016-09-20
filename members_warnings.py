import argparse
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
    # Fetch a cell range
    i = 2
    row = wks.row_values(i)
    row_count = wks.row_count

    date = get(row, 'M')

    first_reminder_timespan_days = -4
    second_reminder_timespan_days = 0
    today = datetime.datetime.today().date()
    members_about_to_expire = []
    members_expired = [""]

    while date and i < row_count:
        d = datetime.datetime.strptime(date, "%m/%d/%Y %H:%M:%S").date()
        if (today-d).days == 365-first_reminder_timespan_days:
            members_about_to_expire.append(u"Name: {};\nDate: {};\nEmail: {}\nRow number: {}\n\n".
                                           format(get(row, 'P'), get(row, 'M'), get(row, 'T'), i))
        elif (today-d).days >= 365-second_reminder_timespan_days:
            members_expired.append(u"Name: {};\nDate: {};\nEmail: {}\nRow number: {}\n\n".
                                           format(get(row, 'P'), get(row, 'M'), get(row, 'T'), i))
        i += 1
        row = wks.row_values(i)
        date = get(row, 'M')

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
    parser = argparse.ArgumentParser(description='ECUSA notifications')
    parser.add_argument('--warn_expired_users', type=bool,
                        help="When False, there won't be a notification if we have only expired users", default=True)

    args = parser.parse_args()
    members_about_to_expire_warning(args.warn_expired_users)



