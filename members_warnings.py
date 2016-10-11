import os
import argparse
import gspread
from oauth2client.service_account import ServiceAccountCredentials      # pip install oauth2client
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

credentials_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "Ecusa-credentials-secretary.json")
mail_pwd_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "mail_pwd.txt")
fromaddr = 'secretary-bos@ecusa.es'


def get(row, letter):
    return row[ord(letter) - ord('A')]


def members_about_to_expire_warning(destination, first_reminder_timespan_days=7, expired_for_at_least_timespan_days=0):
    """
    Send an email warning about the members that are about to expire and the members that already expired
    :param destination: list od email addresses
    :param first_reminder_timespan_days: user is going to expire in X days
    :param expired_for_at_least_timespan_days: user has been expired for at least X days
    :return:
    """
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

    today = datetime.datetime.today().date()
    members_about_to_expire = []
    members_expired = [""]

    if expired_for_at_least_timespan_days > 0:
        expired_for_at_least_timespan_days = -expired_for_at_least_timespan_days

    while date and i < row_count:
        print ("Processing {}".format(i))
        d = datetime.datetime.strptime(date, "%m/%d/%Y %H:%M:%S").date()
        if (today-d).days == 365-first_reminder_timespan_days:
            members_about_to_expire.append(u"Name: {};\nDate: {};\nEmail: {}\nRow number: {}\n\n".
                                           format(get(row, 'P'), get(row, 'M'), get(row, 'T'), i))
        elif (today-d).days >= 365-expired_for_at_least_timespan_days:
            members_expired.append(u"Name: {};\nDate: {};\nEmail: {}\nRow number: {}\n\n".
                                           format(get(row, 'P'), get(row, 'M'), get(row, 'T'), i))
        i += 1
        row = wks.row_values(i)
        date = get(row, 'M')

    if not members_about_to_expire and not members_expired:
        return

    message_text = "This is an automatic notification for ECUSA secretary.\n\n"
    if members_about_to_expire:
        message_text += "Members about to expire:\n\n"
        message_text += "\n".join(members_about_to_expire)
        message_text += "\n\n************************"
    if members_expired:
        message_text += "\nMembers EXPIRED:\n\n"
        message_text += "\n".join(members_expired)

    send_email(destination, "ECUSA users notification", message_text)


def send_email(toaddrs, subject, body):
    with open(mail_pwd_file, 'r') as f:
        password=f.readline()

    server = smtplib.SMTP('smtp.gmail.com:587')
    server.ehlo()
    server.starttls()
    server.login(fromaddr, password)

    msg = u"\r\n".join([
        "From: ECUSA automatic tools",
        "To: {}",
        "Subject: {}",
        "",
        "{}"
    ]).format(toaddrs, subject, body).encode("utf8")

    server.sendmail(fromaddr, toaddrs, msg)
    server.quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='ECUSA notifications')
    parser.add_argument('--destination', type=str, required=True,
                        help="Email address of the destination (if more than one, separate them with ;)")
    parser.add_argument('--first_reminder', type=int, default=7,
                        help="A reminder of the users that are going to expire in X days",
                        )
    parser.add_argument('--expired_for_at_least', type=int, default=0,
                        help="Replace users considered expired. For example, if expired_for_at_least = 7, only warn "
                             "about the users that have been expired for more than 7 days")
    args = parser.parse_args()
    dest = map(str.strip, args.destination.split(';'))
    members_about_to_expire_warning(dest, first_reminder_timespan_days=args.first_reminder,
                                    expired_for_at_least_timespan_days=args.expired_for_at_least)



