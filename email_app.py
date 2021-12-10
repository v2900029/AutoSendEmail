import os
import argparse
import configparser
from google_email_api.quickstart import main as google_email_service

from oauth2client.service_account import ServiceAccountCredentials as SAC
import gspread
from google_email_api.Google import Create_Service
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def init_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('-cp', '--configs_path', default='./configs', type=str, help='The configs path')
    # example:
    # args.add_argument("arg1", nargs=3, type=int,help="這是第 1 個引數，請輸入三個整數")
    return parser.parse_args()


def prepare_GoogleForms_api(configs_path, google_client_url='https://spreadsheets.google.com/feeds'):
    Json = '{}/client_secrets.json'.format(configs_path) 
    Url = [google_client_url]

    config_ini = configparser.ConfigParser()
    config_ini.read(config_ini_path, encoding="utf-8")
    GoogleSheets_key = config_ini.get('GoogleSheets', 'GoogleSheets_key')# GoogleSheets_key = '1O1fx92flO-l8sCXIJisEA_-4iG3JIuXrMPIuSvQZXGc'
    
    Connect = SAC.from_json_keyfile_name(Json, Url)
    GoogleSheets = gspread.authorize(Connect)
    google_form = GoogleSheets.open_by_key(GoogleSheets_key) 
    return google_form


def prepare_Gmail_api(configs_path, API_NAME='gmail', API_VERSION='v1', scope='https://mail.google.com/'):
    CLIENT_SECRET_FILE = '{}/credentials.json'.format(configs_path)
    SCOPES = [scope]
    service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)    
    return service


def create_email_message(email_address = '??@??', download_link='https://drive.google.com/???????=sharing'):
    email = MIMEMultipart()
    email['to'] = email_address
    email['subject'] = 'This is a title.'

    context = '''Welcome, \n\nTo download the file, please click on the following link:\n{0} \n\nAuthor'''.format(download_link)
    email.attach(MIMEText(context, 'plain'))
    return {'raw': base64.urlsafe_b64encode(email.as_string())} 


if __name__ == '__main__':
    args = init_args()

    if not os.path.isdir(args.configs_path):
        # make sure configs are exist.
        os.makedirs(args.configs_path)

    # first time to use Google Api => inintial GOOGLE email API
    google_email_service(args.configs_path)

    # connect to the google form
    google_form = prepare_GoogleForms_api(configs_path=args.configs_path)
    # get target form's sheet
    target_sheet = google_form.sheet1
    # Gmail Service args
    gmail_service = prepare_Gmail_api(configs_path=args.configs_path, API_NAME='gmail', API_VERSION='v1', scope='https://mail.google.com/')

    sending_flag = target_sheet.cell(2,7).value
    endpoint = target_sheet.cell(2,7).value
    i = 2 # What is the meaning of the "i"
    if endpoint != 1:
        target_sheet.update_cell(3,7, "2")

    while sending_flag != '2':    
        email_address = target_sheet.cell(i,2).value
        download_link = 'https://drive.google.com/drive/folders/??????????=sharing'

        if sending_flag == '1':
            print('N') # Is it necessary to use this if-else conditional expressions?
        else:
            email_message = create_email_message(email_address=email_address, download_link=download_link)
            message = gmail_service.users().messages().send(userId='me', body=email_message).execute()
            print(message)
            target_sheet.update_cell(i,7, "1")
    
        i+=1
        sending_flag = target_sheet.cell(i,7).value
        print (sending_flag)