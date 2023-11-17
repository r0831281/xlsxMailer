#!/usr/bin/python3

#app id ff91f935-c699-4f05-8844-0570bd88b673


import argparse
import pandas as pd
import smtplib
import getpass
from O365 import Message, MSGraphProtocol, Account, FileSystemTokenBackend


scopes = ['https://graph.microsoft.com/.default/Mail.ReadWrite', 'https://graph.microsoft.com/Mail.Send', 'basic']


html_template =     """ 
            <html>
            <head>
                <title></title>
            </head>
            <body>
                    {}
            </body>
            </html>
        """

parser = argparse.ArgumentParser(description='get site data.')
parser.add_argument('site', metavar='-s', type=int, nargs='+',
                    help='site name')
parser.add_argument('xlxs', metavar='-x', type=str, nargs='+', help='xlxs file location')
parser.add_argument('--version', '-V', action='version', version='%(prog)s 1.0')
parser.add_argument('--verbose', '-v', action='count', default=0)
parser.add_help = True


protocol = MSGraphProtocol()


def authenticate_to_outlook(sendr):
    try:
           
            client_secret = 'MC~8Q~82rBtOVn2SWA0bkJ8nPJZJUwdpCQWNoa_a'
            client_id = 'ff91f935-c699-4f05-8844-0570bd88b673'
            credentials = (client_id, client_secret)
            token_backend = FileSystemTokenBackend(token_path='/tokenStore', token_filename='my_token.txt')
            o365_auth = Account(credentials = credentials, auth_flow_type='credentials', protocol=protocol, tenant_id='f8cdef31-a31e-4b4a-93e4-5f571e91255a')
            o365_auth.authenticate(scopes=scopes, token_backend=token_backend, account_email=sendr)
            return o365_auth
    except Exception as e:
        print("An error occurred: ", e)
    
def getConfirmation():
        a = input("do you want to continue? (y/n): ")
        if a == 'y':
            return True
        elif a == 'n':
            return False
        else:
            print("invalid input")
            return getConfirmation()
   

def getsendrAndRecievr():
    sendr = input("Enter sender e-mail address: ")
    recievr = input("Enter reciever e-mail address: ")
    return sendr, recievr

def getBySite(site, xlxs):
    df = pd.read_excel(xlxs, index_col=0, sheet_name=None)
    cols =  ['data1', 'data3', 'data4']
    outData = df["Sheet1"].loc[site, cols]._append(df['Sheet2'].loc[site, cols])
    return outData


if __name__ == '__main__':
    print(
    """
__  __  __  ____  __  __ _ _            ___      _            ___      _ _           _             
\ \/ / / / / _\ \/ / / _(_) |_ ___     /   \__ _| |_ __ _    / __\___ | | | ___  ___| |_ ___  _ __ 
 \  / / /  \ \ \  /  \ \| | __/ _ \   / /\ / _` | __/ _` |  / /  / _ \| | |/ _ \/ __| __/ _ \| '__|
 /  \/ /____\ \/  \  _\ \ | ||  __/  / /_// (_| | || (_| | / /__| (_) | | |  __/ (__| || (_) | |   
/_/\_\____/\__/_/\_\ \__/_|\__\___| /___,' \__,_|\__\__,_| \____/\___/|_|_|\___|\___|\__\___/|_|   
                   By Jonas Quintiens                                                                                                                                                                         

    """
    )
    try:
        args = parser.parse_args()
        if not args.site or not args.xlxs:
            parser.print_help()
        else:
            data = getBySite(args.site[0], args.xlxs[0])
            print(data)
            sendr, recievr = getsendrAndRecievr()
            print( 
"""
Mailer name is: %s
reciever is %s
"""
                  %(sendr,recievr))
            while not getConfirmation():
                sendr, recievr = getsendrAndRecievr()

            final_html_data = html_template.format(data.to_html(index=False))
            m = authenticate_to_outlook(sendr)
            if m.is_authenticated:
                m.new_message()
                print("sending mail")
                m.to.add(recievr)
                m.subject = "test"
                m.body = final_html_data
                m.send()
                print("mail sent")
            else:
                print("not authenticated")
            
    except Exception as e:
        print(e)
        parser.print_help()

    
