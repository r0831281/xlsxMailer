#!/usr/bin/python3

#app id ff91f935-c699-4f05-8844-0570bd88b673


import argparse
import time
import pandas as pd 
from O365 import Message, MSGraphProtocol, Account, FileSystemTokenBackend



html_template =     """ 
            <html>
            <head>
                <title>Site data</title>
            </head>
            <body>
                    {}
            </body>
            </html>
        """

parser = argparse.ArgumentParser(description='get site data.')
parser.add_argument('site', metavar='-s', type=int, nargs='+',
                    help='site number')
parser.add_argument('xlxs', metavar='-x', type=str, nargs='+', help='xlxs file location')
parser.add_argument('--version', '-V', action='version', version='%(prog)s 1.0')
parser.add_argument('--verbose', '-v', action='count', default=0)
parser.add_help = True


protocol = MSGraphProtocol()
scopes = protocol.get_scopes_for(['basic', 'message_all'])

def authenticate_to_outlook():
    try:
           
            client_secret = ' '
            client_id = 'ff91f935-c699-4f05-8844-0570bd88b673'
            credentials = (client_id, client_secret)
            token_backend = FileSystemTokenBackend(token_path='/tokenStore', token_filename='my_token.txt')
            o365_auth = Account(credentials)
            o365_auth.authenticate(scopes=scopes, token_backend=token_backend)
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
    #sendr = input("Enter sender e-mail address: ")
    recievr = input("Enter reciever e-mail address: ")
    return recievr

def getBySite(site, xlxs):
    df = pd.read_excel(xlxs, index_col=0, sheet_name=None)
    cols =  ['data1', 'data3', 'data4']
    data1 = df["Sheet1"].loc[site, cols]
    data1.dropna(inplace=True)
    try:
        data1 = data1.to_frame()
        data1 = data1.transpose()
    except(Exception):
        print("error")    
    print(data1)
    data2 = df['Sheet2'].loc[site, cols]
    try:
        data2 = data2.to_frame()
        data2 = data2.transpose()
    except(Exception):
        print("error")
    outData = pd.concat([data1, data2], ignore_index=True)
    print(data2)
    return outData


def main():
    print("""
__  __  __  ____  __  __ _ _            ___      _            ___      _ _           _             
\ \/ / / / / _\ \/ / / _(_) |_ ___     /   \__ _| |_ __ _    / __\___ | | | ___  ___| |_ ___  _ __ 
 \  / / /  \ \ \  /  \ \| | __/ _ \   / /\ / _` | __/ _` |  / /  / _ \| | |/ _ \/ __| __/ _ \| '__|
 /  \/ /____\ \/  \  _\ \ | ||  __/  / /_// (_| | || (_| | / /__| (_) | | |  __/ (__| || (_) | |   
/_/\_\____/\__/_/\_\ \__/_|\__\___| /___,' \__,_|\__\__,_| \____/\___/|_|_|\___|\___|\__\___/|_|   
                   By Jonas Quintiens                                                                                                                                                                         
""")
    try:
        args = parser.parse_args()
        if not args.site or not args.xlxs:
            parser.print_help()
        else:
            data = getBySite(args.site[0], args.xlxs[0])
            print(data)
            recievr = getsendrAndRecievr()
            print( 
"""
reciever is %s
"""
                  %(recievr))
            while not getConfirmation():
                recievr = getsendrAndRecievr()
            final_html_data = html_template.format(data.to_html())
            m = authenticate_to_outlook()
            if m.is_authenticated:
                a = m.new_message()
                print("sending mail")
                a.to.add(recievr)
                a.subject = "test"
                a.body = final_html_data
                a.send()
                print("mail sent")
            else:
                print("not authenticated")
            
    except Exception as e:
            print(e)
            parser.print_help()

    


if __name__ == '__main__':
    main()
