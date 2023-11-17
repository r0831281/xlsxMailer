#!/usr/bin/python3

import argparse
import pandas as pd
import smtplib
import getpass

parser = argparse.ArgumentParser(description='get site data.')
parser.add_argument('site', metavar='-s', type=int, nargs='+',
                    help='site name')
parser.add_argument('xlxs', metavar='-x', type=str, nargs='+', help='xlxs file location')
parser.add_argument('--version', '-V', action='version', version='%(prog)s 1.0')
parser.add_argument('--verbose', '-v', action='count', default=0)
parser.add_help = True


def authenticate_to_outlook(sendr):
    try:
        email = sendr
        password = getpass.getpass("Enter your password: ")

        server = smtplib.SMTP('smtp-mail.outlook.com', 587)
        server.starttls()
        server.login(email, password)

        return server
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

    

def getsendrAndRecievr():
    sendr = input("Enter sender e-mail address: ")
    recievr = input("Enter reciever e-mail address: ")
    return sendr, recievr

def getBySite(site, xlxs):
    df = pd.read_excel(xlxs, index_col=0, sheet_name=None)
    cols =  ['data1', 'data3', 'data4']
    outData = df["Sheet1"].loc[site, cols]._append(df['Sheet2'].loc[site, cols])
    return outData

def parseMail(data, sendr, recievr):
    return data


if __name__ == '__main__':
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
            server = authenticate_to_outlook(sendr)
            print("sending mail")
            data = parseMail(data, sendr, recievr)



    except Exception as e:
        print(e)
        parser.print_help()

    
