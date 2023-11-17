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



def getBySite(site, xlxs):
    df = pd.read_excel(xlxs, index_col="site", sheet_name=None)
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
            sendr = input("Enter sender e-mail address: ")
            recievr = input("Enter reciever e-mail address: ")
            print( 
"""
Mailer name is: %s
reciever is %s
"""
                  %(sendr,recievr))
            input("do you want to continue? (y/n): ")
            authenticate_to_gmail(sendr)
            if input == 'y':
                parseMail(data, sendr, recievr)
            else:
                exit(0)
    except Exception as e:
        print(e)
        parser.print_help()

    
