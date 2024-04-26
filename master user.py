# -*- coding: utf-8 -*-
"""
Created on Thu Jan 18 20:13:14 2024

@author: klecznar
"""

import sys, argparse

from functions import get_credentials, get_expired_OINs, get_selected_OINs, download_certs


def initialize(choice):

    print("What would you like to do today? \n"
          "Press 1 to check all expired certificates \n"
          "Press 2 to only download certificates for selected OIN numbers \n"
          "or press Q to exit. \n"
          "")    
    

    choice = input("My choice: ")
    
    if choice.upper() == 'Q':
        sys.exit()
    elif choice.upper() == '1':
        # get_credentials()
        get_expired_OINs() 
        download_certs()
    elif choice.upper() == '2':
        # get_credentials()
        get_selected_OINs()
        download_certs()
    else:
        print("You've chosen unavailable option! Exiting...")
        sys.exit()





if __name__ == '__main__':

    parser = argparse.ArgumentParser(
                        description='WELCOME TO THE OASIS WEB SCRAPER!')
                                    
    parser.add_argument(
        '-q', '--quit', 
        type=str, 
        help='To quit the program'
        )    
    parser.add_argument(
        '-x', 
        type=str, 
        help='To check for expired AS certificates'
        )     
    parser.add_argument(
        '-z', 
        type=str, 
        help='To check AS certificates as needed'
        ) 
    args = parser.parse_args() 
    # print(args)
    initialize(args)
          
        




