# 
# Cloud Sheets 2024 Python Edition - Sample Project
# 
# This sample project demonstrates the usage of Cloud Sheets in a 
# simple, straightforward way. It is not intended to be a complete 
# application. Error handling and other checks are simplified for clarity.
# 
# www.nsoftware.com/cloudsheets
# 
# This code is subject to the terms and conditions specified in the 
# corresponding product license agreement which outlines the authorized 
# usage and restrictions.
# 

import sys
import string
from cloudsheets import *

input = sys.hexversion < 0x03000000 and raw_input or input


def ensureArg(args, prompt, index):
    if len(args) <= index:
        while len(args) <= index:
            args.append(None)
        args[index] = input(prompt)
    elif args[index] is None:
        args[index] = input(prompt)


def get_arg_value(flag):
    if flag in sys.argv:
        index = sys.argv.index(flag) + 1
        if index < len(sys.argv):
            return sys.argv[index]
    return None 

OAUTH_SCOPES = "offline_access files.readwrite user.read"
OAUTH_AUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
OAUTH_TOKEN_UTL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
OAUTH_CLIENT_ID = get_arg_value("/c")
OAUTH_SECRET_ID = get_arg_value("/s")

oexcel = Office365Excel()
changes_made = False

if len(sys.argv) < 4:
    print("Usage: python office365excel.py /c client /s secret\n" +
            " /c The Id of the client assigned when registering the application.\n" +
            " /s The secret value for the client assigned when registering the application.")
    
def display_menu():
    print("\r\n?\t-\tHelp\n" +
          "lwb\t-\tList Workbooks\n" +
          "lwsh\t-\tList Worksheets\n" +
          "cwb\t-\tCreate Workbook\n" +
          "awsh\t-\tAdd Worksheet\n" +
          "uwsh\t-\tUpdate Worksheet\n" +
          "dwb\t-\tDelete Workbook\n" +
          "dsh\t-\tDelete Worksheet\n" +
          "q\t-\tQuit\n")


def authenticate(client, secret):
    oauth = OAuth()
    oauth.client_profile = 0
    oauth.grant_type = 0
    oauth.client_id = client
    oauth.client_secret = secret
    oauth.authorization_scope = OAUTH_SCOPES
    oauth.server_auth_url = OAUTH_AUTH_URL
    oauth.server_token_url = OAUTH_TOKEN_UTL

    try:
        oexcel.authorization = oauth.get_authorization()
        print("\nAuthorized")
    except CloudSpreadsheetsError as ex:
        print(ex)
        
def workbook_on_workbook_list(e):
    print("ID: " + e.id)
    print("Name: " + e.name)


try:
    oexcel = Office365Excel() 
    oexcel.on_workbook_list = workbook_on_workbook_list
    authenticate(OAUTH_CLIENT_ID, OAUTH_SECRET_ID)
    
    print("\nListing Workbooks...")
    oexcel.list_workbooks()
    oexcel.create_session(True)
    
    display_menu()
    
    while True:
        buffer = input("Enter command: ")
        
        try:
            match buffer:
                case "lwb":
                    oexcel.list_workbooks()

                case "lwsh":
                    current_workbook_id = input("Enter the Workbook Id: ")
                    oexcel.load_workbook(current_workbook_id)
                    
                    for i in range((oexcel.workbooks_count)):
                        print("Id: " + str(oexcel.get_worksheet_worksheet_id(i)))
                        print("Name: " + str(oexcel.get_worksheet_name(i)))
                        print("Position: " + str(oexcel.get_worksheet_position(i)))
                        print("Visibility: " + str(oexcel.get_worksheet_visibility(i)))

                case "cwb":
                    oexcel.create_workbook(input("Enter Workbook Name: "))
                    print("Success...")
                    
                case "awsh":
                    current_workbook_id = input("Enter the Workbook Id: ")
                    oexcel.load_workbook(current_workbook_id)
                    oexcel.add_worksheet(input("Enter worksheet title: "))
                    
                case "uwsh":
                    changes_made = False
                    current_workbook_id = input("Enter the Workbook Id: ")
                    current_worksheet_id = input("Enter the Worksheet Id: ")
                    
                    oexcel.load_workbook(current_workbook_id)
                    oexcel.load_worksheet(current_worksheet_id)
                    loaded_worksheet_name = oexcel.get_worksheet_name(0)
                    loaded_worksheet_position = oexcel.get_worksheet_position(0)
                    
                    new_name = input("Current title: " + loaded_worksheet_name + "\nEnter new title (or press Enter to keep the same): ") or loaded_worksheet_name
                    new_position = int(input("Current position: " + str(loaded_worksheet_position).strip() + "\nEnter new position: ") or loaded_worksheet_position)

                    if new_name != loaded_worksheet_name or new_position != loaded_worksheet_position:
                        oexcel.set_worksheet_name(0, new_name)
                        oexcel.set_worksheet_position(0, new_position)
                        changes_made = True
                        
                    if(changes_made):
                        oexcel.update_worksheet()
                        print("Success...")
                    else:
                        print("Couldn't update Worksheet")
                        
                case "dwb":
                    oexcel.delete_workbook(input("Enter Workbook Id: "))
                    print("Success...")
                    
                case "dsh":
                    current_workbook_id = input("Enter Workbook Id: ")
                    current_worksheet_id = input(f"Enter Worksheet Id (Should be in Workbook with Id '{current_workbook_id}'): ")
                    oexcel.delete_worksheet(current_worksheet_id)
                
                case "?":
                    display_menu()
                    
                case "q":
                    break
                
                case _:
                    print("Not a known command!")
                    
        except CloudSpreadsheetsError as ex:
            print(ex)
            
except CloudSpreadsheetsError as ex:
    print(ex)



