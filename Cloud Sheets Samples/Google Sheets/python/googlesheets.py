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

OAUTH_SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive"
OAUTH_AUTH_URL = "https://accounts.google.com/o/oauth2/auth"
OAUTH_TOKEN_UTL = "https://accounts.google.com/o/oauth2/token"
OAUTH_CLIENT_ID = get_arg_value("/c")
OAUTH_SECRET_ID = get_arg_value("/s")

gsheets = GoogleSheets()
changes_made = False

if len(sys.argv) < 4:
    print("Usage: python googlesheets.py /c client /s secret\n" +
            " /c The Id of the client assigned when registering the application.\n" +
            " /s The secret value for the client assigned when registering the application.")


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
        gsheets.authorization = oauth.get_authorization()
        print("\nAuthorized")
    except CloudSpreadsheetsError as ex:
        print(ex)
        
def display_menu():
    print("\r\n?\t-\tHelp\n" +
          "ls\t-\tList Spreadsheets\n" +
          "lssh\t-\tList Sheets\n" +
          "cs\t-\tCreate Spreadsheet\n" +
          "ash\t-\tAdd Sheet\n" +
          "us\t-\tUpdate Spreadsheet\n" +
          "ush\t-\tUpdate Sheet\n" +
          "ds\t-\tDelete Spreadsheet\n" +
          "dsh\t-\tDelete Sheet\n" +
          "exp\t-\tExport Spreadsheet\n" +
          "q\t-\tQuit\n")
    
def spreadsheet_on_spreadsheet_list(e):
    print("ID: " + e.id)
    print("Name: " + e.name)
    
try: 
    gsheets = GoogleSheets() 
    gsheets.on_spreadsheet_list = spreadsheet_on_spreadsheet_list
    authenticate(OAUTH_CLIENT_ID, OAUTH_SECRET_ID)
    
    print("\nListing Spreadsheets...")
    gsheets.list_spreadsheets()
    
    display_menu()
    
    while True:
        buffer = input("Enter command: ")
        
        try:
            match buffer:
                case "ls":
                    gsheets.list_spreadsheets()
                    
                case "lssh":
                    current_spreadsheet_id = input("Enter the Spreadsheet Id: ")
                    gsheets.load_spreadsheet(current_spreadsheet_id)
                    
                    for i in range((gsheets.sheets_count)):
                        print("Id: " + str(gsheets.get_sheet_sheet_id(i)))
                        print("Title: " + gsheets.get_sheet_title(i))
                        print("Index: " + str(gsheets.get_sheet_sheet_index(i)))
                        print("Total Columns: " + str(gsheets.get_sheet_total_column_count(i)))
                        print("Total Rows: " + str(gsheets.get_sheet_total_row_count(i)))

                case "cs":
                    gsheets.create_spreadsheet(input("Enter Spreadsheet Name: "), input("Enter Locale (optional): "), input("Enter Timezone (optional): "))
                    print("Success...")
                    
                case "ash":
                    current_spreadsheet_id = input("Enter the Spreadsheet Id: ")
                    gsheets.load_spreadsheet(current_spreadsheet_id)
                    gsheets.add_sheet(int(input("Enter the index where sheet should be added: ")), input("Enter the title: "))
                
                case "us":
                    changes_made = False
                    current_spreadsheet_id = input("Enter the Spreadsheet Id: ")
                    gsheets.load_spreadsheet(current_spreadsheet_id)
                    
                    if (input("Current name: " + gsheets.get_spreadsheet_name(0) + "\nChange? y/n: ") or "n") == "y":
                        gsheets.set_spreadsheet_name(0, input("Enter a new name: "))
                        changes_made = True
                    
                    if (input("Current locale: " + gsheets.get_spreadsheet_locale(0) + "\nChange? y/n: ") or "n") == "y":
                        gsheets.set_spreadsheet_name(0, input("Enter a new locale (e.g. en_US): "))
                        changes_made = True
                        
                    if (input("Current timezone: " + gsheets.get_spreadsheet_locale(0) + "\nChange? y/n: ") or "n") == "y":
                        gsheets.set_spreadsheet_name(0, input("Enter a new locale (e.g. America/New_York): "))
                        changes_made = True  
                        
                    if changes_made:
                            gsheets.update_spreadsheet()
                            print("Success...")
                            changes_made = False 
                    else:
                        print("Couldn't update Spreadsheet.")
                        
                case "ush":
                    current_spreadsheet_id = input("Enter Spreadsheet Id: ")
                    current_sheet_id = int(input(f"Enter Sheet Id (Should be in Spreadsheet with Id '{current_spreadsheet_id}'): "))
                    
                    gsheets.load_spreadsheet(current_spreadsheet_id)
                    gsheets.load_sheet(current_sheet_id)
                    loaded_sheet_title = gsheets.get_sheet_title(0)
                    loaded_sheet_index = gsheets.get_sheet_sheet_index(0)
                    
                    new_title = input("Current title: " + loaded_sheet_title + "\nEnter new title (or press Enter to keep the same): ") or loaded_sheet_title
                    new_index = int(input("Current index: " + str(loaded_sheet_index) + "\nEnter new index: ") or loaded_sheet_index)
                
                    if new_title != loaded_sheet_title or new_index != loaded_sheet_index:
                        gsheets.set_sheet_title(0, new_title)
                        gsheets.set_sheet_sheet_index(0, new_index)
                        changes_made = True
                    
                    if (input("Total rows: " + str(gsheets.get_sheet_total_row_count(0)) + "\nChange? y/n: ") or "n") == "y":
                        gsheets.set_sheet_total_row_count(0, int(input("Enter new total row count: ")))
                        changes_made = True
                        
                    if(input("Total columns: " + str(gsheets.get_sheet_total_column_count(0)) + "\nChange? y/n: ") or "n") == "y":
                        gsheets.set_sheet_total_column_count(0, int(input("Enter new total columns count: ")))
                        changes_made = True
                        
                    if changes_made:
                            gsheets.update_sheet()
                            print("Success...")
                            changes_made = False 
                    else:
                        print("Couldn't update Spreadsheet.")
                                    
                case "ds":
                    gsheets.delete_spreadsheet(input("Enter Spreadsheet Id: "))
                    print("Success...")
                
                case "dsh":
                    current_spreadsheet_id = input("Enter Spreadsheet Id: ")
                    current_sheet_id = int(input(f"Enter Sheet Id (Should be in Spreadsheet with Id '{current_spreadsheet_id}'): "))
                    gsheets.delete_sheet(current_sheet_id)
                
                case "exp":
                    print("Supported file extensions: csv, ods, pdf, tsv, xlsx, zip")
                    current_spreadsheet_id = input("\nEnter Spreadsheet Id: ")
                    file_name = input("Enter full filename (e.g. myFile.zip): ")
                    file_path = input("Enter save directory: ") + f"\\{file_name}"
                    
                    print(file_path)
                    gsheets.export(current_spreadsheet_id, file_path)
                    print("Success...")
                                    
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



