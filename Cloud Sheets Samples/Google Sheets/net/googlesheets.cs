/*
 * Cloud Sheets 2024 .NET Edition - Sample Project
 *
 * This sample project demonstrates the usage of Cloud Sheets in a 
 * simple, straightforward way. It is not intended to be a complete 
 * application. Error handling and other checks are simplified for clarity.
 *
 * www.nsoftware.com/cloudsheets
 *
 * This code is subject to the terms and conditions specified in the 
 * corresponding product license agreement which outlines the authorized 
 * usage and restrictions.
 * 
 */

using System;
using System.Collections.Generic;
using nsoftware.CloudSheets;

class googlesheets : ConsoleDemo
{
    private static string oAuthScopes = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive";
    private static string oAuthAuthUrl = "https://accounts.google.com/o/oauth2/auth";
    private static string oAuthTokenUrl = "https://accounts.google.com/o/oauth2/token";
    private static GoogleSheets gsheets;
    private static string currentSpreadsheetId = "";
    private static int currentSheetId = -1;
    private static bool changesMade = false;
    private static GSSheet loadedSheet;
    private static Dictionary<string, string> myArgs;

    public static void Main(string[] args)
    {
        if (args.Length < 4)
        {
            Console.WriteLine("Usage: googlesheets.cs /c client /s secret\n" +
                              " /c The Id of the client assigned when registering the application.\n" +
                              " /s The secret value for the client assigned when registering the application.");
            return;
        }

        try
        {
            string buffer;
            myArgs = ParseArgs(args);
            gsheets = new GoogleSheets();
            gsheets.OnSpreadsheetList += Spreadsheet_OnSpreadsheetList;

            Authenticate(myArgs["c"], myArgs["s"]);

            Console.WriteLine("Listing Spreadsheets...");
            gsheets.ListSpreadsheets();

            DisplayMenu();

            while (true)
            {
                buffer = Prompt("Enter command", "");
                try
                {
                    switch (buffer)
                    {
                        case "ls":
                            gsheets.ListSpreadsheets();
                            break;
                        case "lssh":
                            currentSpreadsheetId = Prompt("Enter the Spreadsheet Id", "");
                            gsheets.LoadSpreadsheet(currentSpreadsheetId);
                            foreach (GSSheet sheet in gsheets.Sheets)
                            {
                                Console.WriteLine("Id: " + sheet.SheetId);
                                Console.WriteLine("Title: " + sheet.Title);
                                Console.WriteLine("Index: " + sheet.SheetIndex);
                                Console.WriteLine("Total Columns: " + sheet.TotalColumnCount);
                                Console.WriteLine("Total Rows: " + sheet.TotalRowCount);
                            }
                            break;
                        case "cs":
                            gsheets.CreateSpreadsheet(Prompt("Enter Spreadsheet name", ""), Prompt("Enter locale (optional", ""), Prompt("Enter timezone (optional)", ""));
                            Console.WriteLine("Success...");
                            break;
                        case "ash":
                            currentSpreadsheetId = Prompt("Enter Spreadsheet Id", "");
                            gsheets.LoadSpreadsheet(currentSpreadsheetId);
                            gsheets.AddSheet(Convert.ToInt32(Prompt("Enter the index where sheet should be added", "")), Prompt("Enter sheet title", ""));
                            break;
                        case "us":
                            changesMade = false;
                            currentSpreadsheetId = Prompt("Enter Spreadsheet Id", "");
                            gsheets.LoadSpreadsheet(currentSpreadsheetId);

                            if (Prompt("Current name: " + gsheets.Spreadsheets[0].Name + "\nChange? y/n", "n") == "y")
                            {
                                gsheets.Spreadsheets[0].Name = Prompt("Enter new name", "");
                                changesMade = true;
                            }
                            if (Prompt("Current locale: " + gsheets.Spreadsheets[0].Locale + "\nChange? y/n", "n") == "y")
                            {
                                gsheets.Spreadsheets[0].Locale = Prompt("Enter new locale (e.g. en_US)", "");
                                changesMade = true;
                            }
                            if (Prompt("Current timezone: " + gsheets.Spreadsheets[0].TimeZone + "\nChange? y/n", "n") == "y")
                            {
                                gsheets.Spreadsheets[0].TimeZone = Prompt("Enter new timezone (e.g. America/New_York)", "");
                                changesMade = true;
                            }

                            if (changesMade)
                            {
                                gsheets.UpdateSpreadsheet();
                                Console.WriteLine("Success...");
                            }
                            else
                            {
                                Console.WriteLine("Couldn't update Spreadsheet");
                            }
                            break;
                        case "ush":
                            changesMade = false;
                            currentSpreadsheetId = Prompt("Enter Spreadsheet Id", "");
                            currentSheetId = Convert.ToInt32(Prompt($"Enter Sheet Id (Should be in Spreadsheet with Id: {currentSpreadsheetId})", ""));

                            gsheets.LoadSpreadsheet(currentSpreadsheetId);
                            gsheets.LoadSheet(currentSheetId);
                            loadedSheet = gsheets.Sheets[0];
                            string newTitle = Prompt("Current title: " + loadedSheet.Title + "Enter new title (or press Enter to keep the same):", loadedSheet.Title);
                            int newIndex = Convert.ToInt32(Prompt("Current index: " + loadedSheet.SheetIndex + "Enter new index:", loadedSheet.SheetIndex.ToString()));

                            if (newTitle != loadedSheet.Title || newIndex != loadedSheet.SheetIndex)
                            {
                                loadedSheet.Title = newTitle;
                                loadedSheet.SheetIndex = newIndex;
                                changesMade = true;
                            }
                            if (Prompt("Total rows: " + loadedSheet.TotalRowCount + "\nChange? y/n", "n") == "y")
                            {
                                loadedSheet.TotalRowCount = Convert.ToInt32(Prompt("Enter new total row count", ""));
                                changesMade = true;
                            }
                            if (Prompt("Total columns: " + loadedSheet.TotalColumnCount + "\nChange? y/n", "n") == "y")
                            {
                                loadedSheet.TotalColumnCount = Convert.ToInt32(Prompt("Enter new total columns count", ""));
                                changesMade = true;
                            }

                            if (changesMade)
                            {
                                gsheets.UpdateSheet();
                                Console.WriteLine("Success...");
                            }
                            else
                            {
                                Console.WriteLine("Couldn't update Sheet");
                            }
                            break;
                        case "ds":
                            gsheets.DeleteSpreadsheet(Prompt("Enter Spreadsheet Id", ""));
                            Console.WriteLine("Success...");
                            break;
                        case "dsh":
                            currentSpreadsheetId = Prompt("Enter Spreadsheet Id", "");
                            currentSheetId = Convert.ToInt32(Prompt($"Enter Sheet Id (Should be in Spreadsheet with Id: {currentSpreadsheetId})", ""));

                            gsheets.DeleteSheet(currentSheetId);
                            break;
                        case "exp":
                            Console.WriteLine("Supported file extensions: csv, ods, pdf, tsv, xlsx, zip");
                            currentSpreadsheetId = Prompt("Enter Spreadsheet Id", "");
                            string fileName = Prompt("Enter full filename (e.g. myFile.zip)", "");
                            string filePath = $"{Prompt("Enter save directory", "")}\\{fileName}";
                            Console.WriteLine(filePath);
                            gsheets.Export(currentSpreadsheetId, filePath);
                            Console.WriteLine("Success...");
                            break;
                        case "?":
                            DisplayMenu();
                            break;
                        case "q":
                            return;
                        default:
                            Console.WriteLine("Not a known command!");
                            break;
                    }
                }
                catch (CloudSpreadsheetsException ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private static void Authenticate(string client, string secret)
    {
        OAuth oauth = new OAuth();
        oauth.ClientProfile = OAuthClientProfiles.ocpApplication;
        oauth.GrantType = OAuthGrantTypes.ogtAuthorizationCode;
        oauth.ClientId = client;
        oauth.ClientSecret = secret;
        oauth.AuthorizationScope = oAuthScopes;
        oauth.ServerAuthURL = oAuthAuthUrl;
        oauth.ServerTokenURL = oAuthTokenUrl;

        try
        {
            gsheets.Authorization = oauth.GetAuthorization();
        }
        catch (CloudSpreadsheetsException ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private static void Spreadsheet_OnSpreadsheetList(object sender, GoogleSheetsSpreadsheetListEventArgs e)
    {
        Console.WriteLine("ID: " + e.Id);
        Console.WriteLine("Name: " + e.Name);
    }

    private static void DisplayMenu()
    {
        Console.WriteLine("\r\n?\t-\tHelp\n" +
          "ls\t-\tList Spreadsheets\n" +
          "lssh\t-\tList Sheets\n" +
          "cs\t-\tCreate Spreadsheet\n" +
          "ash\t-\tAdd Sheet\n" +
          "us\t-\tUpdate Spreadsheet\n" +
          "ush\t-\tUpdate Sheet\n" +
          "ds\t-\tDelete Spreadsheet\n" +
          "dsh\t-\tDelete Sheet\n" +
          "exp\t-\tExport Spreadsheet\n" +
          "q\t-\tQuit\n");
    }
}





class ConsoleDemo
{
  /// <summary>
  /// Takes a list of switch arguments or name-value arguments and turns it into a dictionary.
  /// </summary>
  public static System.Collections.Generic.Dictionary<string, string> ParseArgs(string[] args)
  {
    System.Collections.Generic.Dictionary<string, string> dict = new System.Collections.Generic.Dictionary<string, string>();

    for (int i = 0; i < args.Length; i++)
    {
      // Add a key to the dictionary for each argument.
      if (args[i].StartsWith("/"))
      {
        // If the next argument does NOT start with a "/", then it is a value.
        if (i + 1 < args.Length && !args[i + 1].StartsWith("/"))
        {
          // Save the value and skip the next entry in the list of arguments.
          dict.Add(args[i].ToLower().TrimStart('/'), args[i + 1]);
          i++;
        }
        else
        {
          // If the next argument starts with a "/", then we assume the current one is a switch.
          dict.Add(args[i].ToLower().TrimStart('/'), "");
        }
      }
      else
      {
        // If the argument does not start with a "/", store the argument based on the index.
        dict.Add(i.ToString(), args[i].ToLower());
      }
    }
    return dict;
  }
  /// <summary>
  /// Asks for user input interactively and returns the string response.
  /// </summary>
  public static string Prompt(string prompt, string defaultVal)
  {
    Console.Write(prompt + (defaultVal.Length > 0 ? " [" + defaultVal + "]": "") + ": ");
    string val = Console.ReadLine();
    if (val.Length == 0) val = defaultVal;
    return val;
  }
}