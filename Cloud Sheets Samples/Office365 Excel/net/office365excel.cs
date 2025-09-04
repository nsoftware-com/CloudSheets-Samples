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

﻿using System;
using System.Collections.Generic;
using System.IO;
using nsoftware.CloudSheets;

class office365excel : ConsoleDemo
{
    private static string oAuthScopes = "offline_access files.readwrite user.read";
    private static string oAuthAuthUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
    private static string oAuthTokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
    private static Office365Excel oexcel;
    private static string currentWorkbookId = "";
    private static string currentWorksheetId = "";
    private static bool changesMade = false;
    private static OEWorksheet loadedWorksheet;
    private static Dictionary<string, string> myArgs;

    public static void Main(string[] args)
    {
        if (args.Length < 4)
        {
            Console.WriteLine("Usage: office365excel.cs /c client /s secret\n" +
                              " /c The Id of the client assigned when registering the application.\n" +
                              " /s The secret value for the client assigned when registering the application.");
            return;
        }

        try
        {
            string buffer;
            myArgs = ParseArgs(args);
            oexcel = new Office365Excel();
            oexcel.OnWorkbookList += Workbook_OnWorkbookList;

            Authenticate(myArgs["c"], myArgs["s"]);

            Console.WriteLine("Listing Workbooks...");
            oexcel.ListWorkbooks();
            oexcel.CreateSession(true);

            DisplayMenu();

            while (true)
            {
                buffer = Prompt("Enter command", "");
                try
                {
                    switch (buffer)
                    {
                        case "lwb":
                            oexcel.ListWorkbooks();
                            break;
                        case "lwsh":
                            currentWorkbookId = Prompt("Enter the Workbook Id", "");
                            oexcel.LoadWorkbook(currentWorkbookId);
                            foreach (OEWorksheet wsheet in oexcel.Worksheets)
                            {
                                Console.WriteLine("Id: " + wsheet.WorksheetId);
                                Console.WriteLine("Name: " + wsheet.Name);
                                Console.WriteLine("Position: " + wsheet.Position);
                                Console.WriteLine("Visibility: " + wsheet.Visibility.ToString());
                            }
                            break;
                        case "cwb":
                            oexcel.CreateWorkbook(Prompt("Enter Workbook name", ""));
                            Console.WriteLine("Success...");
                            break;
                        case "awsh":
                            currentWorkbookId = Prompt("Enter Workbook Id", "");
                            oexcel.LoadWorkbook(currentWorkbookId);
                            oexcel.AddWorksheet(Prompt("Enter worksheet title", ""));
                            break;
                        case "uwsh":
                            currentWorkbookId = Prompt("Enter Workbook Id", "");
                            currentWorksheetId = Prompt($"Enter Worksheet Id (Should be in Workbook with Id: {currentWorkbookId})", "");

                            oexcel.LoadWorkbook(currentWorkbookId);
                            oexcel.LoadWorksheet(currentWorksheetId);
                            loadedWorksheet = oexcel.Worksheets[0];
                            string newName = Prompt("Current name: " + loadedWorksheet.Name + "Enter new title (or press Enter to keep the same):", loadedWorksheet.Name);
                            int newPosition = Convert.ToInt32(Prompt("Current position: " + loadedWorksheet.Position + "Enter new position:", loadedWorksheet.Position.ToString()));

                            if (newName != loadedWorksheet.Name || newPosition != loadedWorksheet.Position)
                            {
                                loadedWorksheet.Name = newName;
                                loadedWorksheet.Position = newPosition;
                            }

                            if (changesMade)
                            {
                                oexcel.UpdateWorksheet();
                                Console.WriteLine("Success...");
                            }
                            else
                            {
                                Console.WriteLine("Couldn't update Worksheet");
                            }
                            break;
                        case "dwb":
                            oexcel.DeleteWorkbook(Prompt("Enter Workbook Id", ""));
                            Console.WriteLine("Success...");
                            break;
                        case "dsh":
                            currentWorkbookId = Prompt("Enter Workbook Id", "");
                            currentWorksheetId = Prompt($"Enter Worksheet Id (Should be in Workbook with Id: {currentWorkbookId})", "");

                            oexcel.DeleteWorksheet(currentWorksheetId);
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
            oexcel.Authorization = oauth.GetAuthorization();
        }
        catch (CloudSpreadsheetsException ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private static void Workbook_OnWorkbookList(object sender, Office365ExcelWorkbookListEventArgs e)
    {
        Console.WriteLine("ID: " + e.Id);
        Console.WriteLine("Name: " + e.Name);
    }

    private static void DisplayMenu()
    {
        Console.WriteLine("\r\n?\t-\tHelp\n" +
          "lwb\t-\tList Workbooks\n" +
          "lwsh\t-\tList Worksheets\n" +
          "cwb\t-\tCreate Workbook\n" +
          "awsh\t-\tAdd Worksheet\n" +
          "uwsh\t-\tUpdate Worksheet\n" +
          "dwb\t-\tDelete Workbook\n" +
          "dsh\t-\tDelete Worksheet\n" +
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