/*
 * Cloud Sheets 2024 Java Edition - Sample Project
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
 */

import java.io.*;
import cloudsheets.*;


import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Map;

public class googlesheets extends ConsoleDemo {
    private static final String oAuthScope = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive";
    private static final String oAuthAuthUrl = "https://accounts.google.com/o/oauth2/auth";
    private static final String oAuthTokenUrl = "https://accounts.google.com/o/oauth2/token";
    private static GoogleSheets gsheets;
    public static String currentSpreadsheetId = "";
    public static int currentSheetId = -1;
    public boolean changesMade = false;
    public static GSSheet currentSheet = null;
    private static Map<String, String> myArgs;


    public static void main(String[] args) {
        if(args.length <4){
            System.out.println("usage: GoogleSheetsDemo.java -c client -s secret\n" +
                    " -c The Id of the client assigned when registering the application.\n" +
                    " -s The secret value for the client assigned when registering the application.");
            return;
        }

        try{
            String buffer;
            myArgs = parseArgs(args);
            gsheets = new GoogleSheets();
            gsheets.addGoogleSheetsEventListener(new Events());
            authenticate(myArgs.get("c"), myArgs.get("s"));

            System.out.println("Listing Spreadsheets...");
            gsheets.listSpreadsheets();
            displayMenu();

            while (true){
                buffer = prompt("\nEnter command");

                try{
                    switch (buffer){
                        case "ls":
                            gsheets.listSpreadsheets();
                            break;
                        case "lssh":
                            currentSpreadsheetId = prompt("Enter the Spreadsheet Id");
                            gsheets.loadSpreadsheet(currentSpreadsheetId);
                            for(GSSheet sheet : gsheets.getSheets()){
                                System.out.println("Id: " + sheet.getSheetId());
                                System.out.println("Title: " + sheet.getTitle());
                                System.out.println("Index: " + sheet.getSheetIndex());
                                System.out.println("Total Columns: " + sheet.getTotalColumnCount());
                                System.out.println("Total Rows: " + sheet.getTotalRowCount());
                            }
                            break;
                        case "cs":
                            gsheets.createSpreadsheet(
                                    prompt("Enter Spreadsheet name"),
                                    prompt("Enter locale (optional)"),
                                    prompt("Enter timezone (optional)"));
                            System.out.println("Success...");
                            break;
                        case "ash":
                            currentSpreadsheetId = prompt("Enter Spreadsheet Id");
                            gsheets.loadSpreadsheet(currentSpreadsheetId);
                            gsheets.addSheet(
                                    Integer.parseInt(prompt("Enter the index where sheet should be added")),
                                    prompt("Enter sheet title"));
                            System.out.println("Success...");
                            break;
                        case "us":
                            boolean changesMade = false;
                            currentSpreadsheetId = prompt("Enter Spreadsheet Id");
                            gsheets.loadSpreadsheet(currentSpreadsheetId);
                            GSSpreadsheet spreadsheet = gsheets.getSpreadsheets().get(0);
                            if (prompt("Current name: " + spreadsheet.getName() + "\nChange? y/n: ", "n")
                                    .equalsIgnoreCase("y")) {
                                spreadsheet.setName(prompt("Enter new name"));
                                changesMade = true;
                            }
                            if (prompt("Current locale: " + spreadsheet.getLocale() + "\nChange? y/n", "n")
                                    .equalsIgnoreCase("y")) {
                                spreadsheet.setLocale(prompt("Enter new locale (e.g. en_US)"));
                                changesMade = true;
                            }
                            if (prompt("Current timezone:" + spreadsheet.getTimeZone() + "\nChange? y/n", "n")
                                    .equalsIgnoreCase("y")) {
                                spreadsheet.setTimeZone(prompt("Enter new timezone (e.g. America/New_York)"));
                                changesMade = true;
                            }
                            if (changesMade) {
                                gsheets.updateSpreadsheet();
                                System.out.println("Success...");
                            } else {
                                System.out.println("Couldn't update Spreadsheet");
                            }
                            break;
                        case "ush":
                            currentSpreadsheetId = prompt("Enter Spreadsheet Id");
                            currentSheetId = Integer.parseInt(prompt("Enter Sheet Id (Should be in Spreadsheet with Id: " + currentSpreadsheetId + ")"));
                            gsheets.loadSpreadsheet(currentSpreadsheetId);
                            gsheets.loadSheet(currentSheetId);
                            GSSheet loadedSheet = gsheets.getSheets().get(0);
                            String newTitle = prompt("Current title: " + loadedSheet.getTitle() +
                                    " Enter new title (or press Enter to keep the same): ", loadedSheet.getTitle());
                            int newIndex = Integer.parseInt(prompt("Current index: " + loadedSheet.getSheetIndex() +
                                    " Enter new index: ", String.valueOf(loadedSheet.getSheetIndex())));
                            if (!newTitle.equals(loadedSheet.getTitle()) || newIndex != loadedSheet.getSheetIndex()) {
                                loadedSheet.setTitle(newTitle);
                                loadedSheet.setSheetIndex(newIndex);
                            }
                            boolean sheetChangesMade = false;
                            if (prompt("Total rows: " + loadedSheet.getTotalRowCount() + "\nChange? y/n: ", "n")
                                    .equalsIgnoreCase("y")) {
                                loadedSheet.setTotalRowCount(Integer.parseInt(prompt("Enter new total row count")));
                                sheetChangesMade = true;
                            }
                            if (prompt("Total columns: " + loadedSheet.getTotalColumnCount() + "\nChange? y/n", "n")
                                    .equalsIgnoreCase("y")) {
                                loadedSheet.setTotalColumnCount(Integer.parseInt(prompt("Enter new total columns count")));
                                sheetChangesMade = true;
                            }
                            if (sheetChangesMade) {
                                gsheets.updateSheet();
                                System.out.println("Success...");
                            } else {
                                System.out.println("Couldn't update Sheet");
                            }
                            break;
                        case "ds":
                            gsheets.deleteSpreadsheet(prompt("Enter Spreadsheet Id"));
                            System.out.println("Success...");
                            break;
                        case "dsh":
                            currentSpreadsheetId = prompt("Enter Spreadsheet Id");
                            currentSheetId = Integer.parseInt(prompt("Enter Sheet Id (Should be in Spreadsheet with Id: " + currentSpreadsheetId + ")"));
                            gsheets.deleteSheet(currentSheetId);
                            break;
                        case "exp":
                            System.out.println("Supported file extensions: csv, ods, pdf, tsv, xlsx, zip");
                            currentSpreadsheetId = prompt("Enter Spreadsheet Id");
                            String fileName = prompt("Enter full filename (e.g. myFile.zip)");
                            String filePath = prompt("Enter save directory") + File.separator + fileName;
                            System.out.println(filePath);
                            gsheets.export(currentSpreadsheetId, filePath);
                            System.out.println("Success...");
                            break;
                        case "q":
                            System.exit(0);
                            break;
                        default:
                            System.out.println("Not a known command!");
                            break;
                    }

                }
                catch(CloudSpreadsheetsException ex){
                    System.out.println(ex.getMessage());
                }
            }

        }
        catch(Exception ex){
            System.out.println(ex.getMessage());
        }
    }

    private static void authenticate(String client, String secret) throws Exception{
        OAuth oauth = new OAuth();

        if(client != null && secret != null){
            oauth.setClientId(client);
            oauth.setClientSecret(secret);
            oauth.setAuthorizationScope(oAuthScope);
            oauth.setServerAuthURL(oAuthAuthUrl);
            oauth.setServerTokenURL(oAuthTokenUrl);

            try{
                gsheets.setAuthorization(oauth.getAuthorization());
            }
            catch(CloudSpreadsheetsException ex){
                System.out.println(ex.getMessage());
            }
        }
        else {
            System.out.println("Client Id and Secret cannot be null!");
        }
    }

    private static void displayMenu()
    {
        System.out.println("\r\n?\t-\tHelp\n" +
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


class Events implements GoogleSheetsEventListener
{

    @Override
    public void error(GoogleSheetsErrorEvent googleSheetsErrorEvent) {

    }

    @Override
    public void log(GoogleSheetsLogEvent googleSheetsLogEvent) {

    }

    @Override
    public void spreadsheetList(GoogleSheetsSpreadsheetListEvent googleSheetsSpreadsheetListEvent) {
        System.out.println("ID: " + googleSheetsSpreadsheetListEvent.id);
        System.out.println("Name: " + googleSheetsSpreadsheetListEvent.name);
    }

    @Override
    public void SSLServerAuthentication(GoogleSheetsSSLServerAuthenticationEvent googleSheetsSSLServerAuthenticationEvent) {

    }

    @Override
    public void SSLStatus(GoogleSheetsSSLStatusEvent googleSheetsSSLStatusEvent) {

    }

    @Override
    public void transfer(GoogleSheetsTransferEvent googleSheetsTransferEvent) {

    }
}
class ConsoleDemo {
  private static BufferedReader bf = new BufferedReader(new InputStreamReader(System.in));

  static String input() {
    try {
      return bf.readLine();
    } catch (IOException ioe) {
      return "";
    }
  }
  static char read() {
    return input().charAt(0);
  }

  static String prompt(String label) {
    return prompt(label, ":");
  }
  static String prompt(String label, String punctuation) {
    System.out.print(label + punctuation + " ");
    return input();
  }
  static String prompt(String label, String punctuation, String defaultVal) {
      System.out.print(label + " [" + defaultVal + "]" + punctuation + " ");
      String response = input();
      if (response.equals(""))
        return defaultVal;
      else
        return response;
  }

  static char ask(String label) {
    return ask(label, "?");
  }
  static char ask(String label, String punctuation) {
    return ask(label, punctuation, "(y/n)");
  }
  static char ask(String label, String punctuation, String answers) {
    System.out.print(label + punctuation + " " + answers + " ");
    return Character.toLowerCase(read());
  }

  static void displayError(Exception e) {
    System.out.print("Error");
    if (e instanceof CloudSheetsException) {
      System.out.print(" (" + ((CloudSheetsException) e).getCode() + ")");
    }
    System.out.println(": " + e.getMessage());
    e.printStackTrace();
  }

  /**
   * Takes a list of switch arguments or name-value arguments and turns it into a map.
   */
  static java.util.Map<String, String> parseArgs(String[] args) {
    java.util.Map<String, String> map = new java.util.HashMap<String, String>();
    
    for (int i = 0; i < args.length; i++) {
      // Add a key to the map for each argument.
      if (args[i].startsWith("-")) {
        // If the next argument does NOT start with a "-" then it is a value.
        if (i + 1 < args.length && !args[i + 1].startsWith("-")) {
          // Save the value and skip the next entry in the list of arguments.
          map.put(args[i].toLowerCase().replaceFirst("^-+", ""), args[i + 1]);
          i++;
        } else {
          // If the next argument starts with a "-", then we assume the current one is a switch.
          map.put(args[i].toLowerCase().replaceFirst("^-+", ""), "");
        }
      } else {
        // If the argument does not start with a "-", store the argument based on the index.
        map.put(Integer.toString(i), args[i].toLowerCase());
      }
    }
    return map;
  }
}



