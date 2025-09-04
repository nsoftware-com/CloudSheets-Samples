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
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Map;

import cloudsheets.*;

public class office365excel extends ConsoleDemo {
  private static String oAuthScope = "offline_access files.readwrite user.read";
  private static String oAuthAuthUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
  private static String oAuthTokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
  private static Office365Excel oexcel;
  private static String currentWorkbookId = "";
  private static String currentWorksheetId = "";
  private static boolean changesMade = false;
  private static OEWorksheet loadedWorksheet;
  private static Map<String, String> myArgs;

  public static void main(String[] args) {
    if (args.length < 4) {
      System.out.println("Usage: Office365ExcelDemo.java -c client -s secret\n" +
              " -c The Id of the client assigned when registering the application.\n" +
              " -s The secret value for the client assigned when registering the application.");
      return;
    }

    try {
      String buffer;
      myArgs = parseArgs(args);
      oexcel = new Office365Excel();
      oexcel.addOffice365ExcelEventListener(new Events());

      authenticate(myArgs.get("c"), myArgs.get("s"));

      System.out.println("Listing Workbooks...");
      oexcel.listWorkbooks();
      oexcel.createSession(true);

      displayMenu();

      while (true) {
        buffer = prompt("Enter command", ":");

        try {
          switch (buffer) {
            case "lwb":
              oexcel.listWorkbooks();
              break;
            case "lwsh":
              currentWorkbookId = prompt("Enter the Workbook Id", ":");
              oexcel.loadWorkbook(currentWorkbookId);
              for (OEWorksheet worksheet : oexcel.getWorksheets()) {
                System.out.println("Id: " + worksheet.getWorksheetId());
                System.out.println("Name: " + worksheet.getName());
                System.out.println("Position: " + worksheet.getPosition());
                System.out.println("Visibility: " + worksheet.getVisibility());
              }
              break;
            case "cwb":
              oexcel.createWorkbook(prompt("Enter Workbook name", ":"));
              System.out.println("Success...");
              break;
            case "awsh":
              currentWorkbookId = prompt("Enter Workbook Id", ":");
              oexcel.loadWorkbook(currentWorkbookId);
              oexcel.addWorksheet(prompt("Enter worksheet title", ":"));
              break;
            case "uwsh":
              currentWorkbookId = prompt("Enter Workbook Id", ":");
              currentWorksheetId = prompt("Enter Worksheet Id (Should be in Workbook with Id: " + currentWorkbookId + ")", ":");

              oexcel.loadWorkbook(currentWorkbookId);
              oexcel.loadWorksheet(currentWorksheetId);
              loadedWorksheet = oexcel.getWorksheets().get(0);
              String newName = prompt("Current name: " + loadedWorksheet.getName() +
                      " Enter new title (or press Enter to keep the same):", loadedWorksheet.getName());
              int newPosition = Integer.parseInt(prompt("Current position: " + loadedWorksheet.getPosition() +
                      " Enter new position:", String.valueOf(loadedWorksheet.getPosition())));

              if (!newName.equals(loadedWorksheet.getName()) || newPosition != loadedWorksheet.getPosition()) {
                loadedWorksheet.setName(newName);
                loadedWorksheet.setPosition(newPosition);
              }

              if (changesMade) {
                oexcel.updateWorksheet();
                System.out.println("Success...");
              } else {
                System.out.println("Couldn't update Worksheet");
              }
              break;
            case "dwb":
              oexcel.deleteWorkbook(prompt("Enter Workbook Id", ":"));
              System.out.println("Success...");
              break;
            case "dsh":
              currentWorkbookId = prompt("Enter Workbook Id", ":");
              currentWorksheetId = prompt("Enter Worksheet Id (Should be in Workbook with Id: " + currentWorkbookId + ")", ":");
              oexcel.deleteWorksheet(currentWorksheetId);
              break;
            case "?":
              displayMenu();
              break;
            case "q":
              return;
            default:
              System.out.println("Not a known command!");
              break;
          }
        } catch (CloudSpreadsheetsException ex) {
          System.out.println(ex.getMessage());
        }
      }

    } catch (Exception ex) {
      System.out.println(ex.getMessage());
    }
  }

  private static void authenticate(String client, String secret) throws Exception {
    OAuth oauth = new OAuth();

    if (client != null && secret != null) {
      oauth.setClientId(client);
      oauth.setClientSecret(secret);
      oauth.setAuthorizationScope(oAuthScope);
      oauth.setServerAuthURL(oAuthAuthUrl);
      oauth.setServerTokenURL(oAuthTokenUrl);

      try {
        oexcel.setAuthorization(oauth.getAuthorization());
      } catch (CloudSpreadsheetsException ex) {
        System.out.println(ex.getMessage());
      }
    } else {
      System.out.println("Client Id and Secret cannot be null!");
    }
  }


  private static void displayMenu() {
    System.out.println("\r\n?\t-\tHelp\n" +
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

class Events implements Office365ExcelEventListener
{

  @Override
  public void error(Office365ExcelErrorEvent office365ExcelErrorEvent) {

  }

  @Override
  public void log(Office365ExcelLogEvent office365ExcelLogEvent) {

  }

  @Override
  public void SSLServerAuthentication(Office365ExcelSSLServerAuthenticationEvent office365ExcelSSLServerAuthenticationEvent) {

  }

  @Override
  public void SSLStatus(Office365ExcelSSLStatusEvent office365ExcelSSLStatusEvent) {

  }

  @Override
  public void transfer(Office365ExcelTransferEvent office365ExcelTransferEvent) {

  }

  @Override
  public void workbookList(Office365ExcelWorkbookListEvent office365ExcelWorkbookListEvent) {
    System.out.println("ID: " + office365ExcelWorkbookListEvent.id);
    System.out.println("Name: " + office365ExcelWorkbookListEvent.name);
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



