/*
 * Cloud Sheets 2024 C++ Edition - Sample Project
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


#include <iostream>
#include "../../include/cloudsheets.h"
#include <unordered_map>
#include <string>
#include <cstring>

#define BUFF_SIZE 500
bool changesMade = false;

class MyGoogleSheets : public GoogleSheets
{
public:
	virtual int FireSpreadsheetList(GoogleSheetsSpreadsheetListEventParams* e) {
		printf("Id: %s\n", e->Id);
		printf("Name: %s\n\n", e->Name);

		return 0;
	}
};
std::unordered_map<std::string, std::string> parseArgs(int argc, char* argv[]);
char* prompt(char* prompt);
void printMenu();
void toLowerCase(char* str);
void authenticate(const char* client, const char* secret);

const char* oAuthScopes = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive";
const char* oAuthAuthUrl = "https://accounts.google.com/o/oauth2/auth";
const char* oAuthTokenUrl = "https://accounts.google.com/o/oauth2/token";

MyGoogleSheets gsheets;

int main(int argc, char* argv[]) {
	if (argc < 4)
	{
		printf("usage: GoogleSheetsDemo.cpp -c client -s secret\n" \
			" -c The Id of the client assigned when registering the application.\n" \
			" -s The secret value for the client assigned when registering the application.");

		return 0;
	}

	std::unordered_map<std::string, std::string> myArgs = parseArgs(argc, argv);
	char buffer[BUFF_SIZE];

	authenticate(myArgs.at("c").c_str(), myArgs.at("s").c_str());

	printf("Listing spreadsheets...\n");
	gsheets.ListSpreadsheets();

	while (true)
	{
		int retCode = 0;
		printf("\nEnter command: ");
		scanf("%[^\n]", buffer);

		if (!strcmp(buffer, "ls"))
		{
			retCode = gsheets.ListSpreadsheets();
			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}
			else
			{
				printf("Success...");
			}
		}
		else if (!strcmp(buffer, "lssh"))
		{
			char spreadsheetId[BUFF_SIZE];
			printf("Enter the Spreadsheet Id: ");
			getchar();
			scanf("%[^\n]", spreadsheetId);

			retCode = gsheets.LoadSpreadsheet(spreadsheetId);
			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}
			else
			{
				for (int i = 0; i < gsheets.GetSheets()->GetCount(); i++)
				{
					printf("Id: %d\n", gsheets.GetSheets()->Get(i)->GetSheetId());
					printf("Title: %s\n", gsheets.GetSheets()->Get(i)->GetTitle());
					printf("Index: %d\n", gsheets.GetSheets()->Get(i)->GetSheetIndex());
					printf("Total Columns: %d\n", gsheets.GetSheets()->Get(i)->GetTotalColumnCount());
					printf("Total Rows: %d\n\n", gsheets.GetSheets()->Get(i)->GetTotalRowCount());
				}
			}
		}
		else if (!strcmp(buffer, "cs"))
		{
			char name[BUFF_SIZE], locale[BUFF_SIZE], timezone[BUFF_SIZE];

			printf("Enter spreadsheet name: ");
			getchar();
			scanf("%[^\n]", name);

			printf("Enter locale (e.g. en_US): ");
			getchar();
			scanf("%[^\n]", locale);

			printf("Enter timezone (e.g. America/New_York): ");
			getchar();
			scanf("%[^\n]", timezone);
			retCode = gsheets.CreateSpreadsheet(name, locale, timezone);

			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}
			else
			{
				printf("Success...");
			}
		}
		else if (!strcmp(buffer, "ash"))
		{
			char spreadsheetId[BUFF_SIZE];
			printf("Enter the Spreadsheet Id: ");
			getchar();
			scanf("%[^\n]", spreadsheetId);

			retCode = gsheets.LoadSpreadsheet(spreadsheetId);
			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}
			else 
			{
				char sheetIndexStr[BUFF_SIZE];
				printf("Enter the index where sheet should be added: ");
				getchar();
				scanf("%[^\n]", sheetIndexStr);
				int sheetIndex = atoi(sheetIndexStr);

				char sheetTitle[BUFF_SIZE];
				printf("Enter sheet title: ");
				getchar();
				scanf("%[^\n]", sheetTitle);

				retCode = gsheets.AddSheet(sheetIndex, sheetTitle);
				if (retCode != 0)
				{
					printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
				}
				else
				{
					printf("Success...");
				}
			}

		}
		else if (!strcmp(buffer, "us"))
		{
			changesMade = false;
			char spreadsheetId[BUFF_SIZE];
			printf("Enter the Spreadsheet Id: ");
			getchar();
			scanf("%[^\n]", spreadsheetId);

			retCode = gsheets.LoadSpreadsheet(spreadsheetId);
			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}
			else
			{
				CloudSpreadsheetsGSSpreadsheet* spreadsheet = gsheets.GetSpreadsheets()->Get(0);

				printf("Current name: %s\n", spreadsheet->GetName());
				if (strcmp(prompt("Change? (y/n): "), "y") == 0)
				{
					const char* newName = prompt("Enter new name: ");

					gsheets.GetSpreadsheets()->Get(0)->SetName(newName);
					changesMade = true;
				}

				printf("Current locale: %s\n", spreadsheet->GetLocale());
				if (strcmp(prompt("Change? (y/n): "), "y") == 0)
				{
					const char* newLocale = prompt("Enter new locale (e.g. en_US): ");
					gsheets.GetSpreadsheets()->Get(0)->SetLocale(newLocale);
					changesMade = true;
				}

				printf("Current timezone: %s\n", spreadsheet->GetTimeZone());
				if (strcmp(prompt("Change? (y/n): "), "y") == 0)
				{
					const char* newTimeZone = prompt("Enter new timezone (e.g. America/New_York): ");
					gsheets.GetSpreadsheets()->Get(0)->SetTimeZone(newTimeZone);
					changesMade = true;
				}

				if (changesMade)
				{
					retCode = gsheets.UpdateSpreadsheet();
					if (retCode != 0)
					{
						printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
					}
					else
					{
						printf("Success... Spreadsheet updated!\n");
					}
					changesMade = false;
				}
				else
				{
					printf("No changes were made to the spreadsheet.\n");
				}
			}
		}
		else if (!strcmp(buffer, "ush"))
		{
			changesMade = false;
			char currentSpreadsheetId[BUFF_SIZE];
			printf("Enter Spreadsheet Id: ");
			getchar();
			scanf("%[^\n]", currentSpreadsheetId);

			char currentSheetIdStr[BUFF_SIZE];
			printf("Enter Sheet Id (Should be in Spreadsheet with Id: %s): ", currentSpreadsheetId);
			getchar();
			scanf("%[^\n]", currentSheetIdStr);
			int currentSheetId = atoi(currentSheetIdStr);

			retCode = gsheets.LoadSpreadsheet(currentSpreadsheetId);
			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}

			retCode = gsheets.LoadSheet(currentSheetId);
			
			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}
			else 
			{
				CloudSpreadsheetsGSSheet* sheet = gsheets.GetSheets()->Get(0);

				printf("Current name: %s\n", sheet->GetTitle());
				if (strcmp(prompt("Change? (y/n): "), "y") == 0)
				{
					const char* newTitle = prompt("Enter new title: ");
					gsheets.GetSheets()->Get(0)->SetTitle(newTitle);
					changesMade = true;
				}

				printf("Current index: %d\n", sheet->GetSheetIndex());
				if (strcmp(prompt("Change? (y/n): "), "y") == 0)
				{
					const char* newIndex_input = prompt("Enter new index: ");
					int newIndex = atoi(newIndex_input);
					gsheets.GetSheets()->Get(0)->SetSheetIndex(newIndex);
					changesMade = true;
				}

				if (changesMade)
				{
					retCode = gsheets.UpdateSheet();
					if (retCode != 0)
					{
						printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
					}
					else
					{
						printf("Success... Sheet updated!\n");
					}
					changesMade = false;
				}
				else
				{
					printf("No changes were made to the sheet.\n");
				}
			}
		}
		else if (!strcmp(buffer, "ds"))
		{
			char spreadsheetId[BUFF_SIZE];
			printf("Enter the Spreadsheet Id: ");
			getchar();
			scanf("%[^\n]", spreadsheetId);

			retCode = gsheets.DeleteSpreadsheet(spreadsheetId);
			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}
			else
			{
				printf("Success...");
			}
		}
		else if (!strcmp(buffer, "dsh"))
		{
			char spreadsheetId[BUFF_SIZE];
			printf("Enter the Spreadsheet Id: ");
			getchar();
			scanf("%[^\n]", spreadsheetId);

			retCode = gsheets.LoadSpreadsheet(spreadsheetId);
			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}

			char sheetId_input[BUFF_SIZE];
			printf("Enter the Sheet Id: ");
			getchar();
			scanf("%[^\n]", sheetId_input);
			int sheetId = atoi(sheetId_input);

			gsheets.DeleteSheet(sheetId);
			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}
			else
			{
				printf("Success...");
			}

		}
		else if (!strcmp(buffer, "exp"))
		{
			printf("Supported file extensions: csv, ods, pdf, tsv, xlsx, zip\n");

			char currentSpreadsheetId[BUFF_SIZE];
			printf("Enter Spreadsheet Id: ");
			getchar();
			scanf("%[^\n]", currentSpreadsheetId);

			char fileName[BUFF_SIZE];
			printf("Enter full filename (e.g. myFile.zip): ");
			getchar();
			scanf("%[^\n]", fileName);

			char filePath[BUFF_SIZE];
			printf("Enter save directory: ");
			getchar();
			scanf("%[^\n]", filePath);

			snprintf(filePath, BUFF_SIZE, "%s\\%s", filePath, fileName);
			printf("File will be saved at: %s\n", filePath);

			retCode = gsheets.Export(currentSpreadsheetId, filePath);
			if (retCode != 0)
			{
				printf("Error: [%i] %s\n\n", retCode, gsheets.GetLastError());
			}
			else
			{
				printf("Success... File exported successfully.\n");
			}
		}
		else if (!strcmp(buffer, "q"))
		{
			return 0;
		}
		else if (!strcmp(buffer, "?"))
		{
			printMenu();
		}
		else
		{
			printf("Command not recognized!");
			printMenu();
		}
		getchar();
	}
}

void printMenu()
{
	printf("\r\n?\t-\tHelp\n" \
		"ls\t-\tList Spreadsheets\n" \
		"lssh\t-\tList Sheets\n" \
		"cs\t-\tCreate Spreadsheet\n" \
		"ash\t-\tAdd Sheet\n" \
		"us\t-\tUpdate Spreadsheet\n" \
		"ush\t-\tUpdate Sheet\n" \
		"ds\t-\tDelete Spreadsheet\n" \
		"dsh\t-\tDelete Sheet\n" \
		"exp\t-\tExport Spreadsheet\n" \
		"q\t-\tQuit\n");
}

void authenticate(const char* client, const char* secret)
{
	OAuth oauth;

	oauth.SetClientId(client);
	oauth.SetClientSecret(secret);
	oauth.SetAuthorizationScope(oAuthScopes);
	oauth.SetServerAuthURL(oAuthAuthUrl);
	oauth.SetServerTokenURL(oAuthTokenUrl);

	gsheets.SetAuthorization(oauth.GetAuthorization());
}

void toLowerCase(char* str)
{
	for (size_t i = 0; i < std::strlen(str); ++i) {
		str[i] = std::tolower(static_cast<unsigned char>(str[i]));
	}
}

char* prompt(char* prompt)
{
	static char buffer[BUFF_SIZE];

	printf("%s", prompt);
	getchar();
	scanf("%[^\n]", buffer);

	return buffer;
}


std::unordered_map<std::string, std::string> parseArgs(int argc, char* argv[])
{
	std::unordered_map<std::string, std::string> map;
	for (int i = 0; i < argc; i++)
	{
		if (argv[i][0] == '-')
		{
			if (i + 1 < argc && (argv[i + 1][0] != '-'))
			{
				char* newStr = argv[i] + 1;
				toLowerCase(newStr);
				map[newStr] = argv[i + 1];
				i++;
			}
			else
			{
				char* newStr = argv[i] + 1;
				toLowerCase(newStr);
				map[newStr] = "";
			}
		}
		else
		{
			map[std::to_string(i)] = argv[i];
		}
	}
	return map;
}

