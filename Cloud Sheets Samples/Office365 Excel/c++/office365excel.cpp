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
#include <string>
#include <vector>
#include <algorithm>
#include <sstream>
#include <cctype>
#include "../../include/cloudsheets.h"

class Office365Demo {
private:
    bool running;
    char* current_workbook_id;
    bool isWorkbook = false;
    std::string promptPrefix = "Excel";
    Office365Excel excel;

    // Utility methods
    static std::string trim(const std::string& str) {
        auto start = str.begin();
        while (start != str.end() && std::isspace(*start)) {
            start++;
        }

        auto end = str.end();
        do {
            end--;
        } while (std::distance(str.begin(), end) > 0 && std::isspace(*end));

        return std::string(start, end + 1);
    }

    static std::string to_lowercase(std::string str) {
        std::transform(str.begin(), str.end(), str.begin(),
            [](unsigned char c) { return std::tolower(c); });
        return str;
    }
    bool isEmpty(const char* str) const {
        return (str == nullptr || str[0] == '\0');
    }

    // Parse input into command and arguments
    std::vector<std::string> parse_input(const std::string& input) {
        std::vector<std::string> tokens;
        std::istringstream iss(trim(input));
        std::string token;

        while (iss >> token) {
            tokens.push_back(token);
        }

        return tokens;
    }

    void authenticate(const char* client, const char* secret)
    {
        OAuth oauth;
        const char* oAuthScope = "offline_access files.readwrite user.read";
        const char* oAuthAuthUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
        const char* oAuthTokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        oauth.SetClientId(client);
        oauth.SetClientSecret(secret);
        oauth.SetAuthorizationScope(oAuthScope);
        oauth.SetServerAuthURL(oAuthAuthUrl);
        oauth.SetServerTokenURL(oAuthTokenUrl);
        excel.SetAuthorization(oauth.GetAuthorization());

    }

    // Command implementations
    void listWorkbooks() {
        excel.ListWorkbooks();
        const int bookCount = excel.GetWorkbooks()->GetCount();

        std::cout << std::string(20, '-') << std::endl;
        for (int i = 0; i < bookCount; i++)
        {
            std::cout << excel.GetWorkbooks()->Get(i)->GetName() << std::endl;
        }
    }

    char* getWorkbookIdByName(const std::string& workbook_name)
    {
        excel.ListWorkbooks();
        const int bookCount = excel.GetWorkbooks()->GetCount();

        for (int i = 0; i < bookCount; i++)
        {
            if (excel.GetWorkbooks()->Get(i)->GetName() == workbook_name)
            {
                return excel.GetWorkbookId(i);
           }
        }
        return nullptr;
    }

    void displayMenu() {
        if (!isWorkbook) {
            std::cout << "h                 -  Help\n";
            std::cout << "ls                -  List workbooks\n";
            std::cout << "select [workbook] -  Select a workbook\n";
            std::cout << "new    [name]     -  Create new Workbook\n";
            std::cout << "delete [name]     -  Delete Workbook\n";
            std::cout << "q                 -  Quit\n";
         }
        else {
            std::cout << "ls                - List sheets\n";
            std::cout << "add    [sheet]    - Add new sheet\n";
            std::cout << "update [sheet]    - Update sheet\n";
            std::cout << "delete [sheet]    - Delete sheet\n";
            std::cout << "exit              - Return to Main Menu\n";
        }
    }

    void selectWorkbook(const std::string& workbook_name) {
        current_workbook_id = getWorkbookIdByName(workbook_name);
        if (isEmpty(current_workbook_id))
        {
            std::cout << "Workbook not found, please try again." << std::endl;
            return;
        }
        excel.LoadWorkbook(current_workbook_id);
        isWorkbook = true; 
        promptPrefix = workbook_name;
        std::cout << "Selected workbook: " << workbook_name << std::endl;
        displayMenu();
    }

    void createWorkbook(const std::string& workbook_name) {

        excel.CreateWorkbook(workbook_name.c_str());
        std::cout << "Created workbook: " << workbook_name << std::endl;
    }

    void addSheet(const std::string& sheet_name) {
        excel.AddWorksheet(sheet_name.c_str());
        std::cout << "Added sheet: " << sheet_name << std::endl;
        
    }
    void listSheets()
    {
        for (int i = 0; i < excel.GetWorksheets()->GetCount(); i++)
        {
            std::cout << excel.GetWorksheets()->Get(i)->GetName() << std::endl;
        }                 
    }
    const char* getSheetIdByName(const std::string& sheetName)
    {
        if (isWorkbook)
        {
            for (int i = 0; i < excel.GetWorksheets()->GetCount(); i++)
            {
                auto* found = excel.GetWorksheets()->Get(i)->GetName();
                if (excel.GetWorksheets()->Get(i)->GetName() == sheetName)
                {
                    return excel.GetWorksheets()->Get(i)->GetWorksheetId();
                }
            }
            return nullptr;
        }
    }

    void updateWorkbook(const std::string& workbook_id) {
        std::cout << "Updating workbook: " << workbook_id << std::endl;
    }

    std::string normalizeArg(const std::string& arg) {
        if (!arg.empty()) {
            // Check if the string already ends with .xlsx
            if (arg.size() >= 5 &&
                arg.compare(arg.size() - 5, 5, ".xlsx") == 0) {
                return arg;
            }
            // Append .xlsx if not already present
            return arg + ".xlsx";
        }
        // Return empty string for empty input 
        return "";
    }

    void updateWorksheet(const std::string& worksheet_name) {

        if (isWorkbook)
        {
            const char* sheetId = getSheetIdByName(worksheet_name.c_str());
            if (sheetId == nullptr)
            {
                std::cout << "The selected sheet was not found in the workbook\n";
                return;
            }
            std::string newName;
            std::cout << "Enter new sheet name: ";
            std::getline(std::cin, newName); 
            excel.LoadWorksheet(sheetId);
            excel.GetWorksheets()->Get(0)->SetName(newName.c_str());
            excel.UpdateWorksheet();
            std::cout << "Sheet updated successfully! " << std::endl;

        }
    }

    void deleteWorkbook(const std::string& workbook_name) {

        char* workbook_id = getWorkbookIdByName(workbook_name);
        if (!isEmpty(workbook_id))
        {
            
            excel.DeleteWorkbook(workbook_id);
            if (current_workbook_id == workbook_id) {
                current_workbook_id[0] = '\n';
            }
            std::cout << "Workbook " << workbook_name <<" was deleted successfully!" << std::endl;
            return;
        }

        std::cout << "Workbook " << workbook_name <<" was not found!" << std::endl;
    }

    void deleteSheet(const std::string& sheet_name) {
        const char* sheetId = getSheetIdByName(sheet_name);
        if (sheetId ==nullptr)
        {
            std::cout << "The selected sheet was not found in the workbook";
            return;
        }
        excel.LoadWorksheet(sheetId);
        excel.DeleteWorksheet(sheetId);
        std::cout << "Deleted sheet: " << sheet_name << std::endl;
    }

    // Process command
    void processCommand(const std::string& input) {
        if (input.empty()) return;
        auto args = parse_input(input);

        if (args.empty()) return;

        std::string command = to_lowercase(args[0]);

        if (command == "ls") {
            if (isWorkbook)
            {
                listSheets();
                return;
            }
            listWorkbooks();
        }
        else if (command == "?" || command == "h") {
            displayMenu();
        }
        else if (command == "select" && args.size() > 1) {
            selectWorkbook(normalizeArg(args[1]));
        }
        else if (command == "new") {
            if (args.size() < 2) {
                std::cout << "Please provide a workbook name.\n";
                return;
            }
            createWorkbook(args[1]);
        }
        else if (command == "add") {
            if (!isWorkbook)
            {
                std::cout << "Please select a Workbook first \n";
                return;
            }
            if (args.size() < 2) {
                std::cout << "Please provide a sheet name.\n";
                return;
            }
            addSheet(args[1]);
        }
        else if (command == "us" && args.size() > 1) {
            updateWorkbook(args[1]);
        }
        else if (command == "update" && args.size() > 1) {
            updateWorksheet(args[1]);
        }
        else if ((command == "del" || command == "delete") && args.size() > 1) {
            if (!isWorkbook) {
                deleteWorkbook(normalizeArg(args[1]));
                return;
            }
            else {
                deleteSheet(args[1]);
            }
        }
        else if (command == "exit" || command == "q") {
            if (!isWorkbook)
            {
                std::cout << "Goodbye!\n";
                running = false;
            }
            isWorkbook = false;
            promptPrefix = "Excel";

        }
        else {
            std::cout << "Invalid command. Please try again.\n";
        }
    }

public:
    Office365Demo() : running(true) {}

    void run(const char* client, const char* secret) {
        excel.CreateSession(true);

        displayMenu();

        authenticate(client,secret);
        while (running) {
            std::cout << promptPrefix + ">";
            std::string input;
            std::getline(std::cin, input);

            try {
                processCommand(input);
            }
            catch (const std::exception& e) {
                std::cerr << e.what() << std::endl;
            }
        }
    }
};

int main() {
    Office365Demo excel;

    std::string clientId;
    std::string clientSecret;
    std::cout << "Enter the Id of the client assigned when registering the application: ";
    std::getline(std::cin, clientId);
    std::cout << "Enter the secret value for the client assigned when registering the application: ";
    std::getline(std::cin, clientSecret);

    excel.run(clientId.c_str(), clientSecret.c_str());
    return 0;
}

