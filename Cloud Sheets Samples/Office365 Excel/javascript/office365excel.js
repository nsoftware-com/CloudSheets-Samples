/*
 * Cloud Sheets 2024 JavaScript Edition - Sample Project
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
 
const readline = require("readline");
const cloudsheets = require("@nsoftware/cloudsheets");

if(!cloudsheets) {
  console.error("Cannot find cloudsheets.");
  process.exit(1);
}
let rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});


class Office365ExcelDemo {
    constructor() {
        this.oAuthScope = "offline_access files.readwrite user.read";
        this.oAuthAuthUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
        this.oAuthTokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        this.excel = null;
        this.currentWorkbookId = "";
        this.selectedWorkbookName="";
        this.currentSheetId = -1;
        this.changesMade = false;
        this.currentSheet = null;
        this.rl = rl;

        this.running = true;
    }

    askQuestion(query) {
        return new Promise(resolve => {
            this.rl.question(query, resolve);
        });
    }
    reset = async () => 
    {
        this.currentWorkbookId="";
        await this.excel.listWorkbooks();
    }
    normalizeArg = (arg) => 
    {
        if(arg)
        {
            return `${arg}${arg.endsWith('.xlsx') ? '' : '.xlsx'}`;
        }
    }
    getNumberInput= async (prompt) => {
        const input = (await this.askQuestion(prompt)).trim();
        const num = parseInt(input, 10);
        return (isNaN(num) || num < 0) ? null : num;
    }

    getSheetId = async(sheetName) => 
    {
        if(!sheetName)
        {
            sheetName = await this.askQuestion("Sheet Name:")
        }
        const worksheets = this.excel.getWorksheets();
        for (let i = 0; i < worksheets.size(); i++) {
            const sheet = worksheets.get(i);
            if (sheet.getName() === sheetName) return sheet.getWorksheetId();
        }
        return null; // Not found
    }
    
    getConfirmation = async(question) =>
    {
        const answer = await this.askQuestion(question);
        return answer.toLowerCase() === 'y';
    }
    getCommand = async() => 
    {
        const prompt = this.currentWorkbookId 
        ? `${this.selectedWorkbookName}>`
        : 'Excel> ';
        return await this.askQuestion(prompt);
    }
    
    selectWorkbook = async(WorkbookName) => 
    {
        if (!WorkbookName.endsWith('.xlsx'))
        {
            WorkbookName+='.xlsx'
        }
        try {
            await this.loadWorkbook(WorkbookName);
            this.displayWorksheetMenu();
            
        }catch(error)
        {
            console.error("Error selecting Workbook:", error);
        }
    }
    loadWorkbook = async(WorkbookName) =>
    {
        if(!WorkbookName)
        {
            WorkbookName = await this.askQuestion("Enter the Workbook Name: ");
        }
        this.currentWorkbookId = await this.getWorkbookId(WorkbookName);
        if(this.currentWorkbookId != null)
        {
            await this.excel.loadWorkbook(this.currentWorkbookId);
            this.selectedWorkbookName = WorkbookName
        }
        else
        {
            await this.loadWorkbook()
        }
    }
    displayWorksheetMenu() {
        if (!this.selectedWorkbookName ==="") {
            console.log("No Worksheet selected.");
            return;
        }
        console.log(
            "\nls\t-\tList Worksheets\n" +
            "add\t-\tAdd Worksheet\n" +
            "update\t-\tUpdate Worksheet\n" +
            "delete\t-\tDelete Worksheet\n" +
            "exit\t-\tReturn to Main Menu\n"
        );
    }
    displayMainMenu() {
        console.log(
                "?\t-\tHelp\n" +
                "ls\t-\tList Workbook\n" +
                "select\t-\tSelect Workbook\n" +
                "new\t-\tCreate Workbook\n" +
                "del\t-\tDelete Workbook\n" +
                "q\t-\tQuit\n"
        );
    }
    listSheets = async() => 
    {
        const sheets = this.excel.getWorksheets()
        for (let i = 0; i < sheets.size(); i++) {
            console.log("Title: " + sheets.get(i).getName());
        }
    }

    getWorkbookId = async (WorkbookName) =>
    {
        const Workbooks = this.excel.getWorkbooks();
        for (let i = 0; i < Workbooks.size(); i++) {
            const name = Workbooks.get(i).getName();
            if(name === WorkbookName)
            {
                return Workbooks.get(i).getId();
            }
        }
        if (!this.currentWorkbookId) {
            console.log(`No Workbook found with name: ${WorkbookName}`);
            return null;
        }
    };
    listWorkbooks = async () => {
        if(this.currentWorkbookId != "")
        {
            await this.listSheets();
            return;
        }
        await this.excel.listWorkbooks();
        const Workbooks = this.excel.getWorkbooks();
        const maxNameLength = 24;

        console.log("Name".padEnd(maxNameLength));
        console.log("-".repeat(20)); 

        for (let i = 0; i < Workbooks.size(); i++) {
            const name = Workbooks.get(i).getName();
            console.log(`${name}`);
        }
    };
    createWorkbook = async(WorkbookName) => 
    {
        if(!WorkbookName)
        {
            WorkbookName = await this.askQuestion("Enter Workbook name: ");
        }
        await this.excel.createWorkbook(WorkbookName);
        console.log(`Workbook ${WorkbookName} created successfully.`)
        this.reset()
    };
    addSheet = async(sheetTitle) => 
    {
        if(this.selectedWorkbookName === "")
        {
            console.log("A Workbook must first be selected")
            return;
        }
        if(!sheetTitle)
        {
            sheetTitle = await this.askQuestion("Enter Sheet title: ");
        }
        await this.excel.addWorksheet(sheetTitle);
        console.log("Sheet added successfully!")
    };

    updateWorksheet = async(sheetTitle) => 
    {
        if(this.selectedWorkbookName === "")
        {
            console.log("A Workbook must first be selected")
            return;
        }
        if(!sheetTitle)
        {
            sheetTitle = await this.askQuestion("Enter Sheet title: ");
        }

        const sheetId = await this.getSheetId(sheetTitle);

        await this.excel.loadWorksheet(sheetId); 
        if(await this.getConfirmation("Change sheet name (Y/N)? "))
        {
            const newName = await this.askQuestion("Enter new sheet name: ")
            this.excel.getWorksheets().get(0).setName(newName);
            await this.excel.updateWorksheet();
        }
    };

    deleteSheet = async(sheetTitle) => 
    {
        if(this.selectedWorkbookName === "")
        {
            console.log("A Workbook must first be selected")
            return;
        }
        if(!sheetTitle)
        {
            sheetTitle = await this.askQuestion("Enter Sheet title: ");
        }
        const sheetId = await this.getSheetId(sheetTitle);

        await this.excel.loadWorksheet(sheetId);
        if(await this.getConfirmation("Delete Sheet (Y/N)? "))
        {
            await this.excel.deleteWorksheet(sheetId);
        }
    }
    deleteWorkbook = async(WorkbookName) => 
    {  
        await this.excel.listWorkbooks();
        await this.loadWorkbook(WorkbookName);
        await this.excel.deleteWorkbook(this.currentWorkbookId)
        await this.reset()
    }

    exit = async()=>
    {
        this.currentWorkbookId ="";
        this.currentSheetId="";
        await this.excel.listWorkbooks();
    }
    async main(args) {
        if (args.length < 4) {
            console.log("usage: node googlesheets.js -c client -s secret\n" +
                " -c The Id of the client assigned when registering the application.\n" +
                " -s The secret value for the client assigned when registering the application.");
            return;
        }

        try {
            const myArgs = this.parseArgs(args);
            this.excel = new cloudsheets.office365excel();


            await this.authenticate(myArgs.get("c"), myArgs.get("s"));
            console.log("Loading Workbooks...");
            await this.excel.listWorkbooks();
            this.currentWorkbookId ? this.displayWorksheetMenu() : this.displayMainMenu();

            while (this.running) {
                try {
                    const input= await this.getCommand();
                    const [command, ...args] = input.trim().split(/\s+/);

                    // Process commands
                    switch (command.toLowerCase().trim()) {
                        case 'ls':
                            await this.listWorkbooks();
                            break;
                        case '?':
                            this.currentWorkbookId ? this.displayWorksheetMenu() : this.displayMainMenu();
                            break;
                        case 'select':
                            await this.selectWorkbook(this.normalizeArg(args[0]));
                            break;
                        case 'new':
                            await this.createWorkbook(args[0]);
                            break;
                        case 'add': 
                            await this.addSheet(args[0]);
                            break;
                        case 'us': 
                            await this.updateWorkbook(args[0]);
                            break;
                        case 'update': 
                            await this.updateWorksheet(args[0]);
                            break;
                        case 'del': 
                            await this.deleteWorkbook(this.normalizeArg(args[0]));
                            break;
                        case 'delete': 
                            await this.deleteSheet(args[0]);
                            break;
                        case 'exit': 
                            await this.exit();
                            break;
                        case 'q':
                            console.log('Goodbye!');
                            this.rl.close();
                            this.running = false;
                            return;
                        default:
                            console.log('Invalid command. Please try again.');
                    }
                } catch (error) {
                    console.error('An error occurred:', error.message);
                    break;
                }
            }

        } catch (ex) {
            console.log(ex.message);
        }

    }

    async authenticate(client, secret) {
        if (client && secret) {
            const oauth = new cloudsheets.oauth();
            oauth.setClientId(client)
            oauth.setClientSecret(secret);
            oauth.setAuthorizationScope(this.oAuthScope);
            oauth.setServerAuthURL(this.oAuthAuthUrl);
            oauth.setServerTokenURL(this.oAuthTokenUrl);

            try {
                let authorization = await oauth.getAuthorization()
                await this.excel.setAuthorization(authorization);
            } catch (ex) {
                console.log(ex.message);
            }
        } else {
            console.log("Client Id and Secret cannot be null!");
        }
    }
    parseArgs(args) {
        const myArgs = new Map();
        for (let i = 0; i < args.length; i += 2) {
            if (args[i].startsWith('-')) {
                myArgs.set(args[i].substring(1), args[i + 1]);
            }
        }
        return myArgs;
    }
}

// Run the demo
const demo = new Office365ExcelDemo();
demo.main(process.argv.slice(2));

function prompt(promptName, label, punctuation, defaultVal)
{
  lastPrompt = promptName;
  lastDefault = defaultVal;
  process.stdout.write(`${label} [${defaultVal}]${punctuation} `);
}
