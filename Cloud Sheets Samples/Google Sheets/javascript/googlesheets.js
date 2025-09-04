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


class GoogleSheetsDemo {
    constructor() {
        this.oAuthScope = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive";
        this.oAuthAuthUrl = "https://accounts.google.com/o/oauth2/auth";
        this.oAuthTokenUrl = "https://accounts.google.com/o/oauth2/token";
        this.gsheets = null;
        this.currentSpreadsheetId = "";
        this.selectedSpreadsheetName="";
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
        this.currentSpreadsheetId="";
        await this.gsheets.listSpreadsheets();
    }
    getNumberInput= async (prompt) => {
        const input = (await this.askQuestion(prompt)).trim();
        const num = parseInt(input, 10);
        return (isNaN(num) || num < 0) ? null : num;
    }
    
    getConfirmation = async(question) =>
    {
        const answer = await this.askQuestion(question);
        return answer.toLowerCase() === 'y';
    }
    getCommand = async() => 
    {
        const prompt = this.currentSpreadsheetId 
        ? `${this.selectedSpreadsheetName}>`
        : 'GoogleSheets> ';
        return await this.askQuestion(prompt);
    }
    
    selectSpreadsheet = async(spreadsheetName) => 
    {
        try {
            await this.loadSpreadsheet(spreadsheetName);
            this.displaySpreadsheetMenu();
            
        }catch(error)
        {
            console.error("Error selecting spreadsheet:", error);
        }
    }
    loadSpreadsheet = async(spreadsheetName) =>
    {
        if(!spreadsheetName)
        {
            spreadsheetName = await this.askQuestion("Enter the Spreadsheet Name: ");
        }
        this.currentSpreadsheetId = await this.getSpreadsheetId(spreadsheetName);
        if(this.currentSpreadsheetId != null)
        {
            await this.gsheets.loadSpreadsheet(this.currentSpreadsheetId);
            this.selectedSpreadsheetName = spreadsheetName
        }
        else
        {
            await this.loadSpreadsheet()
        }
    }
    displaySpreadsheetMenu() {
        if (!this.selectedSpreadsheetName ==="") {
            console.log("No spreadsheet selected.");
            return;
        }
        console.log(
            "\nls\t-\tList Sheets\n" +
            "add\t-\tAdd Sheet\n" +
            "update\t-\tUpdate Sheet\n" +
            "delete\t-\tDelete Sheet\n" +
            "exit\t-\tReturn to Main Menu\n"
        );
    }
    displayMainMenu() {
        console.log(
                "?\t-\tHelp\n" +
                "ls\t-\tList Spreadsheets\n" +
                "sel\t-\tSelect Spreadsheet\n" +
                "cs\t-\tCreate Spreadsheet\n" +
                "us\t-\tUpdate Spreadsheet\n" +
                "del\t-\tDelete Spreadsheet\n" +
                "exp\t-\tExport Spreadsheet\n" +
                "q\t-\tQuit\n"
        );
    }
    listSheets = async() => 
    {
        const sheets = this.gsheets.getSheets()
        for (let i = 0; i < sheets.size(); i++) {
            console.log("\nId: " + sheets.get(i).getSheetId());
            console.log("Title: " + sheets.get(i).getTitle());
            console.log("Index: " + sheets.get(i).getSheetIndex());
            console.log("Total Columns: " + sheets.get(i).getTotalColumnCount());
            console.log("Total Rows: " + sheets.get(i).getTotalRowCount());
        }
    }

    getSpreadsheetId = async (spreadsheetName) =>
    {
        const spreadsheets = this.gsheets.getSpreadsheets();
        for (let i = 0; i < spreadsheets.size(); i++) {
            const name = spreadsheets.get(i).getName();
            if(name === spreadsheetName)
            {
                return spreadsheets.get(i).getId();
            }
        }
        if (!this.currentSpreadsheetId) {
            console.log(`No spreadsheet found with name: ${spreadsheetName}`);
            return null;
        }
    };
    listSpreadsheets = async () => {
        if(this.currentSpreadsheetId != "")
        {
            await this.listSheets();
            return;
        }
        const spreadsheets = this.gsheets.getSpreadsheets();
        const maxNameLength = 24;

        console.log("Name".padEnd(maxNameLength));
        console.log("-".repeat(20)); 

        for (let i = 0; i < spreadsheets.size(); i++) {
            const name = spreadsheets.get(i).getName();
            console.log(`${name}`);
        }
    };
    createSpreadsheet = async(spreadSheetName) => 
    {
        if(!spreadSheetName)
        {
            spreadSheetName = await this.askQuestion("Enter Spreadsheet name: ");
        }
        const locale = await this.askQuestion("Enter Spreadsheet locale (optional): ");
        const timezone = await this.askQuestion("Enter Spreadsheet timezone (optional): ");
        await this.gsheets.createSpreadsheet(spreadSheetName,locale,timezone);
        console.log(`Spreadsheet ${spreadSheetName} created successfully.`)
    };
    addSheet = async(sheetTitle) => 
    {
        if(this.selectedSpreadsheetName === "")
        {
            console.log("A spreadsheet must first be selected")
            return;
        }
        if(!sheetTitle)
        {
            sheetTitle = await this.askQuestion("Enter Sheet title: ");
        }
        await this.gsheets.addSheet(this.gsheets.getSheets().size()+1,sheetTitle);
        console.log("Sheet added successfully!")
    };

    updateSpreadsheet = async(spreadSheetName) => 
    {
        await this.loadSpreadsheet(spreadSheetName);
        const spreadsheets = this.gsheets.getSpreadsheets();
        for (let i = 0; i < spreadsheets.size(); i++) {
            if(this.currentSpreadsheetId === spreadsheets.get(i).getId())
            {
                const currentSpreadsheet = spreadsheets.get(i);
                const originalName = currentSpreadsheet.getName();
                const originalLocale=currentSpreadsheet.getLocale();
                const originalTimeZone=currentSpreadsheet.getTimeZone();

                const newName = await this.askQuestion(`Name [${originalName}]: `)
                const newLocale = await this.askQuestion(`Locale [${originalLocale}]: `)
                const newTimeZone = await this.askQuestion(`TimeZone [${originalTimeZone}]: `)

                this.gsheets.getSpreadsheets().get(i).setName(newName || originalName)
                this.gsheets.getSpreadsheets().get(i).setLocale(newLocale || originalLocale)
                this.gsheets.getSpreadsheets().get(i).setTimeZone(newTimeZone || originalTimeZone)

                await this.gsheets.updateSpreadsheet();
                console.log("Spreadsheet updated successfully!")
            } 
        }
        this.reset();
    };

    updateSheet = async() => 
    {
        if(this.selectedSpreadsheetName === "")
        {
            console.log("A spreadsheet must first be selected")
            return;
        }

        this.currentSheetId = await this.getNumberInput("Select a SheetID to update: ")
        await this.gsheets.loadSheet(this.currentSheetId);
        if(await this.getConfirmation("Change sheet name (Y/N)? "))
        {
            const newName = await this.askQuestion("Enter new sheet name: ")
            this.gsheets.getSheets().get(0).setTitle(newName);
            await this.gsheets.updateSheet();
        }
    };

    deleteSheet = async() => 
    {
        if(this.selectedSpreadsheetName === "")
        {
            console.log("A spreadsheet must first be selected")
            return;
        }
        this.currentSheetId = await this.getNumberInput("Select a SheetID to delete: ")
        await this.gsheets.loadSheet(this.currentSheetId);
        if(await this.getConfirmation("Delete Sheet (Y/N)? "))
        {
            await this.gsheets.deleteSheet(this.currentSheetId);
        }
    }
    deleteSpreadsheet = async(spreadSheetName) => 
    {   
        await this.loadSpreadsheet(spreadSheetName);
        await this.gsheets.deleteSpreadsheet(this.currentSpreadsheetId)
        await this.reset()
    }
    export = async() => 
    {
        await this.loadSpreadsheet();
        console.log("Supported file extensions: csv, ods, pdf, tsv, xlsx, zip");
        const fileSavePath = await this.askQuestion("Enter export location ex. (C:/Downloads/myFile.pdf) ")
        await this.gsheets.export(this.currentSpreadsheetId,fileSavePath)
        await this.reset()
    }
    exit = async()=>
    {
        this.currentSpreadsheetId ="";
        this.currentSheetId="";
        await this.gsheets.listSpreadsheets();
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
            this.gsheets = new CloudSheets.googlesheets();

            await this.authenticate(myArgs.get("c"), myArgs.get("s"));
            console.log("Loading Spreadsheets...");
            await this.gsheets.listSpreadsheets();
            this.currentSpreadsheetId ? this.displaySpreadsheetMenu() : this.displayMainMenu();

            while (this.running) {
                try {

                    const input= await this.getCommand();
                    const [command, ...args] = input.trim().split(/\s+/);

                    // Process commands
                    switch (command.toLowerCase().trim()) {
                        case 'ls':
                            await this.listSpreadsheets();
                            break;

                        case '?':
                            this.currentSpreadsheetId ? this.displaySpreadsheetMenu() : this.displayMainMenu();
                            break;
                        case 'sel':
                            await this.selectSpreadsheet(args[0]);
                            break;
                        case 'lssh':
                            await this.listSheets();
                            break;
                        case 'cs':
                            await this.createSpreadsheet(args[0]);
                            break;
                        case 'add': 
                            await this.addSheet(args[0]);
                            break;
                        case 'us': 
                            await this.updateSpreadsheet(args[0]);
                            break;
                        case 'update': 
                            await this.updateSheet();
                            break;
                        case 'del': 
                            await this.deleteSpreadsheet(args[0]);
                            break;
                        case 'delete': 
                            await this.deleteSheet();
                            break;
                        case 'exp': 
                            await this.export();
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
            const oauth = new CloudSheets.oauth();
            oauth.setClientId(client)
            oauth.setClientSecret(secret);
            oauth.setAuthorizationScope(this.oAuthScope);
            oauth.setServerAuthURL(this.oAuthAuthUrl);
            oauth.setServerTokenURL(this.oAuthTokenUrl);

            try {
                let authorization = await oauth.getAuthorization()
                await this.gsheets.setAuthorization(authorization);
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
const demo = new GoogleSheetsDemo();
demo.main(process.argv.slice(2));

function prompt(promptName, label, punctuation, defaultVal)
{
  lastPrompt = promptName;
  lastDefault = defaultVal;
  process.stdout.write(`${label} [${defaultVal}]${punctuation} `);
}
