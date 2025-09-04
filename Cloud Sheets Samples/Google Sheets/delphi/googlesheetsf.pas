(*
 * Cloud Sheets 2024 Delphi Edition - Sample Project
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
 *)
unit googlesheetsf;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, chgooglesheets, chcore,
  chtypes, choauth, Vcl.ComCtrls, Vcl.NumberBox;

type
  TFormGoogleSheets = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    clientId: TEdit;
    clientSecret: TEdit;
    AuthBtn: TButton;
    chOAuth1: TchOAuth;
    chGoogleSheets1: TchGoogleSheets;
    GroupBox2: TGroupBox;
    lvSpreadsheets: TListView;
    ListSpreadsheetsBtn: TButton;
    CreateSpreadShBtn: TButton;
    DeleteSpreadshBtn: TButton;
    GroupBox3: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label11: TLabel;
    spreadsheetNameEdit: TEdit;
    spreadsheetLocaleEdit: TEdit;
    spreadsheetTimezoneEdit: TEdit;
    updateSpreadsheetBtn: TButton;
    GroupBox4: TGroupBox;
    updateSheetBtn: TButton;
    Label5: TLabel;
    Label6: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    hideGridLineschkbox: TCheckBox;
    hiddenchkbox: TCheckBox;
    sheetIdEdit: TEdit;
    sheetTitleEdit: TEdit;
    rowCountNumBox: TNumberBox;
    colCountNumBox: TNumberBox;
    Label7: TLabel;
    exportBtn: TButton;
    procedure AuthBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ListSpreadsheetsBtnClick(Sender: TObject);
    procedure OnSpreadsheetList(Sender: TObject; const Id: String; const Name: String);
    procedure CreateSpreadShBtnClick(Sender: TObject);
    procedure DeleteSpreadshBtnClick(Sender: TObject);
    procedure lvSpreadsheetsSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure updateSpreadsheetBtnClick(Sender: TObject);
    procedure lvSpreadsheetsDblClick(Sender: TObject);
    procedure updateSheetBtnClick(Sender: TObject);
    procedure exportBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    listedSheets, listedSpreadsheets: Boolean;
    SelectedSpreadsheetId: String;
    SelectedSheetId: Integer;
  end;

var
  FormGoogleSheets: TFormGoogleSheets;

implementation

{$R *.dfm}


procedure TFormGoogleSheets.OnSpreadsheetList(Sender: TObject; const Id: string; const Name: string);
var
  ListItem: TListItem;

begin
  ListItem := lvSpreadsheets.Items.Add;
  ListItem.Caption := Id;
  ListItem.SubItems.Add(Name);
end;

procedure TFormGoogleSheets.updateSheetBtnClick(Sender: TObject);
var
  sheet: TchGSSheet;
  sheetIdStr: string;
begin
  if sheetTitleEdit.Text = '' then
    begin
      ShowMessage('Sheet title cannot be empty');
      Exit;
    end
    else
    begin
       if listedSheets then
       begin
         if lvSpreadsheets.Selected = nil then
            Exit;

            sheet := nil;
            sheetIdStr := lvSpreadsheets.Selected.Caption;

            TryStrToInt(sheetIdStr, SelectedSheetId);

            chGoogleSheets1.LoadSheet(SelectedSheetId);

            sheet := chGoogleSheets1.Sheets[0];

            try

            sheet.Title := sheetTitleEdit.Text;
            sheet.TotalRowCount := Trunc(rowCountNumBox.Value);
            sheet.TotalColumnCount := Trunc(colCountNumBox.Value);
            sheet.Hidden := hiddenchkbox.Checked;
            sheet.HideGridlines := hidegridlineschkbox.Checked;

            chGoogleSheets1.UpdateSheet();

            ListSpreadsheetsBtnClick(Sender);

            except
              on E: ECloudspreadsheets do
              begin
                ShowMessage('An error occurred: ' + e.Message);
              end;

            end;

       end;
    end;
end;

procedure TFormGoogleSheets.updateSpreadsheetBtnClick(Sender: TObject);
var
  spreadsheet: TchGSSpreadsheet;
begin
    if spreadsheetNameEdit.Text = '' then
    begin
      ShowMessage('Spreadsheet name cannot be empty');
      Exit;
    end
    else
    begin
       if listedSpreadsheets then
       begin
         if lvSpreadsheets.Selected = nil then
            Exit;

            spreadsheet := nil;
            SelectedSpreadsheetId := lvSpreadsheets.Selected.Caption;

            chGoogleSheets1.LoadSpreadsheet(SelectedSpreadsheetId);

            spreadsheet := chGoogleSheets1.Spreadsheets[0];

            try

            spreadsheet.Name := spreadsheetNameEdit.Text;
            spreadsheet.TimeZone := spreadsheetTimezoneEdit.Text;
            spreadsheet.Locale := spreadsheetLocaleEdit.Text;

            chGoogleSheets1.UpdateSpreadsheet();

            ListSpreadsheetsBtnClick(Sender);

            except
              on E: ECloudspreadsheets do
              begin
                ShowMessage('An error occurred: ' + e.Message);
              end;

            end;

       end;
    end;


end;

procedure TFormGoogleSheets.AuthBtnClick(Sender: TObject);
begin
  if (clientId.Text = '') or (clientSecret.Text = '') then
  begin
    ShowMessage('Client ID and Client Secret cannot be empty!');
    Exit;
  end
  else
  try
    chOAuth1.ClientId := clientId.Text;
    chOAuth1.ClientSecret := clientSecret.Text;
    chOAuth1.AuthorizationScope := 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive';
    chOAuth1.ServerAuthURL := 'https://accounts.google.com/o/oauth2/auth';
    chOAuth1.ServerTokenURL := 'https://accounts.google.com/o/oauth2/token';
    chOAuth1.GrantType := TchOAuthGrantTypes.ogtAuthorizationCode;
    chOAuth1.ClientProfile := TchOAuthClientProfiles.ocpApplication;

    chGoogleSheets1.Authorization := chOAuth1.GetAuthorization();

    ListSpreadsheetsBtnClick(Sender);

  except
    on E: ECloudSheets do
    begin
      ShowMessage('An error occurred: ' + e.Message)
    end;

  end;
end;

procedure TFormGoogleSheets.ListSpreadsheetsBtnClick(Sender: TObject);
begin
  if listedSheets then
  begin
    lvSpreadsheets.Clear;
    lvSpreadsheets.Columns.Clear;
    lvSpreadsheets.Columns.Add.Caption := 'ID';
    lvSpreadsheets.Columns.Add.Caption := 'Name';
    lvSpreadsheets.Columns[0].Width := 300;
    lvSpreadsheets.Columns[1].Width := 300;
    lvSpreadsheets.ViewStyle := vsReport;
    listedSheets := False;
  end;

  lvSpreadsheets.Clear();
  chGoogleSheets1.ListSpreadsheets();
  listedSpreadsheets := True;
end;

procedure TFormGoogleSheets.lvSpreadsheetsDblClick(Sender: TObject);
var
  ListItem: TListItem;
  i: Integer;
  sheet: TchGSSheet;
begin
    if listedSheets then
    Exit;

    SelectedSpreadsheetId := lvSpreadsheets.Selected.Caption;

    chGoogleSheets1.LoadSpreadsheet(SelectedSpreadsheetId);

    lvSpreadsheets.Clear;
    lvSpreadsheets.Columns.Clear;
    lvSpreadsheets.Columns.Add.Caption := 'SheetId';
    lvSpreadsheets.Columns.Add.Caption := 'Title';
    lvSpreadsheets.Columns.Add.Caption := 'Row count';
    lvSpreadsheets.Columns.Add.Caption := 'Column count';
    lvSpreadsheets.Columns[0].Width := 150;
    lvSpreadsheets.Columns[1].Width := 150;
    lvSpreadsheets.Columns[2].Width := 150;
    lvSpreadsheets.Columns[3].Width := 150;
    lvSpreadsheets.ViewStyle := vsReport;

    for i := 0 to chGoogleSheets1.Sheets.Count - 1 do
    begin
      sheet := chGoogleSheets1.Sheets[i];

      ListItem := lvSpreadsheets.Items.Add;
      ListItem.Caption := IntToStr(sheet.sheetId);
      ListItem.SubItems.Add(sheet.Title);
      ListItem.SubItems.Add(IntToStr(sheet.TotalRowCount));
      ListItem.SubItems.Add(IntToStr(sheet.TotalColumnCount));
    end;

    listedSpreadsheets := False;
    listedSheets := True;

end;

procedure TFormGoogleSheets.lvSpreadsheetsSelectItem(Sender: TObject;
  Item: TListItem; Selected: Boolean);
var
    spreadsheet: TchGSSpreadsheet;
    sheet: TchGSSheet;
    sheetIdStr: String;
begin
    if listedSheets then
    begin
      sheetIdEdit.Clear;
      sheetTitleEdit.Clear;
      rowCountNumBox.Clear;
      colCountNumBox.Clear;
      hideGridLineschkbox.Checked := False;
      hiddenchkbox.Checked := False;

      if lvSpreadsheets.Selected = nil then
        Exit;

      sheet := nil;

      sheetIdStr := lvSpreadsheets.Selected.Caption;
      TryStrToInt(sheetIdStr, SelectedSheetId);

      chGoogleSheets1.LoadSheet(SelectedSheetId);

      sheet := chGoogleSheets1.Sheets[0];

      sheetIdEdit.Text := sheet.SheetId.ToString();
      sheetTitleEdit.Text := sheet.Title;
      rowCountNumBox.Value := sheet.TotalRowCount;
      colCountNumBox.Value := sheet.TotalColumnCount;
      hideGridLineschkbox.Checked := sheet.HideGridlines;
      hiddenchkbox.Checked := sheet.Hidden;
    end;


    if listedSpreadsheets then
    begin
      spreadsheetNameEdit.Clear;
      spreadsheetTimezoneEdit.Clear;
      spreadsheetLocaleEdit.Clear;

      if lvSpreadsheets.Selected = nil then
        Exit;

      SelectedSpreadsheetId := lvSpreadsheets.Selected.Caption;
      spreadsheet := nil;

      chGoogleSheets1.LoadSpreadsheet(SelectedSpreadsheetId);

      spreadsheet := chGoogleSheets1.Spreadsheets[0];

      spreadsheetNameEdit.Text := spreadsheet.Name;
      spreadsheetTimezoneEdit.Text := spreadsheet.Timezone;
      spreadsheetLocaleEdit.Text := spreadsheet.Locale;
    end;
end;

procedure TFormGoogleSheets.CreateSpreadShBtnClick(Sender: TObject);
var
  SpreadsheetName: string;
begin
  SpreadsheetName := InputBox('Create Spreadsheet', 'Enter the name of the new spreadsheet:', '');

  if SpreadsheetName.Trim = '' then
  begin
    ShowMessage('Spreadsheet name cannot be empty.');
    Exit;
  end;

  try
    chGoogleSheets1.CreateSpreadsheet(SpreadsheetName, '', '');

    ListSpreadsheetsBtnClick(Sender);
  except
    on E: ECloudSheets do
    begin
      ShowMessage('Failed to create Spreadsheet: ' + E.Message);
    end;
  end;
end;

procedure TFormGoogleSheets.DeleteSpreadshBtnClick(Sender: TObject);
var
  SelectedID: string;
begin

  if listedSheets then
  begin
     ListSpreadsheetsBtnClick(Sender);
     Exit;
  end;

  if lvSpreadsheets.Selected = nil then
  begin
    ShowMessage('Please select a Spreadsheet to delete.');
    Exit;
  end;

  SelectedID := lvSpreadsheets.Selected.Caption;

  try
    chGoogleSheets1.DeleteSpreadsheet(SelectedID);

    ListSpreadsheetsBtnClick(Sender);
  except
    on E: ECloudSheets do
    begin
      ShowMessage('Failed to delete spreadsheet: ' + E.Message);
    end;
  end;
end;

procedure TFormGoogleSheets.exportBtnClick(Sender: TObject);
var
  SaveDialog: TSaveDialog;
  FullPath: string;
begin

  if lvSpreadsheets.Selected = nil then
  Exit;

  SelectedSpreadsheetId := lvSpreadsheets.Selected.Caption;

  SaveDialog := TSaveDialog.Create(nil);
  try
    SaveDialog.Filter :=
      'PDF Files (*.pdf)|*.pdf|' +
      'OpenDocument Spreadsheet (*.ods)|*.ods|' +
      'CSV Files (*.csv)|*.csv|' +
      'Tab-Separated Values (*.tsv)|*.tsv|' +
      'Excel Files (*.xlsx)|*.xlsx|' +
      'ZIP Files (*.zip)|*.zip|' +
      'All Files (*.*)|*.*';
    SaveDialog.DefaultExt := 'pdf';

    if SaveDialog.Execute then
    begin
      try
        FullPath := SaveDialog.FileName;
        chGoogleSheets1.Export(SelectedSpreadsheetId, FullPath);
        ShowMessage('File saved successfully!');
      except
        on E: ECloudSheets do
          ShowMessage('Error Message: ' + E.Message);
      end;
    end;
  finally
    SaveDialog.Free;
  end;
end;

procedure TFormGoogleSheets.FormCreate(Sender: TObject);
begin

  Self.BorderStyle := bsSingle;
  Self.BorderIcons := [biSystemMenu, biMinimize];
  lvSpreadsheets.ReadOnly := True;

  lvSpreadsheets.Columns.Add.Caption := 'ID';
  lvSpreadsheets.Columns.Add.Caption := 'Name';
  lvSpreadsheets.Columns[0].AutoSize := True;
  lvSpreadsheets.Columns[1].AutoSize := True;

  chGoogleSheets1.OnSpreadsheetList := OnSpreadsheetList;

  lvSpreadsheets.ViewStyle := vsReport;
end;

end.

