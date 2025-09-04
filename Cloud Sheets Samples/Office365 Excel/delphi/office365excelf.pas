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
unit office365excelf;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls,
  Vcl.NumberBox, chcore, chtypes, choauth, choffice365excel;

type
  TFormOffice365Excel = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label7: TLabel;
    clientId: TEdit;
    clientSecret: TEdit;
    AuthBtn: TButton;
    GroupBox2: TGroupBox;
    sessionBtn: TButton;
    listWorkbooksBtn: TButton;
    createWorkbookBtn: TButton;
    deleteWorkbookBtn: TButton;
    lvWorkbooks: TListView;
    GroupBox3: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label11: TLabel;
    worksheetNameEdit: TEdit;
    updateWorksheetBtn: TButton;
    visibilityComboBox: TComboBox;
    positionNumBox: TNumberBox;
    chOAuth1: TchOAuth;
    chOffice365Excel1: TchOffice365Excel;
    procedure AuthBtnClick(Sender: TObject);
    procedure createWorkbookBtnClick(Sender: TObject);
    procedure deleteWorkbookBtnClick(Sender: TObject);
    procedure listWorkbooksBtnClick(Sender: TObject);
    procedure OnWorkbookList(Sender: TObject; const Id: String; const Name: String);
    procedure updateWorksheetBtnClick(Sender: TObject);
    procedure lvWorkbooksDblClick(Sender: TObject);
    procedure lvWorkbooksSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure sessionBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    listedWorksheets, listedWorkbooks: Boolean;
    SelectedWorkbookId: String;
    SelectedWorksheetId: String;
  end;

var
  FormOffice365Excel: TFormOffice365Excel;

implementation

{$R *.dfm}

procedure TFormOffice365Excel.AuthBtnClick(Sender: TObject);
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
    chOAuth1.AuthorizationScope := 'offline_access files.readwrite user.read';
    chOAuth1.ServerAuthURL := 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
    chOAuth1.ServerTokenURL := 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
    chOAuth1.GrantType := TchOAuthGrantTypes.ogtAuthorizationCode;
    chOAuth1.ClientProfile := TchOAuthClientProfiles.ocpApplication;

    chOffice365Excel1.Authorization := chOAuth1.GetAuthorization();

    ListWorkbooksBtnClick(Sender);

  except
    on E: ECloudSheets do
    begin
      ShowMessage('An error occurred: ' + e.Message)
    end;

  end;
end;

procedure TFormOffice365Excel.createWorkbookBtnClick(Sender: TObject);
var
  WorkbookName: string;
  InputCancelled: Boolean;
begin
  InputCancelled := InputQuery('Create Workbook', 'Enter the name of the new workbook:', WorkbookName);

  if not InputCancelled then
    Exit;

  if WorkbookName.Trim = '' then
  begin
    ShowMessage('Workbook name cannot be empty.');
    Exit;
  end;

  try
    chOffice365Excel1.CreateWorkbook(WorkbookName);

    listWorkbooksBtnClick(Sender);
  except
    on E: ECloudSheets do
    begin
      ShowMessage('Failed to create Workbook: ' + E.Message);
    end;
  end;
end;

procedure TFormOffice365Excel.deleteWorkbookBtnClick(Sender: TObject);
var
  SelectedID: string;
begin

  if listedWorksheets then
  begin
     listWorkbooksBtnClick(Sender);
     Exit;
  end;

  if lvWorkbooks.Selected = nil then
  begin
    ShowMessage('Please select a Workbook to delete.');
    Exit;
  end;

  SelectedID := lvWorkbooks.Selected.Caption;

  try
    chOffice365Excel1.DeleteWorkbook(SelectedID);

    ListWorkbooksBtnClick(Sender);
  except
    on E: ECloudSheets do
    begin
      ShowMessage('Failed to delete workbook: ' + E.Message);
    end;
  end;
end;

procedure TFormOffice365Excel.FormCreate(Sender: TObject);
begin

  Self.BorderStyle := bsSingle;
  Self.BorderIcons := [biSystemMenu, biMinimize];
  lvWorkbooks.ReadOnly := True;

  lvWorkbooks.Columns.Add.Caption := 'ID';
  lvWorkbooks.Columns.Add.Caption := 'Name';
  lvWorkbooks.Columns[0].AutoSize := True;
  lvWorkbooks.Columns[1].AutoSize := True;

  chOffice365Excel1.OnWorkbookList := OnWorkbookList;

  lvWorkbooks.ViewStyle := vsReport;
end;

procedure TFormOffice365Excel.OnWorkbookList(Sender: TObject; const Id: string; const Name: string);
var
  ListItem: TListItem;

begin
  ListItem := lvWorkbooks.Items.Add;
  ListItem.Caption := Id;
  ListItem.SubItems.Add(Name);
end;


procedure TFormOffice365Excel.listWorkbooksBtnClick(Sender: TObject);
begin
  if listedWorksheets then
  begin
    lvWorkbooks.Clear;
    lvWorkbooks.Columns.Clear;
    lvWorkbooks.Columns.Add.Caption := 'ID';
    lvWorkbooks.Columns.Add.Caption := 'Name';
    lvWorkbooks.Columns[0].Width := 300;
    lvWorkbooks.Columns[1].Width := 300;
    lvWorkbooks.ViewStyle := vsReport;
    listedWorksheets := False;
  end;

  lvWorkbooks.Clear();
  chOffice365Excel1.ListWorkbooks();
  listedWorkbooks := True;
end;

procedure TFormOffice365Excel.lvWorkbooksDblClick(Sender: TObject);
var
  ListItem: TListItem;
  i: Integer;
  worksheet: TchOEWorksheet;
begin
    if listedWorksheets then
    Exit;

    SelectedWorkbookId := lvWorkbooks.Selected.Caption;

    chOffice365Excel1.LoadWorkbook(SelectedWorkbookId);

    lvWorkbooks.Clear;
    lvWorkbooks.Columns.Clear;
    lvWorkbooks.Columns.Add.Caption := 'WorksheetId';
    lvWorkbooks.Columns.Add.Caption := 'Name';
    lvWorkbooks.Columns.Add.Caption := 'Position';
    lvWorkbooks.Columns[0].Width := 200;
    lvWorkbooks.Columns[1].Width := 200;
    lvWorkbooks.Columns[2].Width := 200;
    lvWorkbooks.ViewStyle := vsReport;

    for i := 0 to chOffice365Excel1.Worksheets.Count - 1 do
    begin
      worksheet := chOffice365Excel1.Worksheets[i];

      ListItem := lvWorkbooks.Items.Add;
      ListItem.Caption := worksheet.WorksheetId;
      ListItem.SubItems.Add(worksheet.Name);
      ListItem.SubItems.Add(IntToStr(worksheet.Position));
    end;

    listedWorkbooks := False;
    listedWorksheets := True;
end;

procedure TFormOffice365Excel.lvWorkbooksSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
    workbook: TchOEWorkbook;
    worksheet: TchOEWorksheet;
begin
    if listedworksheets then
    begin
      worksheetNameEdit.Clear;
      positionNumBox.Value := 0;
      visibilityComboBox.ItemIndex := 0;

      if lvWorkbooks.Selected = nil then
        Exit;

      worksheet := nil;

      SelectedWorksheetId := lvWorkbooks.Selected.Caption;

      chOffice365Excel1.LoadWorksheet(SelectedWorksheetId);

      worksheet := chOffice365Excel1.Worksheets[0];

      worksheetNameEdit.Text := worksheet.Name;
      positionNumBox.Value := Trunc(worksheet.Position);
      visibilityComboBox.ItemIndex := Integer(worksheet.Visibility);
    end;
end;

procedure TFormOffice365Excel.sessionBtnClick(Sender: TObject);
var
  UsePersistentSession: Integer;
begin
  if (lvWorkbooks.Selected = nil) or (not listedWorkbooks) then
  begin
    ShowMessage('A workbook must be selected before creating a session.');
    Exit;
  end;

  SelectedWorkbookId := lvWorkbooks.Selected.Caption;

  try
    UsePersistentSession := MessageDlg('Use persistent session?',
                                       mtConfirmation, [mbYes, mbNo, mbCancel], 0);

      Exit;

    chOffice365Excel1.LoadWorkbook(SelectedWorkbookId);

    if UsePersistentSession = mrYes then
      chOffice365Excel1.CreateSession(True)
    else
      chOffice365Excel1.CreateSession(False);
  except
    on E: ECloudSheets do
    begin
      ShowMessage('Failed to create session: ' + E.Message);
    end;
  end;
end;

procedure TFormOffice365Excel.updateWorksheetBtnClick(Sender: TObject);
var
  worksheet: TchOEWorksheet;
begin
  if worksheetNameEdit.Text = '' then
    begin
      ShowMessage('Worksheet title cannot be empty');
      Exit;
    end
    else
    begin
       if listedWorksheets then
       begin
         if lvWorkbooks.Selected = nil then
            Exit;

            worksheet := nil;
            SelectedWorksheetId := lvWorkbooks.Selected.Caption;

            chOffice365Excel1.LoadWorkSheet(SelectedWorksheetId);

            worksheet := chOffice365Excel1.Worksheets[0];

            try

            worksheet.Name := workSheetNameEdit.Text;
            worksheet.Position := Trunc(positionNumBox.Value);
            worksheet.Visibility := TchTOEWorksheetVisibilities(visibilityComboBox.ItemIndex);

            chOffice365Excel1.UpdateWorksheet();

            ListWorkbooksBtnClick(Sender);

            except
              on E: ECloudspreadsheets do
              begin
                ShowMessage('An error occurred: ' + e.Message);
              end;

            end;

       end;
    end;
end;

end.

