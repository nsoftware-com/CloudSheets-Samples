object FormGoogleSheets: TFormGoogleSheets
  Left = 0
  Top = 0
  Caption = 'FormGoogleSheets'
  ClientHeight = 719
  ClientWidth = 813
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 15
  object Label7: TLabel
    Left = 16
    Top = 16
    Width = 670
    Height = 45
    Caption = 
      'The GoogleSheets demo shows how to use the GoogleSheets componen' +
      't to manage spreadsheets in Google Drive.'#13#10'Within the demo you c' +
      'an List, Delete, Create, Update and Export spreadsheets.'#13#10'To beg' +
      'in click Authorize to allow the application to access your accou' +
      'nt. The Authorization String is created during this process.'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clHighlight
    Font.Height = -12
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentColor = False
    ParentFont = False
  end
  object GroupBox1: TGroupBox
    Left = 16
    Top = 72
    Width = 777
    Height = 73
    Caption = 'Authentication'
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 32
      Width = 48
      Height = 15
      Caption = 'Client ID:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 312
      Top = 32
      Width = 69
      Height = 15
      Caption = 'Client Secret:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = []
      ParentFont = False
    end
  end
  object clientId: TEdit
    Left = 104
    Top = 101
    Width = 201
    Height = 23
    TabOrder = 1
  end
  object clientSecret: TEdit
    Left = 426
    Top = 101
    Width = 201
    Height = 23
    TabOrder = 2
  end
  object AuthBtn: TButton
    Left = 666
    Top = 101
    Width = 115
    Height = 24
    Caption = 'Authenticate'
    TabOrder = 3
    OnClick = AuthBtnClick
  end
  object GroupBox2: TGroupBox
    Left = 16
    Top = 151
    Width = 777
    Height = 266
    Caption = 'Spreadsheets and Sheets'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
    object exportBtn: TButton
      Left = 650
      Top = 216
      Width = 115
      Height = 25
      Caption = 'Export'
      TabOrder = 0
      OnClick = exportBtnClick
    end
  end
  object lvSpreadsheets: TListView
    Left = 33
    Top = 176
    Width = 600
    Height = 233
    Columns = <>
    Constraints.MaxHeight = 233
    Constraints.MaxWidth = 600
    Constraints.MinHeight = 233
    Constraints.MinWidth = 600
    TabOrder = 5
    OnDblClick = lvSpreadsheetsDblClick
    OnSelectItem = lvSpreadsheetsSelectItem
  end
  object ListSpreadsheetsBtn: TButton
    Left = 666
    Top = 176
    Width = 115
    Height = 25
    Caption = 'List Spreadsheets'
    TabOrder = 6
    OnClick = ListSpreadsheetsBtnClick
  end
  object CreateSpreadShBtn: TButton
    Left = 666
    Top = 216
    Width = 115
    Height = 25
    Caption = 'Create Spreadsheet'
    TabOrder = 7
    OnClick = CreateSpreadShBtnClick
  end
  object DeleteSpreadshBtn: TButton
    Left = 666
    Top = 255
    Width = 115
    Height = 25
    Caption = 'Delete Spreadsheet'
    TabOrder = 8
    OnClick = DeleteSpreadshBtnClick
  end
  object GroupBox3: TGroupBox
    Left = 16
    Top = 432
    Width = 777
    Height = 130
    Caption = 'Update Spreadsheet'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
    TabOrder = 9
    object Label3: TLabel
      Left = 17
      Top = 24
      Width = 35
      Height = 15
      Caption = 'Name:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBackground
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 18
      Top = 53
      Width = 54
      Height = 15
      Caption = 'Timezone:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = []
      ParentFont = False
    end
    object Label11: TLabel
      Left = 18
      Top = 88
      Width = 37
      Height = 15
      Caption = 'Locale:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = []
      ParentFont = False
    end
    object spreadsheetNameEdit: TEdit
      Left = 168
      Top = 24
      Width = 345
      Height = 23
      TabOrder = 0
    end
    object spreadsheetLocaleEdit: TEdit
      Left = 168
      Top = 82
      Width = 345
      Height = 23
      TabOrder = 1
    end
    object spreadsheetTimezoneEdit: TEdit
      Left = 168
      Top = 53
      Width = 345
      Height = 23
      TabOrder = 2
    end
  end
  object updateSpreadsheetBtn: TButton
    Left = 666
    Top = 481
    Width = 97
    Height = 25
    Caption = 'Update'
    TabOrder = 10
    OnClick = updateSpreadsheetBtnClick
  end
  object GroupBox4: TGroupBox
    Left = 16
    Top = 568
    Width = 777
    Height = 130
    Caption = 'Update Sheet'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
    TabOrder = 11
    object Label5: TLabel
      Left = 17
      Top = 35
      Width = 42
      Height = 15
      Caption = 'SheetId:'
    end
    object Label6: TLabel
      Left = 18
      Top = 91
      Width = 25
      Height = 15
      Caption = 'Title:'
    end
    object Label8: TLabel
      Left = 269
      Top = 35
      Width = 60
      Height = 15
      Caption = 'Row count:'
    end
    object Label9: TLabel
      Left = 269
      Top = 91
      Width = 80
      Height = 15
      Caption = 'Column count:'
    end
    object rowCountNumBox: TNumberBox
      Left = 376
      Top = 27
      Width = 121
      Height = 23
      TabOrder = 0
    end
    object hideGridLineschkbox: TCheckBox
      Left = 520
      Top = 33
      Width = 97
      Height = 17
      Caption = 'Hide Gridlines'
      TabOrder = 1
    end
    object hiddenchkbox: TCheckBox
      Left = 520
      Top = 89
      Width = 97
      Height = 17
      Caption = 'Hidden'
      TabOrder = 2
    end
    object colCountNumBox: TNumberBox
      Left = 376
      Top = 83
      Width = 121
      Height = 23
      TabOrder = 3
    end
  end
  object updateSheetBtn: TButton
    Left = 666
    Top = 626
    Width = 97
    Height = 25
    Caption = 'Update Sheet'
    TabOrder = 12
    OnClick = updateSheetBtnClick
  end
  object sheetIdEdit: TEdit
    Left = 104
    Top = 600
    Width = 137
    Height = 23
    Enabled = False
    TabOrder = 13
  end
  object sheetTitleEdit: TEdit
    Left = 104
    Top = 656
    Width = 137
    Height = 23
    TabOrder = 14
  end
  object chOAuth1: TchOAuth
    SSLAcceptServerCertStore = 'MY'
    SSLCertStore = 'MY'
    WebServerSSLCertStore = 'MY'
    Left = 480
    Top = 136
  end
  object chGoogleSheets1: TchGoogleSheets
    SSLAcceptServerCertStore = 'MY'
    SSLCertStore = 'MY'
    Left = 432
    Top = 136
  end
end


