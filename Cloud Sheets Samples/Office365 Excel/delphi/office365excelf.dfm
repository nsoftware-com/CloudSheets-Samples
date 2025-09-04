object FormOffice365Excel: TFormOffice365Excel
  Left = 0
  Top = 0
  Caption = 'FormOffice365Excel'
  ClientHeight = 553
  ClientWidth = 808
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
      'The Office365Excel demo shows how to use the Office365Excel comp' +
      'onent to manage workbooks in OneDrive.'#13#10'You can List, Create, De' +
      'lete and manage both workbooks and worksheets.'#13#10'To begin click A' +
      'uthorize to allow the application to access your account. The Au' +
      'thorization String is created during this process.'
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
    object clientId: TEdit
      Left = 88
      Top = 26
      Width = 201
      Height = 23
      TabOrder = 0
    end
    object clientSecret: TEdit
      Left = 418
      Top = 26
      Width = 201
      Height = 23
      TabOrder = 1
    end
    object AuthBtn: TButton
      Left = 650
      Top = 25
      Width = 115
      Height = 24
      Caption = 'Authenticate'
      TabOrder = 2
      OnClick = AuthBtnClick
    end
  end
  object GroupBox2: TGroupBox
    Left = 16
    Top = 151
    Width = 777
    Height = 266
    Caption = 'Workbooks and worksheets'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    object sessionBtn: TButton
      Left = 650
      Top = 216
      Width = 115
      Height = 25
      Caption = 'Create session'
      TabOrder = 0
      OnClick = sessionBtnClick
    end
    object listWorkbooksBtn: TButton
      Left = 650
      Top = 17
      Width = 115
      Height = 25
      Caption = 'List Workbooks'
      TabOrder = 1
      OnClick = listWorkbooksBtnClick
    end
    object createWorkbookBtn: TButton
      Left = 650
      Top = 64
      Width = 115
      Height = 25
      Caption = 'Create Workbook'
      TabOrder = 2
      OnClick = createWorkbookBtnClick
    end
    object deleteWorkbookBtn: TButton
      Left = 650
      Top = 111
      Width = 115
      Height = 25
      Caption = 'Delete Workbook'
      TabOrder = 3
      OnClick = deleteWorkbookBtnClick
    end
    object lvWorkbooks: TListView
      Left = 16
      Top = 17
      Width = 600
      Height = 233
      Columns = <>
      Constraints.MaxHeight = 233
      Constraints.MaxWidth = 600
      Constraints.MinHeight = 233
      Constraints.MinWidth = 600
      TabOrder = 4
      OnDblClick = lvWorkbooksDblClick
      OnSelectItem = lvWorkbooksSelectItem
    end
  end
  object GroupBox3: TGroupBox
    Left = 16
    Top = 431
    Width = 777
    Height = 106
    Caption = 'Update worksheet'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    object Label3: TLabel
      Left = 29
      Top = 53
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
      Left = 267
      Top = 53
      Width = 46
      Height = 15
      Caption = 'Position:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = []
      ParentFont = False
    end
    object Label11: TLabel
      Left = 418
      Top = 53
      Width = 47
      Height = 15
      Caption = 'Visibility:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = []
      ParentFont = False
    end
    object worksheetNameEdit: TEdit
      Left = 88
      Top = 50
      Width = 153
      Height = 23
      TabOrder = 0
    end
    object updateWorksheetBtn: TButton
      Left = 650
      Top = 48
      Width = 115
      Height = 25
      Caption = 'Update'
      TabOrder = 1
      OnClick = updateWorksheetBtnClick
    end
    object visibilityComboBox: TComboBox
      Left = 483
      Top = 48
      Width = 121
      Height = 23
      TabOrder = 2
      Text = '0 - Visible'
      Items.Strings = (
        '0 - Visible'
        '1 - Hidden'
        '2 - Very hidden')
    end
    object positionNumBox: TNumberBox
      Left = 344
      Top = 50
      Width = 57
      Height = 23
      TabOrder = 3
    end
  end
  object chOAuth1: TchOAuth
    SSLAcceptServerCertStore = 'MY'
    SSLCertStore = 'MY'
    WebServerSSLCertStore = 'MY'
    Left = 648
    Top = 8
  end
  object chOffice365Excel1: TchOffice365Excel
    SSLAcceptServerCertStore = 'MY'
    SSLCertStore = 'MY'
    Left = 584
    Top = 15
  end
end


