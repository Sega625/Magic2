object MDBForm: TMDBForm
  Left = 501
  Top = 226
  BorderStyle = bsSingle
  Caption = 'Magic2'
  ClientHeight = 481
  ClientWidth = 687
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  TextHeight = 13
  object LoadMDBLab: TLabel
    Left = 122
    Top = 43
    Width = 22
    Height = 71
    Caption = #8226
    Font.Charset = DEFAULT_CHARSET
    Font.Color = 13302029
    Font.Height = -53
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
  end
  object LoadNormsLab: TLabel
    Left = 122
    Top = 93
    Width = 22
    Height = 71
    Caption = #8226
    Font.Charset = DEFAULT_CHARSET
    Font.Color = 13302029
    Font.Height = -53
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
  end
  object LoadMapLab: TLabel
    Left = 122
    Top = 143
    Width = 22
    Height = 71
    Caption = #8226
    Font.Charset = DEFAULT_CHARSET
    Font.Color = 13302029
    Font.Height = -53
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
  end
  object Label1: TLabel
    Left = 245
    Top = 13
    Width = 57
    Height = 16
    Caption = #1055#1083#1072#1089#1090#1080#1085#1099
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 465
    Top = 13
    Width = 68
    Height = 16
    Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090#1099
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label3: TLabel
    Left = 21
    Top = 438
    Width = 106
    Height = 16
    Caption = #1042#1088#1077#1084#1103' '#1086#1073#1088#1072#1073#1086#1090#1082#1080':'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object TimeLab: TLabel
    Left = 17
    Top = 456
    Width = 110
    Height = 16
    Alignment = taCenter
    AutoSize = False
    Caption = '0.0 '#1089#1077#1082'.'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label4: TLabel
    Left = 10
    Top = 9
    Width = 142
    Height = 16
    Caption = #1048#1079#1084#1077#1088#1080#1090#1077#1083#1100#1085#1072#1103' '#1089#1080#1089#1090#1077#1084#1072
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object OpenDirLab: TLabel
    Left = 122
    Top = 193
    Width = 22
    Height = 71
    Caption = #8226
    Font.Charset = DEFAULT_CHARSET
    Font.Color = 13302029
    Font.Height = -53
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
    Visible = False
  end
  object LoadMDBBtn: TButton
    Left = 8
    Top = 70
    Width = 111
    Height = 25
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' MDB'
    TabOrder = 0
    OnClick = LoadMDBBtnClick
  end
  object WafersLB: TListBox
    Left = 161
    Top = 31
    Width = 217
    Height = 410
    BevelKind = bkFlat
    BorderStyle = bsNone
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Roboto'
    Font.Style = []
    ItemHeight = 15
    MultiSelect = True
    ParentFont = False
    Sorted = True
    TabOrder = 1
    OnDrawItem = WafersLBDrawItem
  end
  object ProcessBtn: TButton
    Left = 161
    Top = 447
    Width = 219
    Height = 25
    Caption = #1054#1073#1088#1072#1073#1086#1090#1072#1090#1100
    Enabled = False
    TabOrder = 2
    OnClick = ProcessBtnClick
  end
  object LoadNormsBtn: TButton
    Left = 8
    Top = 120
    Width = 111
    Height = 25
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1085#1086#1088#1084#1099
    TabOrder = 3
    OnClick = LoadNormsBtnClick
  end
  object LoadMapBtn: TButton
    Left = 7
    Top = 170
    Width = 111
    Height = 25
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1082#1072#1088#1090#1091
    TabOrder = 4
    OnClick = LoadMapBtnClick
  end
  object ResultRE: TRichEdit
    Left = 385
    Top = 31
    Width = 296
    Height = 441
    BevelKind = bkFlat
    BorderStyle = bsNone
    Ctl3D = False
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Roboto'
    Font.Style = []
    ParentCtl3D = False
    ParentFont = False
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 5
  end
  object ClearBtn: TButton
    Left = 607
    Top = 6
    Width = 73
    Height = 21
    Caption = #1054#1095#1080#1089#1090#1080#1090#1100
    TabOrder = 6
    OnClick = ClearBtnClick
  end
  object MSystemCB: TComboBox
    Left = 8
    Top = 31
    Width = 144
    Height = 21
    ItemIndex = 0
    TabOrder = 7
    Text = #1043#1072#1084#1084#1072'-156'
    OnChange = MSystemCBChange
    Items.Strings = (
      #1043#1072#1084#1084#1072'-156'
      'Schuster TSM 664')
  end
  object OpenDirBtn: TButton
    Left = 7
    Top = 220
    Width = 111
    Height = 25
    Caption = #1054#1090#1082#1088#1099#1090#1100' '#1087#1072#1087#1082#1091
    TabOrder = 8
    OnClick = OpenDirBtnClick
  end
  object MainMenu1: TMainMenu
    Left = 329
    Top = 39
    object PrefMenu: TMenuItem
      Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1080
      OnClick = PrefMenuClick
    end
    object ExitMenu: TMenuItem
      Caption = #1042#1099#1093#1086#1076
      OnClick = ExitMenuClick
    end
  end
end
