object PrefForm: TPrefForm
  Left = 694
  Top = 353
  BorderStyle = bsToolWindow
  Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1080
  ClientHeight = 199
  ClientWidth = 295
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Position = poDesigned
  TextHeight = 13
  object MainGroup: TGroupBox
    Left = 10
    Top = 5
    Width = 274
    Height = 151
    BiDiMode = bdLeftToRight
    Caption = ' '#1054#1089#1085#1086#1074#1085#1099#1077' '
    Color = clBtnFace
    Ctl3D = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = 'Segoe UI Semilight'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentBiDiMode = False
    ParentColor = False
    ParentCtl3D = False
    ParentFont = False
    TabOrder = 0
    object CreateSTSChB: TCheckBox
      Tag = 1
      Left = 13
      Top = 63
      Width = 150
      Height = 17
      Caption = #1057#1086#1079#1076#1072#1090#1100' STS '#1092#1072#1081#1083#1099
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'Segoe UI Semibold'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
    object NoNormsChB: TCheckBox
      Tag = 2
      Left = 13
      Top = 94
      Width = 257
      Height = 17
      Caption = #1056#1072#1079#1088#1077#1096#1080#1090#1100' '#1088#1072#1073#1086#1090#1091' '#1073#1077#1079' '#1085#1086#1088#1084
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'Segoe UI Semilight'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
    end
    object ToFirstFailChB: TCheckBox
      Left = 14
      Top = 30
      Width = 150
      Height = 17
      Caption = #1044#1086' 1-'#1075#1086' '#1073#1088#1072#1082#1072
      Checked = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'Segoe UI Semibold'
      Font.Style = [fsBold]
      ParentFont = False
      State = cbChecked
      TabOrder = 2
    end
    object MapByParamsChB: TCheckBox
      Tag = 3
      Left = 13
      Top = 125
      Width = 255
      Height = 17
      Caption = #1050#1072#1088#1090#1072' '#1075#1086#1076#1085#1086#1089#1090#1080' '#1087#1086' '#1087#1072#1088#1072#1084#1077#1090#1088#1072#1084
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'Segoe UI Semibold'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 3
    end
  end
  object CloseBtn: TBitBtn
    Left = 9
    Top = 163
    Width = 275
    Height = 27
    Caption = #1047#1072#1082#1088#1099#1090#1100
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Segoe UI Semibold'
    Font.Style = []
    Kind = bkOK
    NumGlyphs = 2
    ParentFont = False
    TabOrder = 1
  end
end
