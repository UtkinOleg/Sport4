inherited RazrDlg: TRazrDlg
  Caption = #1056#1072#1079#1088#1103#1076#1099
  ClientHeight = 335
  ClientWidth = 330
  OnShow = FormShow
  ExplicitWidth = 336
  ExplicitHeight = 367
  PixelsPerInch = 96
  TextHeight = 13
  inherited Bevel1: TBevel
    Width = 313
    Height = 289
    ExplicitWidth = 313
    ExplicitHeight = 289
  end
  inherited OKBtn: TButton
    Left = 86
    Top = 303
    OnClick = OKBtnClick
    ExplicitLeft = 86
    ExplicitTop = 303
  end
  inherited CancelBtn: TButton
    Left = 166
    Top = 303
    Caption = #1054#1090#1084#1077#1085#1072
    ExplicitLeft = 166
    ExplicitTop = 303
  end
  object HelpBtn: TButton
    Left = 246
    Top = 303
    Width = 75
    Height = 25
    Caption = '&'#1055#1086#1084#1086#1097#1100
    TabOrder = 2
    OnClick = HelpBtnClick
  end
  object ListBox1: TListBox
    Left = 16
    Top = 16
    Width = 217
    Height = 273
    ItemHeight = 13
    TabOrder = 3
    OnDblClick = ListBox1DblClick
  end
  object Button1: TButton
    Left = 239
    Top = 16
    Width = 75
    Height = 25
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100
    TabOrder = 4
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 239
    Top = 47
    Width = 75
    Height = 25
    Caption = #1048#1079#1084#1077#1085#1080#1090#1100
    TabOrder = 5
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 239
    Top = 78
    Width = 75
    Height = 25
    Caption = #1059#1076#1072#1083#1080#1090#1100
    TabOrder = 6
    OnClick = Button3Click
  end
end
