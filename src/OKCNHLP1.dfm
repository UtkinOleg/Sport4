inherited OKHelpBottomDlg: TOKHelpBottomDlg
  Caption = #1042#1074#1077#1076#1080#1090#1077
  ClientHeight = 105
  ClientWidth = 563
  OnShow = FormShow
  ExplicitWidth = 569
  ExplicitHeight = 137
  PixelsPerInch = 96
  TextHeight = 13
  inherited Bevel1: TBevel
    Width = 547
    Height = 57
    ExplicitWidth = 547
    ExplicitHeight = 57
  end
  inherited OKBtn: TButton
    Left = 318
    Top = 72
    ModalResult = 0
    OnClick = OKBtnClick
    ExplicitLeft = 318
    ExplicitTop = 72
  end
  inherited CancelBtn: TButton
    Left = 398
    Top = 72
    Caption = #1054#1090#1084#1077#1085#1072
    ExplicitLeft = 398
    ExplicitTop = 72
  end
  object HelpBtn: TButton
    Left = 478
    Top = 72
    Width = 75
    Height = 25
    Caption = '&'#1055#1086#1084#1086#1097#1100
    TabOrder = 2
    OnClick = HelpBtnClick
  end
  object Edit1: TEdit
    Left = 24
    Top = 24
    Width = 513
    Height = 21
    TabOrder = 3
  end
end
