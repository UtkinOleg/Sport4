object SelectSportDlg: TSelectSportDlg
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1042#1099#1073#1086#1088' '#1074#1080#1076#1072' '#1089#1087#1086#1088#1090#1072
  ClientHeight = 352
  ClientWidth = 377
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 8
    Top = 8
    Width = 281
    Height = 336
    Shape = bsFrame
  end
  object OKBtn: TButton
    Left = 296
    Top = 288
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 296
    Top = 318
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
  end
  object ListBox1: TListBox
    Left = 16
    Top = 16
    Width = 265
    Height = 321
    ItemHeight = 13
    TabOrder = 2
    OnDblClick = ListBox1DblClick
  end
end
