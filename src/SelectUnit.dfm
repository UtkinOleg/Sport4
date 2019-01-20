object SelectForm: TSelectForm
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #1069#1082#1089#1087#1086#1088#1090
  ClientHeight = 119
  ClientWidth = 307
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object ButtonNext: TButton
    Left = 143
    Top = 87
    Width = 75
    Height = 25
    Caption = #1044#1072#1083#1077#1077
    ModalResult = 1
    TabOrder = 0
  end
  object CancelBtn: TButton
    Left = 224
    Top = 86
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 291
    Height = 73
    Caption = #1059#1082#1072#1078#1080#1090#1077' '#1082#1072#1082#1091#1102' '#1080#1085#1092#1086#1088#1084#1072#1094#1080#1102' '#1089#1083#1077#1076#1091#1077#1090' '#1101#1082#1089#1087#1086#1088#1090#1080#1088#1086#1074#1072#1090#1100
    TabOrder = 2
    object ComboBox1: TComboBox
      Left = 73
      Top = 27
      Width = 137
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
    end
  end
end
