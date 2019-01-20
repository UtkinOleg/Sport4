object OptionForm: TOptionForm
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1080
  ClientHeight = 160
  ClientWidth = 428
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PrintScale = poNone
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object OKBtn: TButton
    Left = 181
    Top = 127
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
  end
  object CancelBtn: TButton
    Left = 262
    Top = 127
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
  end
  object HelpBtn: TButton
    Left = 342
    Top = 127
    Width = 75
    Height = 25
    Caption = '&'#1055#1086#1084#1086#1097#1100
    TabOrder = 2
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 410
    Height = 113
    Caption = #1054#1073#1085#1086#1074#1083#1077#1085#1080#1077
    TabOrder = 3
    object Label1: TLabel
      Left = 16
      Top = 21
      Width = 117
      Height = 13
      Caption = 'URL '#1092#1072#1081#1083#1072' '#1086#1073#1085#1086#1074#1083#1077#1085#1080#1081
      Transparent = True
    end
    object Edit1: TEdit
      Left = 16
      Top = 40
      Width = 377
      Height = 21
      TabOrder = 0
      Text = 'http://'
    end
    object Button1: TButton
      Left = 16
      Top = 72
      Width = 137
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1090#1100' '#1086#1073#1085#1086#1074#1083#1077#1085#1080#1077
      TabOrder = 1
      OnClick = Button1Click
    end
  end
  object SaveDialog1: TSaveDialog
    Left = 120
    Top = 128
  end
end
