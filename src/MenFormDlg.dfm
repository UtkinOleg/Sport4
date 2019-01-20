inherited FIODlg: TFIODlg
  Caption = #1057#1087#1086#1088#1090#1089#1084#1077#1085
  ClientHeight = 274
  ClientWidth = 568
  OnShow = FormShow
  ExplicitWidth = 574
  ExplicitHeight = 306
  PixelsPerInch = 96
  TextHeight = 13
  inherited Bevel1: TBevel
    Width = 552
    Height = 225
    ExplicitWidth = 552
    ExplicitHeight = 225
  end
  object Label1: TLabel [1]
    Left = 24
    Top = 147
    Width = 36
    Height = 13
    Caption = #1056#1072#1079#1088#1103#1076
  end
  object Label2: TLabel [2]
    Left = 288
    Top = 21
    Width = 65
    Height = 13
    Caption = #1042#1080#1076#1099' '#1089#1087#1086#1088#1090#1072
  end
  inherited OKBtn: TButton
    Left = 326
    Top = 239
    ModalResult = 0
    OnClick = OKBtnClick
    ExplicitLeft = 326
    ExplicitTop = 239
  end
  inherited CancelBtn: TButton
    Left = 406
    Top = 239
    Caption = #1054#1090#1084#1077#1085#1072
    ExplicitLeft = 406
    ExplicitTop = 239
  end
  object HelpBtn: TButton
    Left = 486
    Top = 239
    Width = 75
    Height = 25
    Caption = '&'#1055#1086#1084#1086#1097#1100
    TabOrder = 2
    OnClick = HelpBtnClick
  end
  object LabeledEdit1: TLabeledEdit
    Left = 24
    Top = 40
    Width = 241
    Height = 21
    EditLabel.Width = 44
    EditLabel.Height = 13
    EditLabel.Caption = #1060#1072#1084#1080#1083#1080#1103
    EditLabel.Transparent = True
    TabOrder = 3
  end
  object LabeledEdit2: TLabeledEdit
    Left = 24
    Top = 80
    Width = 241
    Height = 21
    EditLabel.Width = 19
    EditLabel.Height = 13
    EditLabel.Caption = #1048#1084#1103
    EditLabel.Transparent = True
    TabOrder = 4
  end
  object LabeledEdit3: TLabeledEdit
    Left = 24
    Top = 120
    Width = 241
    Height = 21
    EditLabel.Width = 49
    EditLabel.Height = 13
    EditLabel.Caption = #1054#1090#1095#1077#1089#1090#1074#1086
    EditLabel.Transparent = True
    TabOrder = 5
  end
  object ComboBox1: TComboBox
    Left = 24
    Top = 166
    Width = 241
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 6
  end
  object ListBox1: TListBox
    Left = 288
    Top = 40
    Width = 257
    Height = 147
    ItemHeight = 13
    TabOrder = 7
  end
  object Button1: TButton
    Left = 389
    Top = 196
    Width = 75
    Height = 25
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100
    TabOrder = 8
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 470
    Top = 196
    Width = 75
    Height = 25
    Caption = #1059#1076#1072#1083#1080#1090#1100
    TabOrder = 9
    OnClick = Button2Click
  end
end
