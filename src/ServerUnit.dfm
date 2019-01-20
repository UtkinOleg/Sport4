object ServerForm: TServerForm
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #1057#1086#1076#1077#1088#1078#1072#1085#1080#1077' '#1089#1077#1088#1074#1077#1088#1072' Startedu.ru'
  ClientHeight = 442
  ClientWidth = 732
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
  object ListView1: TListView
    Left = 8
    Top = 8
    Width = 716
    Height = 393
    Checkboxes = True
    Columns = <
      item
        Caption = #1057#1087#1086#1088#1090#1089#1084#1077#1085
        Width = 300
      end
      item
        Caption = #1060#1072#1081#1083
        Width = 200
      end
      item
        Caption = #1044#1072#1090#1072
        Width = 70
      end
      item
        Caption = #1056#1072#1079#1084#1077#1088
        Width = 70
      end>
    ReadOnly = True
    RowSelect = True
    TabOrder = 0
    ViewStyle = vsReport
  end
  object Button1: TButton
    Left = 568
    Top = 408
    Width = 75
    Height = 25
    Caption = 'OK'
    ModalResult = 1
    TabOrder = 1
  end
  object Button2: TButton
    Left = 649
    Top = 408
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 2
  end
end
