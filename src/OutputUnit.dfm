object OutputForm: TOutputForm
  Left = 0
  Top = 0
  Caption = #1040#1085#1072#1083#1080#1079
  ClientHeight = 472
  ClientWidth = 699
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Visible = True
  WindowState = wsMaximized
  OnActivate = FormActivate
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 113
    Width = 699
    Height = 3
    Cursor = crVSplit
    Align = alTop
    Beveled = True
    ExplicitTop = 118
  end
  object Splitter2: TSplitter
    Left = 0
    Top = 309
    Width = 699
    Height = 3
    Cursor = crVSplit
    Align = alBottom
    Beveled = True
    ExplicitTop = 116
    ExplicitWidth = 276
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 699
    Height = 113
    Align = alTop
    TabOrder = 0
    object lv: TListView
      Left = 1
      Top = 1
      Width = 697
      Height = 111
      Align = alClient
      Columns = <
        item
          Caption = #1055#1072#1088#1072#1084#1077#1090#1088
          Width = 300
        end
        item
          AutoSize = True
          Caption = #1047#1085#1072#1095#1077#1085#1080#1077
        end>
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      GridLines = True
      HideSelection = False
      RowSelect = True
      ParentFont = False
      TabOrder = 0
      ViewStyle = vsReport
    end
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 116
    Width = 699
    Height = 193
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 1
    object TabSheet1: TTabSheet
      Caption = #1057#1086#1086#1090#1085#1086#1096#1077#1085#1080#1077' '#1085#1072#1075#1088#1091#1079#1082#1080' '#1087#1086' '#1087#1077#1088#1080#1086#1076#1072#1084
      object Chart1: TChart
        Left = 0
        Top = 0
        Width = 691
        Height = 165
        AllowPanning = pmNone
        BackWall.Brush.Color = clWhite
        BackWall.Brush.Style = bsClear
        Foot.Visible = False
        LeftWall.Brush.Color = clWhite
        Legend.Alignment = laBottom
        Legend.Frame.Visible = False
        Title.Text.Strings = (
          'TChart')
        Title.Visible = False
        BottomAxis.Grid.Visible = False
        BottomAxis.GridCentered = True
        BottomAxis.LabelsAngle = 90
        BottomAxis.LabelsMultiLine = True
        BottomAxis.LabelsSeparation = 0
        ClipPoints = False
        LeftAxis.Grid.Visible = False
        LeftAxis.TicksInner.Visible = False
        LeftAxis.TickOnLabelsOnly = False
        LeftAxis.Title.Caption = #1054#1073#1098#1077#1084
        LeftAxis.Title.Font.Color = clRed
        LeftAxis.Title.Font.Style = [fsBold]
        RightAxis.Grid.Visible = False
        RightAxis.Title.Caption = #1048#1085#1090#1077#1085#1089#1080#1074#1085#1086#1089#1090#1100
        RightAxis.Title.Font.Color = clGreen
        RightAxis.Title.Font.Style = [fsBold]
        View3D = False
        Zoom.Allow = False
        Align = alClient
        Color = clBtnHighlight
        TabOrder = 0
        object Series1: TLineSeries
          Marks.Arrow.Visible = True
          Marks.Callout.Brush.Color = clBlack
          Marks.Callout.Arrow.Visible = True
          Marks.Visible = False
          Title = #1054#1073#1098#1077#1084
          Pointer.HorizSize = 3
          Pointer.InflateMargins = True
          Pointer.Style = psCircle
          Pointer.VertSize = 3
          Pointer.Visible = True
          XValues.Name = 'X'
          XValues.Order = loAscending
          YValues.Name = 'Y'
          YValues.Order = loNone
        end
        object Series2: TLineSeries
          Marks.Arrow.Visible = True
          Marks.Callout.Brush.Color = clBlack
          Marks.Callout.Arrow.Visible = True
          Marks.Visible = False
          SeriesColor = clBlue
          Title = #1048#1085#1090#1077#1085#1089#1080#1074#1085#1086#1089#1090#1100
          VertAxis = aRightAxis
          Pointer.HorizSize = 3
          Pointer.InflateMargins = True
          Pointer.Style = psCircle
          Pointer.VertSize = 3
          Pointer.Visible = True
          XValues.Name = 'X'
          XValues.Order = loAscending
          YValues.Name = 'Y'
          YValues.Order = loNone
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1054#1073#1098#1077#1084' '#1080' '#1080#1085#1090#1077#1085#1089#1080#1074#1085#1086#1089#1090#1100' '#1079#1072' '#1052#1047#1062
      ImageIndex = 1
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 189
      object Chart2: TChart
        Left = 0
        Top = 0
        Width = 691
        Height = 165
        AllowPanning = pmNone
        BackWall.Brush.Color = clWhite
        BackWall.Brush.Style = bsClear
        Foot.Visible = False
        LeftWall.Brush.Color = clWhite
        Legend.Alignment = laBottom
        Legend.Frame.Visible = False
        Title.Text.Strings = (
          'TChart')
        Title.Visible = False
        BottomAxis.Grid.Visible = False
        BottomAxis.GridCentered = True
        BottomAxis.LabelsAngle = 90
        BottomAxis.LabelsMultiLine = True
        BottomAxis.LabelsSeparation = 0
        ClipPoints = False
        LeftAxis.Grid.Visible = False
        LeftAxis.TicksInner.Visible = False
        LeftAxis.TickOnLabelsOnly = False
        LeftAxis.Title.Caption = #1054#1073#1098#1077#1084
        LeftAxis.Title.Font.Color = clRed
        LeftAxis.Title.Font.Style = [fsBold]
        RightAxis.Grid.Visible = False
        RightAxis.Title.Caption = #1048#1085#1090#1077#1085#1089#1080#1074#1085#1086#1089#1090#1100
        RightAxis.Title.Font.Color = clGreen
        RightAxis.Title.Font.Style = [fsBold]
        View3D = False
        Zoom.Allow = False
        Align = alClient
        Color = clBtnHighlight
        TabOrder = 0
        ExplicitHeight = 189
        object LineSeries1: TLineSeries
          Marks.Arrow.Visible = True
          Marks.Callout.Brush.Color = clBlack
          Marks.Callout.Arrow.Visible = True
          Marks.Visible = False
          Title = #1054#1073#1098#1077#1084
          Pointer.HorizSize = 3
          Pointer.InflateMargins = True
          Pointer.Style = psCircle
          Pointer.VertSize = 3
          Pointer.Visible = True
          XValues.Name = 'X'
          XValues.Order = loAscending
          YValues.Name = 'Y'
          YValues.Order = loNone
        end
        object LineSeries2: TLineSeries
          Marks.Arrow.Visible = True
          Marks.Callout.Brush.Color = clBlack
          Marks.Callout.Arrow.Visible = True
          Marks.Visible = False
          SeriesColor = clBlue
          Title = #1048#1085#1090#1077#1085#1089#1080#1074#1085#1086#1089#1090#1100
          VertAxis = aRightAxis
          Pointer.HorizSize = 3
          Pointer.InflateMargins = True
          Pointer.Style = psCircle
          Pointer.VertSize = 3
          Pointer.Visible = True
          XValues.Name = 'X'
          XValues.Order = loAscending
          YValues.Name = 'Y'
          YValues.Order = loNone
        end
      end
    end
    object TabSheet3: TTabSheet
      Caption = #1057#1086#1086#1090#1085#1086#1096#1077#1085#1080#1077' '#1091#1087#1088#1072#1078#1085#1077#1085#1080#1081
      ImageIndex = 2
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 189
      object Chart3: TChart
        Left = 0
        Top = 0
        Width = 691
        Height = 165
        AllowPanning = pmNone
        BackWall.Brush.Color = clWhite
        BackWall.Brush.Style = bsClear
        BackWall.Pen.Visible = False
        Foot.Visible = False
        LeftWall.Brush.Color = clWhite
        Legend.Alignment = laBottom
        Title.Text.Strings = (
          'TChart')
        Title.Visible = False
        AxisVisible = False
        BottomAxis.LabelsAngle = 90
        BottomAxis.LabelsMultiLine = True
        BottomAxis.LabelsOnAxis = False
        BottomAxis.LabelsSeparation = 0
        BottomAxis.RoundFirstLabel = False
        ClipPoints = False
        Frame.Visible = False
        LeftAxis.Automatic = False
        LeftAxis.AutomaticMinimum = False
        LeftAxis.Title.Caption = #1054#1073#1098#1077#1084
        LeftAxis.Title.Font.Color = clRed
        LeftAxis.Title.Font.Style = [fsBold]
        RightAxis.Automatic = False
        RightAxis.AutomaticMinimum = False
        RightAxis.Title.Caption = #1048#1085#1090#1077#1085#1089#1080#1074#1085#1086#1089#1090#1100
        RightAxis.Title.Font.Color = clGreen
        RightAxis.Title.Font.Style = [fsBold]
        View3DOptions.Elevation = 315
        View3DOptions.Orthogonal = False
        View3DOptions.Perspective = 0
        View3DOptions.Rotation = 360
        View3DWalls = False
        Zoom.Allow = False
        Align = alClient
        Color = clSilver
        TabOrder = 0
        ExplicitHeight = 189
        object LineSeries3: TPieSeries
          Marks.Arrow.Visible = True
          Marks.Callout.Brush.Color = clBlack
          Marks.Callout.Arrow.Visible = True
          Marks.Style = smsLabelPercent
          Marks.Visible = True
          Title = #1054#1073#1098#1077#1084
          Gradient.Direction = gdRadial
          OtherSlice.Legend.Visible = False
          OtherSlice.Text = 'Other'
          PieValues.Name = 'Pie'
          PieValues.Order = loNone
        end
      end
    end
    object TabSheet4: TTabSheet
      Caption = #1057#1086#1086#1090#1085#1086#1096#1077#1085#1080#1077' '#1057#1060#1055'/'#1054#1060#1055
      ImageIndex = 3
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 189
      object Chart4: TChart
        Left = 0
        Top = 0
        Width = 691
        Height = 165
        AllowPanning = pmNone
        BackWall.Brush.Color = clWhite
        BackWall.Brush.Style = bsClear
        Foot.Visible = False
        LeftWall.Brush.Color = clWhite
        Legend.ColorWidth = 30
        Legend.DividingLines.Visible = True
        Legend.Frame.Visible = False
        Legend.Symbol.Width = 30
        Title.Text.Strings = (
          'TChart')
        Title.Visible = False
        BottomAxis.Grid.Visible = False
        BottomAxis.GridCentered = True
        BottomAxis.LabelsAngle = 90
        BottomAxis.LabelsMultiLine = True
        BottomAxis.LabelsSeparation = 0
        ClipPoints = False
        LeftAxis.Grid.Visible = False
        LeftAxis.TicksInner.Visible = False
        LeftAxis.TickOnLabelsOnly = False
        LeftAxis.Title.Caption = #1054#1073#1098#1077#1084
        LeftAxis.Title.Font.Color = clRed
        LeftAxis.Title.Font.Style = [fsBold]
        RightAxis.Grid.Visible = False
        RightAxis.Title.Caption = #1048#1085#1090#1077#1085#1089#1080#1074#1085#1086#1089#1090#1100
        RightAxis.Title.Font.Color = clGreen
        RightAxis.Title.Font.Style = [fsBold]
        Zoom.Allow = False
        Align = alClient
        Color = clBtnHighlight
        TabOrder = 0
        ExplicitHeight = 189
        object LineSeries4: TAreaSeries
          Marks.Arrow.Visible = True
          Marks.Callout.Brush.Color = clBlack
          Marks.Callout.Arrow.Visible = True
          Marks.Visible = False
          Title = #1057#1060#1055
          DrawArea = True
          Pointer.HorizSize = 3
          Pointer.InflateMargins = True
          Pointer.Style = psCircle
          Pointer.VertSize = 3
          Pointer.Visible = False
          XValues.Name = 'X'
          XValues.Order = loAscending
          YValues.Name = 'Y'
          YValues.Order = loNone
        end
        object Series3: TBarSeries
          Marks.Arrow.Visible = True
          Marks.Callout.Brush.Color = clBlack
          Marks.Callout.Arrow.Visible = True
          Marks.Visible = False
          SeriesColor = clBlue
          Title = #1054#1060#1055
          BarWidthPercent = 25
          Gradient.Direction = gdTopBottom
          XValues.Name = 'X'
          XValues.Order = loAscending
          YValues.Name = 'Bar'
          YValues.Order = loNone
        end
      end
    end
    object TabSheet5: TTabSheet
      Caption = #1042#1086#1089#1089#1090#1072#1085#1086#1074#1083#1077#1085#1080#1077
      ImageIndex = 4
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 189
      object Chart5: TChart
        Left = 0
        Top = 0
        Width = 691
        Height = 165
        AllowPanning = pmNone
        BackWall.Brush.Color = clWhite
        BackWall.Brush.Style = bsClear
        Foot.Visible = False
        LeftWall.Brush.Color = clWhite
        Legend.ColorWidth = 30
        Legend.DividingLines.Visible = True
        Legend.Frame.Visible = False
        Legend.Symbol.Width = 30
        Title.Text.Strings = (
          'TChart')
        Title.Visible = False
        BottomAxis.Grid.Visible = False
        BottomAxis.GridCentered = True
        BottomAxis.LabelsAngle = 90
        BottomAxis.LabelsMultiLine = True
        BottomAxis.LabelsSeparation = 0
        ClipPoints = False
        LeftAxis.Grid.Visible = False
        LeftAxis.TicksInner.Visible = False
        LeftAxis.TickOnLabelsOnly = False
        LeftAxis.Title.Font.Color = clRed
        LeftAxis.Title.Font.Style = [fsBold]
        RightAxis.Grid.Visible = False
        RightAxis.Title.Caption = #1048#1085#1090#1077#1085#1089#1080#1074#1085#1086#1089#1090#1100
        RightAxis.Title.Font.Color = clGreen
        RightAxis.Title.Font.Style = [fsBold]
        Zoom.Allow = False
        Align = alClient
        Color = clBtnHighlight
        TabOrder = 0
        ExplicitHeight = 189
        object AreaSeries1: TBarSeries
          Marks.Arrow.Visible = True
          Marks.Callout.Brush.Color = clBlack
          Marks.Callout.Arrow.Visible = True
          Marks.Callout.Length = 8
          Marks.Visible = False
          SeriesColor = clLime
          ShowInLegend = False
          Title = #1042#1086#1089#1089#1090#1072#1085#1086#1074#1083#1077#1085#1080#1077
          Gradient.Direction = gdTopBottom
          XValues.Name = 'X'
          XValues.Order = loAscending
          YValues.Name = 'Bar'
          YValues.Order = loNone
        end
      end
    end
    object TabSheet6: TTabSheet
      Caption = #1050#1086#1085#1090#1088#1086#1083#1100
      ImageIndex = 5
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 189
      object Chart6: TChart
        Left = 0
        Top = 0
        Width = 691
        Height = 165
        AllowPanning = pmNone
        BackWall.Brush.Color = clWhite
        BackWall.Brush.Style = bsClear
        Foot.Visible = False
        LeftWall.Brush.Color = clWhite
        Legend.ColorWidth = 30
        Legend.DividingLines.Visible = True
        Legend.Frame.Visible = False
        Legend.Symbol.Width = 30
        Title.Text.Strings = (
          'TChart')
        BottomAxis.Grid.Visible = False
        BottomAxis.GridCentered = True
        BottomAxis.LabelsAngle = 90
        BottomAxis.LabelsMultiLine = True
        BottomAxis.LabelsSeparation = 0
        ClipPoints = False
        LeftAxis.Grid.Visible = False
        LeftAxis.TicksInner.Visible = False
        LeftAxis.TickOnLabelsOnly = False
        LeftAxis.Title.Font.Color = clRed
        LeftAxis.Title.Font.Style = [fsBold]
        RightAxis.Grid.Visible = False
        RightAxis.Title.Caption = #1048#1085#1090#1077#1085#1089#1080#1074#1085#1086#1089#1090#1100
        RightAxis.Title.Font.Color = clGreen
        RightAxis.Title.Font.Style = [fsBold]
        Zoom.Allow = False
        Align = alClient
        Color = clBtnHighlight
        TabOrder = 0
        ExplicitHeight = 189
        object BarSeries1: TBarSeries
          Marks.Arrow.Visible = True
          Marks.Callout.Brush.Color = clBlack
          Marks.Callout.Arrow.Visible = True
          Marks.Callout.Length = 8
          Marks.Visible = False
          SeriesColor = clBlue
          ShowInLegend = False
          Title = #1050#1086#1085#1090#1088#1086#1083#1100
          Gradient.Direction = gdTopBottom
          XValues.Name = 'X'
          XValues.Order = loAscending
          YValues.Name = 'Bar'
          YValues.Order = loNone
        end
      end
    end
  end
  object PageControl2: TPageControl
    Left = 0
    Top = 312
    Width = 699
    Height = 160
    ActivePage = TabSheet7
    Align = alBottom
    TabOrder = 2
    object TabSheet7: TTabSheet
      Caption = #1042#1077#1083#1080#1095#1080#1085#1099' '#1085#1072#1075#1088#1091#1079#1082#1080
      object lv2: TListView
        Left = 0
        Top = 0
        Width = 691
        Height = 132
        Align = alClient
        Columns = <
          item
            Caption = #1055#1086#1082#1072#1079#1072#1090#1077#1083#1100
            Width = 300
          end
          item
            Caption = #1047#1085#1072#1095#1077#1085#1080#1077
            Width = 100
          end>
        GridLines = True
        HideSelection = False
        TabOrder = 0
        ViewStyle = vsReport
      end
    end
    object TabSheet8: TTabSheet
      Caption = #1042#1072#1088#1080#1072#1085#1090#1099' '#1088#1072#1089#1087#1088#1077#1076#1077#1083#1077#1085#1080#1103' '#1085#1072#1075#1088#1091#1079#1082#1080
      ImageIndex = 1
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 281
      ExplicitHeight = 165
      object StringGrid1: TStringGrid
        Left = 0
        Top = 0
        Width = 691
        Height = 132
        Align = alClient
        ColCount = 6
        DefaultColWidth = 130
        DefaultRowHeight = 17
        FixedCols = 0
        RowCount = 2
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing]
        ScrollBars = ssVertical
        TabOrder = 0
        ExplicitLeft = -418
        ExplicitTop = 29
        ExplicitWidth = 699
        ExplicitHeight = 136
      end
    end
  end
end