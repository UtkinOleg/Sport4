unit OutputUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  ComObj, Dialogs, TeEngine, Series, ExtCtrls, TeeProcs, Chart, Grids, ComCtrls, ToolWin,
  StdCtrls;

type
  TOutputForm = class(TForm)
    Panel1: TPanel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    Chart1: TChart;
    Series1: TLineSeries;
    Series2: TLineSeries;
    Chart3: TChart;
    LineSeries3: TPieSeries;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet6: TTabSheet;
    Chart4: TChart;
    lv: TListView;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    Series3: TBarSeries;
    LineSeries4: TAreaSeries;
    Chart5: TChart;
    AreaSeries1: TBarSeries;
    Chart6: TChart;
    BarSeries1: TBarSeries;
    PageControl2: TPageControl;
    TabSheet7: TTabSheet;
    TabSheet8: TTabSheet;
    StringGrid1: TStringGrid;
    lv2: TListView;
    Chart2: TChart;
    LineSeries1: TLineSeries;
    LineSeries2: TLineSeries;
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ToolButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure ActivateExcel;
    procedure ActivateExcel_TopTable;
    procedure ActivateExcel_BottomTable;
  end;

var
  OutputForm: TOutputForm;

implementation

uses Main;

{$R *.dfm}

function ReplaceStr(const S, Srch, Replace: string): string;
var
  I: Integer;
  Source: string;
begin
  Source := S;
  Result := '';
  repeat
    I := Pos(Srch, Source);
    if I > 0 then begin
      Result := Result + Copy(Source, 1, I - 1) + Replace;
      Source := Copy(Source, I + Length(Srch), MaxInt);
    end
    else Result := Result + Source;
  until I <= 0;
end;

Function ConvertString(S:string):string;
Begin
  Result := StringReplace(S, '-', #173,[]);
end;

procedure TOutputForm.ActivateExcel_TopTable;
var
  Area, Excel, Sheet: Variant;
  len, i,j,jj : integer;
begin
  try
    Excel := GetActiveOleObject('Excel.Application');
    Excel.Visible := False;
  except
    try
      Excel := CreateOleObject('Excel.Application');
      Excel.Visible := False;
    except
      Exception.Create('Ошибка Microsoft Excel.');
      Exit;
    end;
  end;
  Excel.WorkBooks.Add;
  Excel.DisplayAlerts := False;
  Excel.Visible:= false;

  Excel.WorkBooks[1].WorkSheets[1].Cells[1, 1] := 'Параметр';
  Excel.WorkBooks[1].WorkSheets[1].Cells[1, 2] := 'Значение';
  Excel.WorkBooks[1].WorkSheets[1].Columns[1].ColumnWidth := 60;
  Excel.WorkBooks[1].WorkSheets[1].Columns[2].ColumnWidth := 20;

  for I := 0 to lv.Items.Count - 1 do
   begin
     Excel.WorkBooks[1].WorkSheets[1].Cells[i+2, 1] := lv.Items[i].Caption;
     if lv.Items[i].Subitems.Count >0 then
      Excel.WorkBooks[1].WorkSheets[1].Cells[i+2, 2] := ReplaceStr(lv.Items[i].Subitems[0],',','.');
   end;
  Excel.Visible:= true;

end;

procedure TOutputForm.ActivateExcel_BottomTable;
var
  Area, Excel, Sheet: Variant;
  len, i,j,jj : integer;
begin
  try
    Excel := GetActiveOleObject('Excel.Application');
    Excel.Visible := False;
  except
    try
      Excel := CreateOleObject('Excel.Application');
      Excel.Visible := False;
    except
      Exception.Create('Ошибка Microsoft Excel.');
      Exit;
    end;
  end;
  Excel.WorkBooks.Add;
  Excel.DisplayAlerts := False;
  Excel.Visible:= false;

  Excel.WorkBooks[1].WorkSheets[1].Columns[1].ColumnWidth := 20;
  Excel.WorkBooks[1].WorkSheets[1].Columns[2].ColumnWidth := 20;
  Excel.WorkBooks[1].WorkSheets[1].Columns[3].ColumnWidth := 20;
  Excel.WorkBooks[1].WorkSheets[1].Columns[4].ColumnWidth := 20;

  for I := 0 to StringGrid1.RowCount - 1 do
   begin
     Excel.WorkBooks[1].WorkSheets[1].Cells[i+1, 1] := ReplaceStr(StringGrid1.Cells[0,i],',','.');
     Excel.WorkBooks[1].WorkSheets[1].Cells[i+1, 2] := ReplaceStr(StringGrid1.Cells[1,i],',','.');
     Excel.WorkBooks[1].WorkSheets[1].Cells[i+1, 3] := ReplaceStr(StringGrid1.Cells[2,i],',','.');
     Excel.WorkBooks[1].WorkSheets[1].Cells[i+1, 4] := ReplaceStr(StringGrid1.Cells[3,i],',','.');
   end;

  Excel.Visible:= true;

end;


procedure TOutputForm.ActivateExcel;
var
  Area, Excel, Sheet, Chart: Variant;
  len, i,j,jj : integer;
  dt : TDateTime;
  execfound, ofpfound, manfound : boolean;
  summ_ofpspeed, summ_ofpvolume, ofpspeed, ofpvolume, summ_speed, summ_volume, speed, volume : real;
begin
  try
    Excel := GetActiveOleObject('Excel.Application');
    Excel.Visible := False;
  except
    try
      Excel := CreateOleObject('Excel.Application');
      Excel.Visible := False;
    except
      Exception.Create('Ошибка Microsoft Excel.');
      Exit;
    end;
  end;
  Excel.DisplayAlerts := False;
  Sheet := Excel.WorkBooks.Add;
  Chart := Excel.Charts.Add;

  for i:=1 to 10 do
   for j:=1 to 10 do
     Excel.WorkBooks[1].WorkSheets[1].Cells[i,j] := '';

{  if PageControl1.ActivePageIndex = 0 then
      Chart.ChartType := 4
  else
  if PageControl1.ActivePageIndex = 1 then
      Chart.ChartType := 65
  else     // 1 - ãðàôèê  , 5 - êðóãîâàÿ
  if PageControl1.ActivePageIndex = 2 then
      Chart.ChartType := -4102;  }

  Chart.SeriesCollection.NewSeries;

  Chart.HasLegend := True;

//  Chart.SeriesCollection(1).AxisGroup := '2';
  Chart.SeriesCollection(1).XValues := ConvertString('=Лист1!R1C1:R'+IntToStr(StringGrid1.RowCount-1)+'C1');
  Chart.SeriesCollection(1).Values := ConvertString('=Лист1!R1C2:R'+IntToStr(StringGrid1.RowCount-1)+'C2');
//  Chart.SeriesCollection(1).Values := ConvertString('=Лист1!R1C1:R'+IntToStr(StringGrid1.ColCount)+'C2');
//  Chart.SeriesCollection(1).Values := ConvertString('=Лист!R1C2:R'+IntToStr(StringGrid1.ColCount)+'C2');
//  Chart.SeriesCollection(1).Select;

  Chart.SeriesCollection.NewSeries;
//  Chart.SeriesCollection(2).Select;
  Chart.SeriesCollection(2).XValues := ConvertString('=Лист1!R1C1:R'+IntToStr(StringGrid1.RowCount-1)+'C1');
  Chart.SeriesCollection(2).Values := ConvertString('=Лист1!R1C3:R'+IntToStr(StringGrid1.RowCount-1)+'C3');
//  Chart.SeriesCollection(2).Values := ConvertString('=Лист!R1C3:R'+IntToStr(StringGrid1.ColCount)+'C3');
  Chart.SeriesCollection(2).AxisGroup := 2;

{  Chart.Axes.item(2).HasTitle := True;
  Chart.Axes.item(2).AxisTitle.Characters.Text := 'Объем';
  Chart.Axes.item(2).AxisTitle.Characters.Font.Size := 10;
  Chart.Axes.item(2).TickLabels.Font.Size := 9;
  Chart.Axes.item(2,2).HasTitle := True;
  Chart.Axes.item(2,2).AxisTitle.Characters.Text := 'Интенсивность';
  Chart.Axes.item(2,2).AxisTitle.Characters.Font.Size := 10;
  Chart.Axes.item(2,2).TickLabels.Font.Size := 9;}

//  Chart.ChartType := 4;


  summ_speed := 0;
  summ_volume := 0;
  summ_ofpspeed := 0;
  summ_ofpvolume := 0;

  for I := 0 to StringGrid1.RowCount - 1 do
   begin
     Excel.WorkBooks[1].WorkSheets[1].Cells[i+1, 1] := ReplaceStr(StringGrid1.Cells[0,i+1],',','.');
     Excel.WorkBooks[1].WorkSheets[1].Cells[i+1, 2] := ReplaceStr(StringGrid1.Cells[1,i+1],',','.');
     Excel.WorkBooks[1].WorkSheets[1].Cells[i+1, 3] := ReplaceStr(StringGrid1.Cells[2,i+1],',','.');
   end;

   Chart.ChartType := 4;
   Chart.SeriesCollection(2).AxisGroup := 2;
   Chart.SeriesCollection(1).Name := 'Объем';
   Chart.SeriesCollection(2).Name := 'Интенсивность';
//   Chart.Axes.item(2,2).AxisTitle.Characters.Text := 'Интенсивность';

   Excel.Visible:= true;
end;


procedure TOutputForm.FormActivate(Sender: TObject);
var
 i : integer;
begin
 for I := 0 to MainForm.DockTabSet1.Tabs.Count - 1 do
  if Caption = Trim(MainForm.DockTabSet1.Tabs[i]) then
  begin
    MainForm.DockTabSet1.TabIndex := i;
    break;
  end;

 MainForm.N18.Enabled := false;
 MainForm.N9.Enabled := true;
 MainForm.ToolButton2.Enabled := true;
 MainForm.FileSaveItem.Enabled := true;
 MainForm.C1.Enabled := false;
 MainForm.ToolButton8.Enabled := false;
 MainForm.N14.Enabled := false;

 MainForm.ToolButton14.Enabled := true;
 MainForm.N24.Enabled := true;

 MainForm.ToolButton11.Enabled := false;
 MainForm.N16.Enabled := false;
 MainForm.ToolButton12.Enabled := false;
 MainForm.N17.Enabled := false;
 MainForm.SB.Panels[0].Text := '';
 MainForm.N27.Enabled := true;
 MainForm.ToolButton4.Enabled := true;

end;

procedure TOutputForm.FormClose(Sender: TObject; var Action: TCloseAction);
var
 i : integer;
begin
  for I := 0 to MainForm.DockTabSet1.Tabs.Count - 1 do
   if Trim(MainForm.DockTabSet1.Tabs[i]) = Caption then
    begin
     MainForm.DockTabSet1.Tabs.Delete(i);
     break;
    end;

  MainForm.ToolButton14.Enabled := false;
  MainForm.N9.Enabled := false;
  MainForm.ToolButton4.Enabled := false;
  MainForm.N27.Enabled := false;

  if MainForm.MDIChildCount = 1 then
   begin
    MainForm.ToolButton2.Enabled := false;
    MainForm.FileSaveItem.Enabled := false;
    MainForm.DockTabSet1.Visible := false;
    MainForm.N18.Enabled := false;
    MainForm.ToolButton14.Enabled := false;
    MainForm.N24.Enabled := false;
   end;

 Action := caFree;
end;

procedure TOutputForm.ToolButton1Click(Sender: TObject);
begin
 Chart1.Print;
end;

end.
