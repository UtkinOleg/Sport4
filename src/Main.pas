unit MAIN;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, Menus,
  StdCtrls, Dialogs, Buttons, Messages, ExtCtrls, ComCtrls, StdActns,
  ComObj, OleCtnrs, OleServer, ActnList, ToolWin, ImgList, XPMan, ActnMan, ActnColorMaps,
  ExcelXP, Grids, Tabs, DockTabSet, Registry, Chart, WinInet, Printers,
  OleCtrls, HTMLUnit, Math, ActiveX, Sockets,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdExplicitTLSClientServerBase, IdFTP, IdFtpCommon, IdFTPListParseBase, IdFTPListParseUnix, About;

const
  ExecTitle = 'Виды спорта и упражнения';
  ManTitle = 'Спортсмены';
  RazrTitle = 'Разряды';
  ExecVersion = 1;
  RazrVersion = 1;
  ManVersion = 1;
  DBVersion = 4;
  MainTitle = 'Спорт 4.0.0';
  demostr = 'Демонстрационная версия';
  regstr = 'Зарегистрированния версия';
  thstr = 'Спасибо за регистрацию программы.';

{$IFDEF WIN32}
const
  BadDllLoad = 0;
{$ELSE}
const
  BadDllLoad = 32;
{$ENDIF}

type

  Month = record
    Micro, Name, Year : integer;
    Days : TStringList;
  end;
  pMonth = ^Month;

  Micro = record
    M : integer;
    V,I,VP,IP : real;
  end;
  pMicro = ^Micro;

  TMainForm = class(TForm)
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    FileOpenItem: TMenuItem;
    FileCloseItem: TMenuItem;
    Window1: TMenuItem;
    Help1: TMenuItem;
    N1: TMenuItem;
    FileExitItem: TMenuItem;
    WindowCascadeItem: TMenuItem;
    WindowTileItem: TMenuItem;
    WindowArrangeItem: TMenuItem;
    HelpAboutItem: TMenuItem;
    OpenDialog: TOpenDialog;
    FileSaveItem: TMenuItem;
    Edit1: TMenuItem;
    CutItem: TMenuItem;
    CopyItem: TMenuItem;
    PasteItem: TMenuItem;
    WindowMinimizeItem: TMenuItem;
    ActionList1: TActionList;
    EditCut1: TEditCut;
    EditCopy1: TEditCopy;
    EditPaste1: TEditPaste;
    FileNew1: TAction;
    FileSave1: TAction;
    FileExit1: TAction;
    FileOpen1: TAction;
    FileSaveAs1: TAction;
    WindowCascade1: TWindowCascade;
    WindowTileHorizontal1: TWindowTileHorizontal;
    WindowArrangeAll1: TWindowArrange;
    WindowMinimizeAll1: TWindowMinimizeAll;
    HelpAbout1: TAction;
    FileClose1: TWindowClose;
    WindowTileVertical1: TWindowTileVertical;
    WindowTileItem2: TMenuItem;
    ImageList1: TImageList;
    XPManifest1: TXPManifest;
    ControlBar1: TControlBar;
    ToolBar2: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    XPColorMap1: TXPColorMap;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    C1: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N10: TMenuItem;
    SaveDialog1: TSaveDialog;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    ToolButton8: TToolButton;
    N14: TMenuItem;
    N15: TMenuItem;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    N16: TMenuItem;
    N17: TMenuItem;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton14: TToolButton;
    DockTabSet1: TDockTabSet;
    sb: TStatusBar;
    N18: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    N21: TMenuItem;
    N22: TMenuItem;
    N23: TMenuItem;
    N24: TMenuItem;
    N25: TMenuItem;
    N9: TMenuItem;
    pd: TPrintDialog;
    psd: TPrinterSetupDialog;
    N26: TMenuItem;
    OpenDialog1: TOpenDialog;
    Timer1: TTimer;
    N27: TMenuItem;
    N28: TMenuItem;
    N29: TMenuItem;
    N30: TMenuItem;
    N31: TMenuItem;
    N32: TMenuItem;
    ToolButton4: TToolButton;
    IdFTP1: TIdFTP;
    IdFTP2: TIdFTP;
    IdFTP3: TIdFTP;
    Timer2: TTimer;
    procedure N30Click(Sender: TObject);
    procedure DockTabSet1DrawTab(Sender: TObject; TabCanvas: TCanvas; R: TRect;
      Index: Integer; Selected: Boolean);
    procedure N27Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure N26Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure N24Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure ToolButton12Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure ToolButton11Click(Sender: TObject);
    procedure N19Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N18Click(Sender: TObject);
    procedure ToolButton14Click(Sender: TObject);
    procedure DockTabSet1Change(Sender: TObject; NewTab: Integer;
      var AllowChange: Boolean);
    procedure ToolButton9Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure ToolButton8Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure C1Click(Sender: TObject);
    procedure FileSave1Execute(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure FileOpen1Execute(Sender: TObject);
    procedure HelpAbout1Execute(Sender: TObject);
    procedure FileExit1Execute(Sender: TObject);
    procedure ToolButton4Click(Sender: TObject);
    procedure IdFTP1Work(ASender: TObject; AWorkMode: TWorkMode;
      AWorkCount: Int64);
    procedure IdFTP1WorkEnd(ASender: TObject; AWorkMode: TWorkMode);
    procedure IdFTP3Work(ASender: TObject; AWorkMode: TWorkMode;
      AWorkCount: Int64);
    procedure IdFTP3WorkEnd(ASender: TObject; AWorkMode: TWorkMode);
    procedure Timer2Timer(Sender: TObject);
    procedure IdFTP3AfterGet(ASender: TObject; AStream: TStream);
  private
    { Private declarations }
    procedure CreateMDIChild(Name : string; filename: string);
    procedure CreatePreviewChild(Name : string);
    procedure RegistryS30(path : string);
  public
    { Public declarations }
    h12, h34 : Int64;
    ImportMode, AdminMode, RegDone : boolean;
    ImportFilename, ImportFileFIO, RegData, RegFIO : string;
    ImportFile : TFileStream;
    ImportFileMem : TMemorystream;
    UsersList : TStringList;
    AboutBox2 : TAboutBox;
    procedure GenerateData;
    procedure ExportData;
    procedure ExportExcel;
    procedure ImportData;
    procedure RegComplete;
    procedure ReadExecsToString(var st : TStringList);
    procedure ReadExecs2ToString(index : integer; var st : TStringList);
    procedure UpdateDataBase;
    function  IfRegOK : boolean;
    function  MonthNumber(m : string):integer;
    function  GetMonth(d : TDateTime) : string;
    function  IsUpdateRequared(fname : string) : boolean;
    procedure ProcessInternetUpdate;

  end;

var
  MainForm: TMainForm;
  tMonth : pMonth;
  tMicro : pMicro;


implementation

{$R *.dfm}


uses CHILDWIN, Spr, Man, RazrForm, DiagUnit, OutputUnit, RegUnit,
  OptionUnit, ProgressUnit, UnitPreview, SelectUnit, PbUnit, BaseUnit,
  ServerUnit;


procedure PrintStringGrid(Grid: TStringGrid; Title: string;
   Orientation: TPrinterOrientation);
 var
   P, I, J, YPos, XPos, HorzSize, VertSize: Integer;
   AnzSeiten, Seite, Zeilen, HeaderSize, FooterSize, ZeilenSize, FontHeight: Integer;
   mmx, mmy: Extended;
   Footer: string;
 begin
   //Kopfzeile, Fu?zeile, Zeilenabstand, Schriftgro?e festlegen
  HeaderSize := 100;
   FooterSize := 200;
   ZeilenSize := 36;
   FontHeight := 36;
   //Printer initializieren
  Printer.Orientation := Orientation;
   Printer.Title  := Title;
   Printer.BeginDoc;
   //Druck auf mm einstellen
  mmx := GetDeviceCaps(Printer.Canvas.Handle, PHYSICALWIDTH) /
     GetDeviceCaps(Printer.Canvas.Handle, LOGPIXELSX) * 25.4;
   mmy := GetDeviceCaps(Printer.Canvas.Handle, PHYSICALHEIGHT) /
     GetDeviceCaps(Printer.Canvas.Handle, LOGPIXELSY) * 25.4;

   VertSize := Trunc(mmy) * 10;
   HorzSize := Trunc(mmx) * 10;
   SetMapMode(Printer.Canvas.Handle, MM_LOMETRIC);

   //Zeilenanzahl festlegen
  Zeilen := (VertSize - HeaderSize - FooterSize) div ZeilenSize;
   //Seitenanzahl ermitteln
  if Grid.RowCount mod Zeilen <> 0 then
     AnzSeiten := Grid.RowCount div Zeilen + 1
   else
     AnzSeiten := Grid.RowCount div Zeilen;

   Seite := 1;
   //Grid Drucken
  for P := 1 to AnzSeiten do
   begin
     //Kopfzeile
    Printer.Canvas.Font.Height := 48;
     Printer.Canvas.TextOut((HorzSize div 2 - (Printer.Canvas.TextWidth(Title) div 2)),
       - 20,Title);
     Printer.Canvas.Pen.Width := 5;
     Printer.Canvas.MoveTo(0, - HeaderSize);
     Printer.Canvas.LineTo(HorzSize, - HeaderSize);
     //Fu?zeile
    Printer.Canvas.MoveTo(0, - VertSize + FooterSize);
     Printer.Canvas.LineTo(HorzSize, - VertSize + FooterSize);
     Printer.Canvas.Font.Height := 36;
     Footer := 'Страница: ' + IntToStr(Seite) + ' из ' + IntToStr(AnzSeiten);
     Printer.Canvas.TextOut((HorzSize div 2 - (Printer.Canvas.TextWidth(Footer) div 2)),
       - VertSize + 150,Footer);
     //Zeilen drucken
    Printer.Canvas.Font.Height := FontHeight;
     YPos := HeaderSize + 10;
     for I := 1 to Zeilen do
     begin
       if Grid.RowCount >= I + (Seite - 1) * Zeilen then
       begin
         XPos := 100;
         for J := 0 to Grid.ColCount - 1 do
         begin
           Printer.Canvas.TextOut(XPos, - YPos,
             Grid.Cells[J, I + (Seite - 1) * Zeilen - 1]);
           XPos := XPos + Grid.ColWidths[J] * 3;
         end;
         YPos := YPos + ZeilenSize;
       end;
     end;
     //Seite hinzufugen
    Inc(Seite);
     if Seite <= AnzSeiten then Printer.NewPage;
   end;
   Printer.EndDoc;
 end;

procedure PrintListView(lv: TListView; Title: string;
   Orientation: TPrinterOrientation);
 var
   P, I, J, YPos, XPos, HorzSize, VertSize: Integer;
   AnzSeiten, Seite, Zeilen, HeaderSize, FooterSize, ZeilenSize, FontHeight: Integer;
   mmx, mmy: Extended;
   Footer: string;
 begin
   //Kopfzeile, Fu?zeile, Zeilenabstand, Schriftgro?e festlegen
  HeaderSize := 100;
   FooterSize := 200;
   ZeilenSize := 36;
   FontHeight := 36;
   //Printer initializieren
  Printer.Orientation := Orientation;
   Printer.Title  := Title;
   Printer.BeginDoc;
   //Druck auf mm einstellen
  mmx := GetDeviceCaps(Printer.Canvas.Handle, PHYSICALWIDTH) /
     GetDeviceCaps(Printer.Canvas.Handle, LOGPIXELSX) * 25.4;
   mmy := GetDeviceCaps(Printer.Canvas.Handle, PHYSICALHEIGHT) /
     GetDeviceCaps(Printer.Canvas.Handle, LOGPIXELSY) * 25.4;

   VertSize := Trunc(mmy) * 10;
   HorzSize := Trunc(mmx) * 10;
   SetMapMode(Printer.Canvas.Handle, MM_LOMETRIC);

   //Zeilenanzahl festlegen
  Zeilen := (VertSize - HeaderSize - FooterSize) div ZeilenSize;
   //Seitenanzahl ermitteln
  if lv.Items.Count mod Zeilen <> 0 then
     AnzSeiten := lv.Items.Count div Zeilen + 1
   else
     AnzSeiten := lv.Items.Count div Zeilen;

   Seite := 1;
   //Grid Drucken
  for P := 1 to AnzSeiten do
   begin
     //Kopfzeile
    Printer.Canvas.Font.Height := 48;
     Printer.Canvas.TextOut((HorzSize div 2 - (Printer.Canvas.TextWidth(Title) div 2)),
       - 20,Title);
     Printer.Canvas.Pen.Width := 5;
     Printer.Canvas.MoveTo(0, - HeaderSize);
     Printer.Canvas.LineTo(HorzSize, - HeaderSize);
     //Fu?zeile
    Printer.Canvas.MoveTo(0, - VertSize + FooterSize);
     Printer.Canvas.LineTo(HorzSize, - VertSize + FooterSize);
     Printer.Canvas.Font.Height := 36;
     Footer := 'Страница: ' + IntToStr(Seite) + ' из ' + IntToStr(AnzSeiten);
     Printer.Canvas.TextOut((HorzSize div 2 - (Printer.Canvas.TextWidth(Footer) div 2)),
       - VertSize + 150,Footer);
     //Zeilen drucken
    Printer.Canvas.Font.Height := FontHeight;
     YPos := HeaderSize + 10;
     for I := 1 to Zeilen do
     begin
       if lv.Items.Count >= I + (Seite - 1) * Zeilen then
       begin
         XPos := 100;
           Printer.Canvas.TextOut(XPos, - YPos,
             lv.Items[I + (Seite - 1) * Zeilen - 1].Caption);
           XPos := XPos + 1400;
           if lv.Items[I + (Seite - 1) * Zeilen - 1].Subitems.Count>0 then
            Printer.Canvas.TextOut(XPos, - YPos,
             lv.Items[I + (Seite - 1) * Zeilen - 1].Subitems[0]);
         YPos := YPos + ZeilenSize;
       end;
     end;
     //Seite hinzufugen
    Inc(Seite);
     if Seite <= AnzSeiten then Printer.NewPage;
   end;
   Printer.EndDoc;
 end;

function ListViewConfHTML(Listview:TListview; output:string; center: Boolean) : Boolean;
 var
   i,f: Integer;
   tfile: TextFile;
 begin
   try
     ForceDirectories(ExtractFilePath(output));
     AssignFile(tfile,output);
     ReWrite(tfile);
     WriteLn(tfile,'<html>');
     WriteLn(tfile,'<head>');
     WriteLn(tfile,'<title>HTML-Ansicht: '+listview.Name+'</title>');
     WriteLn(tfile,'</head>');
     WriteLn(tfile,'<table border="1" bordercolor="#000000">');
     WriteLn(tfile,'<tr>');
     for i := 0 to listview.Columns.Count - 1 do
     begin
       if center then
         WriteLn(tfile,'<td><b><center>'+listview.columns[i].caption+'</center></b></td>') else
         WriteLn(tfile,'<td><b>'+listview.columns[i].caption+'</b></td>');
     end;
     WriteLn(tfile,'</tr>');
     WriteLn(tfile,'<tr>');
     for i := 0 to listview.Items.Count-1 do
     begin
       WriteLn(tfile,'<td>'+listview.items.item[i].caption+'</td>');
       if listview.items.item[i].Subitems.Count > 0 then
        for f := 0 to listview.Columns.Count-2 do
        begin
         if listview.items.item[i].subitems[f]='' then Write(tfile,'<td>-</td>') else
           Write(tfile,'<td>'+listview.items.item[i].subitems[f]+'</td>');
        end;
       Write(tfile,'</tr>');
     end;
     WriteLn(tfile,'</table>');
     WriteLn(tfile,'</html>');
     CloseFile(tfile);
     Result := True;
   except
   Result := False;
   end;
 end;

function StringGridConfHTML(StringGrid: TStringGrid; output:string; center: Boolean) : Boolean;
 var
   i,f: Integer;
   tfile: TextFile;
 begin
  with StringGrid do
   try
     ForceDirectories(ExtractFilePath(output));
     AssignFile(tfile,output);
     ReWrite(tfile);
     WriteLn(tfile,'<html>');
     WriteLn(tfile,'<head>');
     WriteLn(tfile,'</head>');
     WriteLn(tfile,'<table border="1" bordercolor="#000000">');
     WriteLn(tfile,'<tr>');
     for i := 0 to ColCount - 1 do
     begin
       if center then
         WriteLn(tfile,'<td><b><center>'+Cells[i, 0]+'</center></b></td>') else
         WriteLn(tfile,'<td><b>'+Cells[i, 0]+'</b></td>');
     end;

     WriteLn(tfile,'</tr>');
     WriteLn(tfile,'<tr>');

     for i := 1 to RowCount-1 do
     begin
        for f := 0 to ColCount-1 do
        begin
         if Cells[f,i]='' then Write(tfile,'<td>-</td>') else
           Write(tfile,'<td>'+Cells[f,i]+'</td>');
        end;
       Write(tfile,'</tr>');
     end;
     WriteLn(tfile,'</table>');
     WriteLn(tfile,'</html>');
     CloseFile(tfile);
     Result := True;
   except
   Result := False;
   end;
 end;

procedure SaveLV(lv: TListView; const FileName: TFileName);
 var
   f:    TextFile;
   i, k: Integer;

 begin
   AssignFile(f, FileName);
   Rewrite(f);
   with lv do
   begin
    for i := 0 to Items.Count - 1 do
    begin
     Write(f,Items[i].Caption + ' - ');
     if Items[i].Subitems.Count>0 then
      Writeln(f, Items[i].Subitems[0])
     else
      Writeln(f, '');
    end;
   end;
   CloseFile(F);
 end;

// Save a TStringGrid to a file

procedure SaveStringGrid(StringGrid: TStringGrid; const FileName: TFileName);
 var
   f:    TextFile;
   i, k: Integer;
 begin
   AssignFile(f, FileName);
   Rewrite(f);
   with StringGrid do
   begin
     // Write number of Columns/Rows
    Writeln(f, ColCount);
     Writeln(f, RowCount);
     // loop through cells
    for i := 0 to ColCount - 1 do
       for k := 0 to RowCount - 1 do
         Writeln(F, Cells[i, k]);
   end;
   CloseFile(F);
 end;

 // Load a TStringGrid from a file

procedure LoadStringGrid(StringGrid: TStringGrid; const FileName: TFileName);
 var
   f:          TextFile;
   iTmp, i, k: Integer;
   strTemp:    String;
 begin
   AssignFile(f, FileName);
   Reset(f);
   with StringGrid do
   begin
     // Get number of columns
    Readln(f, iTmp);
     ColCount := iTmp;
     // Get number of rows
    Readln(f, iTmp);
     RowCount := iTmp;
     // loop through cells & fill in values
    for i := 0 to ColCount - 1 do
       for k := 0 to RowCount - 1 do
       begin
         Readln(f, strTemp);
         Cells[i, k] := strTemp;
       end;
   end;
   CloseFile(f);
 end;

function IsConnectedToInternet: Boolean;
var
   dwConnectionTypes: DWORD;
begin
   dwConnectionTypes :=
     INTERNET_CONNECTION_MODEM +
     INTERNET_CONNECTION_LAN +
     INTERNET_CONNECTION_PROXY;
   Result := InternetGetConnectedState(@dwConnectionTypes, 0);
end;

function DownloadURL_NOCache(const aUrl: string; var F: TFileStream): Boolean;
 var
   hSession: HINTERNET;
   hService: HINTERNET;
   lpBuffer: array[0..1024 + 1] of Char;
   dwBytesRead: DWORD;
begin
  Result := False;
  try
   Screen.Cursor := crHourGlass;
   hSession := InternetOpen('MyApp', INTERNET_OPEN_TYPE_PRECONFIG, nil, nil, 0);
   try
     if Assigned(hSession) then
     begin
       hService := InternetOpenUrl(hSession, PChar(aUrl), nil, 0, INTERNET_FLAG_RELOAD, 0);
       if Assigned(hService) then
         try
           while True do
           begin
             dwBytesRead := 1024;
             InternetReadFile(hService, @lpBuffer, 1024, dwBytesRead);
             if dwBytesRead = 0 then break;
             lpBuffer[dwBytesRead] := #0;
             F.Write(lpBuffer,dwBytesRead);
           end;
           Result := True;
         finally
           InternetCloseHandle(hService);
         end;
     end;
   finally
     InternetCloseHandle(hSession);
     Screen.Cursor := crDefault;
   end;
  except
   Screen.Cursor := crDefault;
  end;
end;

function DownloadURL_NOCache_Mem(const aUrl: string; var F: TMemoryStream): Boolean;
 var
   hSession: HINTERNET;
   hService: HINTERNET;
   lpBuffer: array[0..1024 + 1] of Char;
   dwBytesRead: DWORD;
begin
  Result := False;
  try
   Screen.Cursor := crHourGlass;
   hSession := InternetOpen('MyApp', INTERNET_OPEN_TYPE_PRECONFIG, nil, nil, 0);
   try
     if Assigned(hSession) then
     begin
       hService := InternetOpenUrl(hSession, PChar(aUrl), nil, 0, INTERNET_FLAG_RELOAD, 0);
       if Assigned(hService) then
         try
           while True do
           begin
             dwBytesRead := 1024;
             InternetReadFile(hService, @lpBuffer, 1024, dwBytesRead);
             if dwBytesRead = 0 then break;
             lpBuffer[dwBytesRead] := #0;
             F.Write(lpBuffer,dwBytesRead);
           end;
           Result := True;
         finally
           InternetCloseHandle(hService);
         end;
     end;
   finally
     InternetCloseHandle(hSession);
     Screen.Cursor := crDefault;
   end;
  except
   Screen.Cursor := crDefault;
  end;
end;

function TMainForm.MonthNumber(m : string):integer;
var
 mi : integer;
begin
 if m='январь' then mi:=1;
 if m='февраль' then mi:=2;
 if m='март' then mi:=3;
 if m='апрель' then mi:=4;
 if m='май' then mi:=5;
 if m='июнь' then mi:=6;
 if m='июль' then mi:=7;
 if m='август' then mi:=8;
 if m='сентябрь' then mi:=9;
 if m='октябрь' then mi:=10;
 if m='ноябрь' then mi:=11;
 if m='декабрь' then mi:=12;
 Result := mi;
end;

function TMainForm.GetMonth(d : TDateTime) : string;
var
 year, month, day : word;
 smonth : string;
begin
 DecodeDate(d,year,month,day);
 case month of
  1: smonth := 'январь';
  2: smonth := 'февраль';
  3: smonth := 'март';
  4: smonth := 'апрель';
  5: smonth := 'май';
  6: smonth := 'июнь';
  7: smonth := 'июль';
  8: smonth := 'август';
  9: smonth := 'сентябрь';
  10: smonth := 'октябрь';
  11: smonth := 'ноябрь';
  12: smonth := 'декабрь';
 end;
 Result := smonth;
end;

procedure TMainForm.RegistryS30(path : string);
var
 Reg : TRegistry;
begin
  Reg:=TRegistry.Create;
  try
   Reg.RootKey:=HKEY_CLASSES_ROOT;
   if not Reg.KeyExists('\.s40') then
   begin
    if Reg.OpenKey('\.s40',True) then
    begin
      Reg.WriteString('','Sport40DataBase.File');
      Reg.CloseKey;
    end;
    if Reg.OpenKey('\Sport40DataBase.File',True) then
    begin
      Reg.WriteString('','Спорт 4.0 база данных');
      Reg.CloseKey;
    end;
    if Reg.OpenKey('\Sport40DataBase.File\Shell\Open\command',True) then
    begin
      Reg.WriteString('','"'+path+'sport4.exe" "%1"');
      Reg.CloseKey;
    end;
    if Reg.OpenKey('\Sport40DataBase.File\DefaultIcon',True) then
    begin
      Reg.WriteString('','%SystemRoot%\system32\SHELL32.dll,213');
      Reg.CloseKey;
    end;
   end;
   Reg.Free;
  except
  end;
end;

procedure TMainForm.ReadExecsToString(var st : TStringList);
var
 F : TFileStream;
 len,j,i,cnt : integer;
 buffer : PAnsiChar;
 s : string;
begin
 if FileExists(ExtractFilePath(paramstr(0)) + 'exec.mtx') then
 begin
  try
   F := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'exec.mtx',fmOpenRead);
   F.Read(len,4);
   if len > ExecVersion then
    MessageBOX(Handle, 'Справочник "'+ExecTitle+'" не загружен...', MainTitle ,
    MB_ICONWARNING)
   else
   begin
    while F.Position < F.Size do
    begin
     F.Read(len,4);
     buffer := AllocMem(len+1);
     buffer[len] := Chr(0);
     F.Read(buffer^,len);
     buffer[len]:=Chr(0);
     s := StrPas(buffer);
     FreeMem(buffer);
     st.Add(s);

     F.Read(cnt,4);
     if cnt>0 then
     begin
       for j := 0 to cnt - 1 do
       begin
        F.Read(len,4);
        F.Seek(len,soFromCurrent);
       end;
     end;
    end;
    end;
   F.Free;
  except
   MessageBOX(Handle, 'Ошибка при чтении справочника "'+ExecTitle+'"...', MainTitle ,
   MB_ICONERROR);
  end;
 end;
end;

procedure TMainForm.ReadExecs2ToString(index : integer; var st : TStringList);
var
 F : TFileStream;
 jj, len,j,i,cnt : integer;
 buffer : PAnsiChar;
 s : string;
begin
 if FileExists(ExtractFilePath(paramstr(0)) + 'exec.mtx') then
 begin
  try
   jj := 0;
   F := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'exec.mtx',fmOpenRead);
   F.Read(len,4);
   if len > ExecVersion then
    MessageBOX(Handle, 'Справочник "'+ExecTitle+'" не загружен...', MainTitle ,
    MB_ICONWARNING)
   else
   begin
    while F.Position < F.Size do
    begin
     F.Read(len,4);
     F.Seek(len,soFromCurrent);

     F.Read(cnt,4);
     if cnt>0 then
     begin
       for j := 0 to cnt - 1 do
       begin
        F.Read(len,4);
        buffer := AllocMem(len+1);
        buffer[len] := Chr(0);
        F.Read(buffer^,len);
        buffer[len]:=Chr(0);
        s := StrPas(buffer);
        FreeMem(buffer);
        if Index = jj then
         st.Add(s);
       end;
     end;
     inc(jj);
    end;
    end;
   F.Free;
  except
   MessageBOX(Handle, 'Ошибка при чтении справочника "'+ExecTitle+'"...', MainTitle ,
   MB_ICONERROR);
  end;
 end;
end;

procedure TMainForm.C1Click(Sender: TObject);
begin
 ManForm.fname := TMDIChild(ActiveMDIChild).fname;
 if ManForm.ShowModal = mrOk then
  UpdateDataBase;
end;

procedure TMainForm.CreateMDIChild(Name : string; filename: string);
var
  Child: TMDIChild;
begin
  { create a new MDI child window }
  Child := TMDIChild.Create(Application);
  DockTabSet1.Visible := true;
  DockTabSet1.Tabs.Add(Name + '       ');
  DockTabSet1.TabIndex := DockTabSet1.Tabs.Count-1;
  Child.Caption := Name;
  Child.fname := fileName;
  Child.OpenFile
end;

procedure TMainForm.CreatePreviewChild(Name : string);
var
  Child: TFormPreview;
begin
  { create a new MDI child window }
  Child := TFormPreview.Create(Application);
  Child.Caption := Child.Caption  + ' - ' + Name;
  DockTabSet1.Visible := true;
  DockTabSet1.Tabs.Add(Child.Caption);
  DockTabSet1.TabIndex := DockTabSet1.Tabs.Count-1;
end;

procedure TMainForm.DockTabSet1Change(Sender: TObject; NewTab: Integer;
  var AllowChange: Boolean);
var
 i : integer;
 fnd : boolean;
begin
 AllowChange := true;
  fnd := false;
  for i := 0 to MDIChildCount - 1 do
   begin
     if MDIChildren[i].Caption = Trim(DockTabSet1.Tabs[NewTab]) then
       begin
        fnd := true;
        Break;
       end;
   end;
 if fnd then
 begin
   MDIChildren[i].Show;
  if Copy(ActiveMDIChild.Caption,1,5) = 'Отчет' then
   with THTMLForm(ActiveMDIChild).wb do
    if Document <> nil then
     with Application as IOleobject do
       DoVerb(OLEIVERB_UIACTIVATE, nil, THTMLForm(ActiveMDIChild).wb, 0, Handle, GetClientRect);
 end;
end;

procedure TMainForm.DockTabSet1DrawTab(Sender: TObject; TabCanvas: TCanvas;
  R: TRect; Index: Integer; Selected: Boolean);
begin
with TabCanvas do
  begin
      FillRect(R);
      Font.Style := [];
      if Trim(DockTabSet1.Tabs[Index]) = 'Ввод данных' then
       ImageList1.Draw(TabCanvas,R.Left+1,R.Top+2,22)
      else
      if Copy(Trim(DockTabSet1.Tabs[Index]),1,9) = 'Диаграмма' then
       ImageList1.Draw(TabCanvas,R.Left+1,R.Top+2,26)
      else
      if Copy(Trim(DockTabSet1.Tabs[Index]),1,5) = 'Отчет' then
       ImageList1.Draw(TabCanvas,R.Left+1,R.Top+2,26)
      else
      if Trim(DockTabSet1.Tabs[Index]) = ExecTitle then
       ImageList1.Draw(TabCanvas,R.Left+1,R.Top+2,30)
      else
       ImageList1.Draw(TabCanvas,R.Left+1,R.Top+2,33);
      TextOut(R.left + 20, R.top + 2, Trim(DockTabSet1.Tabs[Index]));
  end;
end;

procedure TMainForm.FileOpen1Execute(Sender: TObject);
var
 i,j,len : integer;
 fnd : boolean;
 modalres : integer;
 Fname, FIO : string;
 LI : TListItem;
 found : boolean;
begin
  if not IfRegOK then exit;

  if AdminMode then
   begin
   modalres := BaseForm.ShowModal;
   if (modalres = mrYes) and (IsConnectedToInternet) then
   begin
   try
    Screen.Cursor := crHourGlass;
    RegComplete;
    ServerForm.ListView1.Items.Clear;
    RegisterFTPListParser(TIdFTPLPUnix);
    IdFTP2.Host := 'startedu.ru';
    IdFTP2.Username := 'u222';
    IdFTP2.Password := '747';
    ImportMode := false;
    IdFTP2.Connect;
    if IdFTP2.Connected then
    begin
     IdFTP2.ChangeDir('/');
     IdFTP2.List;
     for j:=0 to Pred(IdFTP2.DirectoryListing.Count) do
     begin
      found := false;
      for i:= 0 to UsersList.Count - 1 do
      begin
       FName := copy(UsersList[i],2,pos(']',UsersList[i])-2)+'.s40';
       FIO := copy(UsersList[i],pos(']',UsersList[i])+2,length(UsersList[i])-pos(']',UsersList[i])-2);
       if FName=IdFTP2.DirectoryListing.Items[j].FileName then
       begin
        found := true;
        break;
       end;
      end;
      if found then
      begin
       LI := ServerForm.ListView1.Items.Add;
       LI.Caption := FIO;
       LI.SubItems.Add(Fname);
       LI.SubItems.Add(DateToStr(IdFTP2.DirectoryListing.Items[j].ModifiedDate));
       LI.SubItems.Add(IntToStr(IdFTP2.DirectoryListing.Items[j].Size));
      end;
     end;
     IdFTP2.Disconnect;
    end;
    Screen.Cursor := crDefault;
    except
     Screen.Cursor := crDefault;
    end;
    if ServerForm.ShowModal=mrOk then
     begin
      for j := 0 to ServerForm.ListView1.Items.Count - 1 do
      begin
       if ServerForm.ListView1.Items[j].Checked then
       begin
        // Закачаем файл
        ProgressForm.pb.Max := 0;
        ProgressForm.Caption := 'Импорт информации...';
        ProgressForm.pb.Position := 0;
        ProgressForm.Show;
        Application.ProcessMessages;
        IdFTP3.Host := 'startedu.ru';
        IdFTP3.Username := 'u222';
        IdFTP3.Password := '747';
        IdFTP3.Connect;
        if IdFTP3.Connected then
        begin
         IdFTP3.ChangeDir('/');
         len := IdFTP3.Size(ServerForm.ListView1.Items[j].Subitems[0]);
         if len>0 then
         begin
          ProgressForm.pb.Max := len;
          IdFTP3.TransferType := ftBinary;
          ImportMode := true;
          ImportFileMem := TMemoryStream.Create;
          ImportFileFIO := ServerForm.ListView1.Items[j].Caption;
          ImportFileName := ServerForm.ListView1.Items[j].Subitems[0];
          IdFTP3.Get(ServerForm.ListView1.Items[j].Subitems[0],ImportFileMem);
         end
         else
         begin
          ProgressForm.Close;
          MessageBOX(Handle, 'База данных не найдена на сервере.', MainTitle , MB_ICONWARNING);
          IdFTP3.Disconnect;
         end;
        end
        else
         ProgressForm.Close;
        break;
       end;
      end;
     end;
   end
   else
   if modalres = mrNo then
   begin
   if OpenDialog.Execute then
   begin
   fnd := false;
   for i := 0 to MDIChildCount - 1 do
   begin
     if  MDIChildren[i].Caption = OpenDialog.FileName then
       begin
        fnd := true;
        Break;
       end;
   end;
   if not fnd then
    CreateMDIChild(OpenDialog.FileName, OpenDialog.FileName)
   else
   begin
    MessageBOX(Handle, PChar('База данных с именем '+OpenDialog.FileName+' уже открыта.'), PChar(MainTitle) ,
    MB_ICONInformation);
    MDIChildren[i].Show;
   end;
   end;
   end;
  end
  else
  if OpenDialog.Execute then
  begin
   fnd := false;
   for i := 0 to MDIChildCount - 1 do
   begin
     if  MDIChildren[i].Caption = OpenDialog.FileName then
       begin
        fnd := true;
        Break;
       end;
   end;
   if not fnd then
    CreateMDIChild(OpenDialog.FileName, OpenDialog.FileName)
   else
   begin
    MessageBOX(Handle, PChar('База данных с именем '+OpenDialog.FileName+' уже открыта.'), PChar(MainTitle) ,
    MB_ICONInformation);
    MDIChildren[i].Show;
   end;
  end;
end;

procedure TMainForm.Timer1Timer(Sender: TObject);
begin
 Timer1.Enabled := false;
 N12Click(Sender);
end;

procedure TMainForm.Timer2Timer(Sender: TObject);
var
 i : integer;
 fnd : boolean;
begin
        Timer2.Enabled := false;
        fnd := false;
        for i := 0 to MDIChildCount - 1 do
         begin
         if  MDIChildren[i].Caption = ImportfileFIO then
          begin
           fnd := true;
           Break;
          end;
         end;
        if not fnd then
         CreateMDIChild(ImportFileFIO,ExtractFilePath(paramstr(0)) + ImportfileName)
        else
        begin
         MessageBOX(Handle, PChar('База данных с именем "'+ImportFileFIO+'" уже открыта.'), PChar(MainTitle) ,
         MB_ICONInformation);
         MDIChildren[i].Show;
        end;

end;

procedure TMainForm.ToolButton11Click(Sender: TObject);
var
 len, success: Integer;
 s : string;
 F : TFileStream;
begin
 if not IfRegOK then exit;

 if IsConnectedToInternet then
 begin
  RegComplete;
  if not IfRegOK then exit;

  try
   ImportFile := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'local.mtx',fmOpenRead);
   len := ImportFile.Size;
  except
  end;

  IdFTP1.Host := 'startedu.ru';
  IdFTP1.Username := 'u222';
  IdFTP1.Password := '747';

  ProgressForm.pb.Max := len;
  ProgressForm.Caption := 'Экспорт информации...';
  ProgressForm.pb.Position := 0;
  ProgressForm.Show;
  Application.ProcessMessages;

  IdFTP1.Connect;
  if IdFTP1.Connected then
  begin
     IdFTP1.ChangeDir('/');
     IdFTP1.TransferType := ftBinary;
     IdFTP1.Put(ImportFile, RegData+'.s40');
  end
  else
   ProgressForm.Close;

 end
 else
  MessageBOX(Handle, 'Отсутствует активное подключение к сети Интернет.', MainTitle , MB_ICONWARNING);

 { DiagForm.fname :=  TMDIChild(ActiveMDIChild).fname;
 DiagForm.MasterType := 1; // Экспорт
 if DiagForm.ShowModal = mrOk then
   ExportData;}
end;

procedure TMainForm.ToolButton12Click(Sender: TObject);
var
 len, success: Integer;
 s : string;
begin
 if not IfRegOK then exit;

 if IsConnectedToInternet then
 begin
 if MessageBOX(Handle, PChar('Внимание! Будет произведен импорт Вашей локальной базы данных. Продолжить?'), PChar(MainTitle) ,
      MB_YESNO or MB_ICONQUESTION) = IDYes then
 begin
  RegComplete;
  if not IfRegOK then exit;

  IdFTP1.Host := 'startedu.ru';
  IdFTP1.Username := 'u222';
  IdFTP1.Password := '747';

  ProgressForm.pb.Max := 0;
  ProgressForm.Caption := 'Импорт информации...';
  ProgressForm.pb.Position := 0;
  ProgressForm.Show;
  Application.ProcessMessages;

  IdFTP1.Connect;
  if IdFTP1.Connected then
  begin
     IdFTP1.ChangeDir('/');
     len := IdFTP1.Size(RegData+'.s40');
     if len>0 then
     begin
      ProgressForm.pb.Max := len;
      IdFTP1.TransferType := ftBinary;
      ImportMode := true;
      ImportFileMem := TMemoryStream.Create;
      IdFTP1.Get(RegData+'.s40',ImportFileMem);
     end
     else
     begin
      ProgressForm.Close;
      MessageBOX(Handle, 'База данных не найдена на сервере.', MainTitle , MB_ICONWARNING);
      IdFTP1.Disconnect;
     end;
  end
  else
   ProgressForm.Close;
 end;
 end
 else
  MessageBOX(Handle, 'Отсутствует активное подключение к сети Интернет.', MainTitle , MB_ICONWARNING);

 { DiagForm.MasterType := 2; // Импорт
 if DiagForm.ShowModal = mrOk then
   ImportData;}
end;

procedure TMainForm.ExportExcel;
var
   XL: variant;
   dl : TDataLeaf;
   fnd, manfound : boolean;
   k, k2, k3, cursheet, i,j, jj : integer;
   s : string;
   ST : TStringList;
begin
  if DiagForm.ListView2.Items.Count > 1 then
   MessageBOX(Handle, PChar('Экспорт информации в Microsoft Excel будет произведен только по первому спрортсмену.'), Pchar(MainTitle) ,
   MB_ICONINFORMATION);

  try
    XL := GetActiveOleObject('Excel.Application');
  except
    try
      XL := CreateOleObject('Excel.Application');
    except
      Exception.Create('Невозможно запустить приложение Microsoft Excel');
      Exit;
    end;
  end;

  ST := TStringList.Create;

  try
  Application.ProcessMessages;

  RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'control.mtx';
  RazrDlg.GetRazr(ST);

  XL.DisplayAlerts := false;
  XL.WorkBooks.Add;
  Screen.Cursor := crHourGlass;

  XL.Visible := false;

  ProgressForm.Caption := 'Экспорт информации...';
  ProgressForm.pb.Position := 0;
  ProgressForm.Show;

//  XL.WorkBooks[1].WorkSheets[1].Delete;

  cursheet := 1;
  k := 1;
  k2 := 1;
  k3 := 1;

  XL.WorkBooks[1].WorkSheets[1].Name := 'Нагрузка';
  XL.WorkBooks[1].WorkSheets[2].Name := 'Восстановление';
  XL.WorkBooks[1].WorkSheets[3].Name := 'Контроль';

  XL.WorkBooks[1].WorkSheets[1].Columns[1].ColumnWidth := 20;
  XL.WorkBooks[1].WorkSheets[1].Columns[2].ColumnWidth := 30;
  XL.WorkBooks[1].WorkSheets[2].Columns[1].ColumnWidth := 20;
  XL.WorkBooks[1].WorkSheets[2].Columns[2].ColumnWidth := 30;

  XL.WorkBooks[1].WorkSheets[3].Columns[1].ColumnWidth := 20;
  XL.WorkBooks[1].WorkSheets[3].Columns[2].ColumnWidth := 30;

  s := DiagForm.ListView2.Items[j].Subitems[0];

  XL.WorkBooks[1].WorkSheets[1].Cells[1, 1] := s;
  XL.WorkBooks[1].WorkSheets[1].Range['A1:B1'].Merge;
  XL.WorkBooks[1].WorkSheets[1].Range['A1:B1'].Borders.LineStyle := 1;
  XL.WorkBooks[1].WorkSheets[1].Range['A1:B1'].Borders.Weight := 2;
  XL.WorkBooks[1].WorkSheets[1].Range['A1:B1'].Borders.ColorIndex := 1;
  XL.WorkBooks[1].WorkSheets[1].Cells[1, 1].HorizontalAlignment := 3;

  XL.WorkBooks[1].WorkSheets[2].Cells[1, 1] := s;
  XL.WorkBooks[1].WorkSheets[2].Range['A1:B1'].Merge;
  XL.WorkBooks[1].WorkSheets[2].Range['A1:B1'].Borders.LineStyle := 1;
  XL.WorkBooks[1].WorkSheets[2].Range['A1:B1'].Borders.Weight := 2;
  XL.WorkBooks[1].WorkSheets[2].Range['A1:B1'].Borders.ColorIndex := 1;
  XL.WorkBooks[1].WorkSheets[2].Cells[1, 1].HorizontalAlignment := 3;
  XL.WorkBooks[1].WorkSheets[3].Cells[1, 1] := s;
  XL.WorkBooks[1].WorkSheets[3].Range['A1:B1'].Merge;
  XL.WorkBooks[1].WorkSheets[3].Range['A1:B1'].Borders.LineStyle := 1;
  XL.WorkBooks[1].WorkSheets[3].Range['A1:B1'].Borders.Weight := 2;
  XL.WorkBooks[1].WorkSheets[3].Range['A1:B1'].Borders.ColorIndex := 1;
  XL.WorkBooks[1].WorkSheets[3].Cells[1, 1].HorizontalAlignment := 3;

      for I := 0 to TMDIChild(ActiveMDIChild).BigDataLeaf.Count - 1 do
      if TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i]).menindex<>-1 then
       begin
        ProgressForm.pb.Position := Trunc(i/TMDIChild(ActiveMDIChild).BigDataLeaf.Count*100);
        // ProgressForm.pb.Repaint;
        Application.ProcessMessages;

        dl := TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i]);

        manfound := false;
        for j := 0 to DiagForm.ListView2.Items.Count - 1 do
        if dl.men = DiagForm.ListView2.Items[j].Subitems[0] then
         begin
           manfound := true;
           break;
         end;

        fnd := false;
        if manfound then
         if DiagForm.RadioButton2.Checked or ((dl.mezo >= StrToInt(DiagForm.MezoB.Text)) and
         (dl.mezo <= StrToInt(DiagForm.MezoE.Text)) and (dl.micro >= StrToInt(DiagForm.MicroB.Text)) and
         (dl.micro <= StrToInt(DiagForm.MicroE.Text)) )then
          fnd := true;

        if fnd then
         begin

           if dl.RelaxList.Count>0 then
           begin

           inc(k2);
           inc(k2);
           XL.WorkBooks[1].WorkSheets[2].Cells[k2, 1] := DateToStr(dl.dt);
           XL.WorkBooks[1].WorkSheets[2].Range['A'+IntToStr(k2)+':B'+IntToStr(k2)].Merge;
           XL.WorkBooks[1].WorkSheets[2].Range['A'+IntToStr(k2)+':B'+IntToStr(k2)].Borders.LineStyle := 1;
           XL.WorkBooks[1].WorkSheets[2].Range['A'+IntToStr(k2)+':B'+IntToStr(k2)].Borders.Weight := 2;
           XL.WorkBooks[1].WorkSheets[2].Range['A'+IntToStr(k2)+':B'+IntToStr(k2)].Borders.ColorIndex := 1;
           XL.WorkBooks[1].WorkSheets[2].Cells[k, 1].HorizontalAlignment := 3;

           for j := 0 to dl.RelaxList.Count - 1 do
           begin
            XL.WorkBooks[1].WorkSheets[2].Cells[k2+(j+1)*2, 1].VerticalAlignment := 2;
            XL.WorkBooks[1].WorkSheets[2].Cells[k2+(j+1)*2, 1] := IntToStr(j+1);
            XL.WorkBooks[1].WorkSheets[2].Range['A'+IntToStr(k2+(j+1)*2)+':A'+IntToStr(1+k2+(j+1)*2)].Merge;
            XL.WorkBooks[1].WorkSheets[2].Range['A'+IntToStr(k2+(j+1)*2)+':A'+IntToStr(1+k2+(j+1)*2)].Borders.LineStyle := 1;
            XL.WorkBooks[1].WorkSheets[2].Range['A'+IntToStr(k2+(j+1)*2)+':A'+IntToStr(1+k2+(j+1)*2)].Borders.Weight := 2;
            XL.WorkBooks[1].WorkSheets[2].Range['A'+IntToStr(k2+(j+1)*2)+':A'+IntToStr(1+k2+(j+1)*2)].Borders.ColorIndex := 1;

            XL.WorkBooks[1].WorkSheets[2].Cells[k2+(j+1)*2, 2] := dl.RelaxList[j];
            XL.WorkBooks[1].WorkSheets[2].Cells[k2+(j+1)*2, 2].VerticalAlignment := 2;
            XL.WorkBooks[1].WorkSheets[2].Range['B'+IntToStr(k2+(j+1)*2)+':B'+IntToStr(1+k2+(j+1)*2)].Merge;
            XL.WorkBooks[1].WorkSheets[2].Range['B'+IntToStr(k2+(j+1)*2)+':B'+IntToStr(1+k2+(j+1)*2)].Borders.LineStyle := 1;
            XL.WorkBooks[1].WorkSheets[2].Range['B'+IntToStr(k2+(j+1)*2)+':B'+IntToStr(1+k2+(j+1)*2)].Borders.Weight := 2;
            XL.WorkBooks[1].WorkSheets[2].Range['B'+IntToStr(k2+(j+1)*2)+':B'+IntToStr(1+k2+(j+1)*2)].Borders.ColorIndex := 1;
           end;

           k2 := 2 + k2 + dl.RelaxList.Count*2;
           end;

           if dl.ControlList.Count>0 then
           begin
           inc(k3);
           inc(k3);
           XL.WorkBooks[1].WorkSheets[3].Cells[k3, 1] := DateToStr(dl.dt);
           XL.WorkBooks[1].WorkSheets[3].Range['A'+IntToStr(k3)+':B'+IntToStr(k3)].Merge;
           XL.WorkBooks[1].WorkSheets[3].Range['A'+IntToStr(k3)+':B'+IntToStr(k3)].Borders.LineStyle := 1;
           XL.WorkBooks[1].WorkSheets[3].Range['A'+IntToStr(k3)+':B'+IntToStr(k3)].Borders.Weight := 2;
           XL.WorkBooks[1].WorkSheets[3].Range['A'+IntToStr(k3)+':B'+IntToStr(k3)].Borders.ColorIndex := 1;
           XL.WorkBooks[1].WorkSheets[3].Cells[k3, 1].HorizontalAlignment := 3;

           for j := 0 to dl.ControlList.Count - 1 do
           begin
            XL.WorkBooks[1].WorkSheets[3].Cells[k3+(j+1)*2, 1].VerticalAlignment := 2;
            XL.WorkBooks[1].WorkSheets[3].Cells[k3+(j+1)*2, 1] := IntToStr(j+1);
            XL.WorkBooks[1].WorkSheets[3].Range['A'+IntToStr(k3+(j+1)*2)+':A'+IntToStr(1+k3+(j+1)*2)].Merge;
            XL.WorkBooks[1].WorkSheets[3].Range['A'+IntToStr(k3+(j+1)*2)+':A'+IntToStr(1+k3+(j+1)*2)].Borders.LineStyle := 1;
            XL.WorkBooks[1].WorkSheets[3].Range['A'+IntToStr(k3+(j+1)*2)+':A'+IntToStr(1+k3+(j+1)*2)].Borders.Weight := 2;
            XL.WorkBooks[1].WorkSheets[3].Range['A'+IntToStr(k3+(j+1)*2)+':A'+IntToStr(1+k3+(j+1)*2)].Borders.ColorIndex := 1;

            XL.WorkBooks[1].WorkSheets[3].Cells[k3+(j+1)*2, 2] := ST[j] + ' = ' + dl.ControlList[j];
            XL.WorkBooks[1].WorkSheets[3].Cells[k3+(j+1)*2, 2].VerticalAlignment := 2;
            XL.WorkBooks[1].WorkSheets[3].Range['B'+IntToStr(k3+(j+1)*2)+':B'+IntToStr(1+k3+(j+1)*2)].Merge;
            XL.WorkBooks[1].WorkSheets[3].Range['B'+IntToStr(k3+(j+1)*2)+':B'+IntToStr(1+k3+(j+1)*2)].Borders.LineStyle := 1;
            XL.WorkBooks[1].WorkSheets[3].Range['B'+IntToStr(k3+(j+1)*2)+':B'+IntToStr(1+k3+(j+1)*2)].Borders.Weight := 2;
            XL.WorkBooks[1].WorkSheets[3].Range['B'+IntToStr(k3+(j+1)*2)+':B'+IntToStr(1+k3+(j+1)*2)].Borders.ColorIndex := 1;
           end;
           k3 := 2 + k3 + dl.ControlList.Count*2;
           end;

           if dl.BigDataGrid.Count>0 then
           begin

           inc(k);
           inc(k);

           XL.WorkBooks[1].WorkSheets[1].Cells[k, 1] := DateToStr(dl.dt);
           XL.WorkBooks[1].WorkSheets[1].Range['A'+IntToStr(k)+':B'+IntToStr(k)].Merge;
           XL.WorkBooks[1].WorkSheets[1].Range['A'+IntToStr(k)+':B'+IntToStr(k)].Borders.LineStyle := 1;
           XL.WorkBooks[1].WorkSheets[1].Range['A'+IntToStr(k)+':B'+IntToStr(k)].Borders.Weight := 2;
           XL.WorkBooks[1].WorkSheets[1].Range['A'+IntToStr(k)+':B'+IntToStr(k)].Borders.ColorIndex := 1;
           XL.WorkBooks[1].WorkSheets[1].Cells[k, 1].HorizontalAlignment := 3;

           for j := 0 to dl.BigDataGrid.Count - 1 do
           begin

            XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, 1].VerticalAlignment := 2;
            XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, 1] := IntToStr(j+1);
            XL.WorkBooks[1].WorkSheets[1].Range['A'+IntToStr(k+(j+1)*2)+':A'+IntToStr(1+k+(j+1)*2)].Merge;
            XL.WorkBooks[1].WorkSheets[1].Range['A'+IntToStr(k+(j+1)*2)+':A'+IntToStr(1+k+(j+1)*2)].Borders.LineStyle := 1;
            XL.WorkBooks[1].WorkSheets[1].Range['A'+IntToStr(k+(j+1)*2)+':A'+IntToStr(1+k+(j+1)*2)].Borders.Weight := 2;
            XL.WorkBooks[1].WorkSheets[1].Range['A'+IntToStr(k+(j+1)*2)+':A'+IntToStr(1+k+(j+1)*2)].Borders.ColorIndex := 1;

            if TDataContent(dl.BigDataGrid.Items[j]).IsOFP = 1 then
             XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, 2] := TDataContent(dl.BigDataGrid.Items[j]).OFP + ' (ОФП)'
            else
             XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, 2] := TDataContent(dl.BigDataGrid.Items[j]).Exec;

            XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, 2].VerticalAlignment := 2;
            XL.WorkBooks[1].WorkSheets[1].Range['B'+IntToStr(k+(j+1)*2)+':B'+IntToStr(1+k+(j+1)*2)].Merge;
            XL.WorkBooks[1].WorkSheets[1].Range['B'+IntToStr(k+(j+1)*2)+':B'+IntToStr(1+k+(j+1)*2)].Borders.LineStyle := 1;
            XL.WorkBooks[1].WorkSheets[1].Range['B'+IntToStr(k+(j+1)*2)+':B'+IntToStr(1+k+(j+1)*2)].Borders.Weight := 2;
            XL.WorkBooks[1].WorkSheets[1].Range['B'+IntToStr(k+(j+1)*2)+':B'+IntToStr(1+k+(j+1)*2)].Borders.ColorIndex := 1;

            // обрисовка диапазона ячеек
          {  XL.WorkBooks[1].WorkSheets[1].Range['A3:' + chr(ord('C') + dl.BigDataGrid.Count) +
            IntToStr(i)].Borders.LineStyle := 1;
            XL.WorkBooks[1].WorkSheets[1].Range['A3:' + chr(ord('C') + dl.BigDataGrid.Count) +
            IntToStr(i)].Borders.Weight := 2;
            XL.WorkBooks[1].WorkSheets[1].Range['A3:' + chr(ord('C') + dl.BigDataGrid.Count) +
            IntToStr(i)].Borders.ColorIndex := 1; }

            for jj := 0 to TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Count - 1 do
            begin
              XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, (jj+1)*2+1] := TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res1;
              XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, (jj+1)*2+1].VerticalAlignment := 2;
              XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, (jj+1)*2+1].HorizontalAlignment := 3;

              XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2+1, (jj+1)*2+1] := TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res2;
              XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2+1, (jj+1)*2+1].VerticalAlignment := 2;
              XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2+1, (jj+1)*2+1].HorizontalAlignment := 3;

              case jj of
               0: s := 'C';
               1: s := 'E';
               2: s := 'G';
               3: s := 'I';
               4: s := 'K';
               5: s := 'M';
               6: s := 'O';
               7: s := 'Q';
               8: s := 'S';
               9: s := 'U';
               10: s := 'W';
               11: s := 'Y';
               12: s := 'AA';
               13: s := 'AC';
               14: s := 'AE';
               15: s:= 'AG';
               16: s:= 'AI';
               17: s:= 'AK';
               18: s:= 'AM';
               19: s:= 'AO';
               20: s:= 'AQ';
               21: s:= 'AS';
               22: s:= 'AU';
               23: s:= 'AW';
               24: s:= 'AY';
              end;
              XL.WorkBooks[1].WorkSheets[1].Range[s+IntToStr(k+(j+1)*2)+':'+s+IntToStr(k+(j+1)*2)].Borders.LineStyle := 1;
              XL.WorkBooks[1].WorkSheets[1].Range[s+IntToStr(k+(j+1)*2)+':'+s+IntToStr(k+(j+1)*2)].Borders.Weight := 2;
              XL.WorkBooks[1].WorkSheets[1].Range[s+IntToStr(k+(j+1)*2)+':'+s+IntToStr(k+(j+1)*2)].Borders.ColorIndex := 1;

              XL.WorkBooks[1].WorkSheets[1].Range[s+IntToStr(1+k+(j+1)*2)+':'+s+IntToStr(1+k+(j+1)*2)].Borders.LineStyle := 1;
              XL.WorkBooks[1].WorkSheets[1].Range[s+IntToStr(1+k+(j+1)*2)+':'+s+IntToStr(1+k+(j+1)*2)].Borders.Weight := 2;
              XL.WorkBooks[1].WorkSheets[1].Range[s+IntToStr(1+k+(j+1)*2)+':'+s+IntToStr(1+k+(j+1)*2)].Borders.ColorIndex := 1;

              if TDataContent(dl.BigDataGrid.Items[j]).IsOFP = 0 then
              begin
              XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, (jj+1)*2+2] := TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res3;
              XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, (jj+1)*2+2].VerticalAlignment := 2;
              XL.WorkBooks[1].WorkSheets[1].Cells[k+(j+1)*2, (jj+1)*2+2].HorizontalAlignment := 3;

              case jj of
               0: s := 'D';
               1: s := 'F';
               2: s := 'H';
               3: s := 'J';
               4: s := 'L';
               5: s := 'N';
               6: s := 'P';
               7: s := 'R';
               8: s := 'T';
               9: s := 'V';
               10: s := 'X';
               11: s := 'Z';
               12: s := 'AB';
               13: s := 'AD';
               14: s := 'AF';
               15: s:= 'AH';
               16: s:= 'AJ';
               17: s:= 'AL';
               18: s:= 'AN';
               19: s:= 'AP';
               20: s:= 'AR';
               21: s:= 'AT';
               22: s:= 'AV';
               23: s:= 'AX';
               24: s:= 'AZ';
              end;
              XL.WorkBooks[1].WorkSheets[1].Range[s+IntToStr(k+(j+1)*2)+':'+s+IntToStr(1+k+(j+1)*2)].Merge;
              XL.WorkBooks[1].WorkSheets[1].Range[s+IntToStr(k+(j+1)*2)+':'+s+IntToStr(1+k+(j+1)*2)].Borders.LineStyle := 1;
              XL.WorkBooks[1].WorkSheets[1].Range[s+IntToStr(k+(j+1)*2)+':'+s+IntToStr(1+k+(j+1)*2)].Borders.Weight := 2;
              XL.WorkBooks[1].WorkSheets[1].Range[s+IntToStr(k+(j+1)*2)+':'+s+IntToStr(1+k+(j+1)*2)].Borders.ColorIndex := 1;
              end;

            end;
           end;
           k := 2 + k + (dl.BigDataGrid.Count * 2);
           end;

         end;
       end;

  ProgressForm.Close;
  Screen.Cursor := crDefault;
  XL.Visible := true;
 except
  ProgressForm.Close;
  Screen.Cursor := crDefault;
 end;
 ST.Free;
end;

procedure TMainForm.ExportData;
var
  mensize, dtcount, len, i,j,jj : integer;
  dt : TDateTime;
  dl : TDataLeaf;
  execfound2, fnd, datafound, execfound, ofpfound, manfound : boolean;
  volume1, volume2, summ_ofpspeed, summ_ofpvolume, ofpspeed, ofpvolume, summ_speed, summ_volume, speed, volume : real;
  ManStream, M : TMemoryStream;
  h : tHandle;
  FirstWeekDay: Integer;
  FirstWeekDate: Integer;
  S1 , S2, S3, S4, S5 : TStringList;
  F : TFileStream;
  buffer : PAnsiChar;
  pbuffer : PByte;
  s : string;
begin
  try
     Screen.Cursor := crHourGlass;
     ManStream := TMemoryStream.Create;

     F := TFileStream.Create(TMDIChild(ActiveMDIChild).fname,fmOpenRead);
     F.Read(len,4);
     if len<=DBVersion then
     begin
      F.Read(len,4);
      mensize := len;
      while F.Position < mensize do
      begin
      // Размер записи
      F.Read(len,4);
      M := TMemoryStream.Create;
      pBuffer := AllocMem(len-4);
      F.Read(pbuffer^,len-4);
      M.Write(pbuffer^,len-4);
      FreeMem(pbuffer);
      M.Position := 0;
      // Фамилия
      M.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      M.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := StrPas(buffer);
      FreeMem(buffer);
      // Имя
      M.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      M.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := s + ' ' + StrPas(buffer);
      FreeMem(buffer);
      // Отчество
      M.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      M.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := s + ' ' + StrPas(buffer);
      FreeMem(buffer);
      fnd := false;
      for I := 0 to TMDIChild(ActiveMDIChild).BigDataLeaf.Count - 1 do
      if TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i]).menindex<>-1 then
       begin
        dl := TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i]);

        manfound := false;
        for j := 0 to DiagForm.ListView2.Items.Count - 1 do
        if dl.men = DiagForm.ListView2.Items[j].Subitems[0] then
         begin
           manfound := true;
           break;
         end;

        if (dl.men = s) and manfound then
         if DiagForm.RadioButton2.Checked or ((dl.mezo >= StrToInt(DiagForm.MezoB.Text)) and
         (dl.mezo <= StrToInt(DiagForm.MezoE.Text)) and (dl.micro >= StrToInt(DiagForm.MicroB.Text)) and
         (dl.micro <= StrToInt(DiagForm.MicroE.Text)) )then
         begin
          fnd := true;
          break;
         end;
       end;
      if fnd then
      begin
        M.Position := 0;
        len := M.Size + 4;
        ManStream.Write(len,4);
        ManStream.CopyFrom(M,M.Size);
      end;
      M.Free;
      end;
      F.Free;
     end;

   if copy(DiagForm.ImportExportName,Pos('.',DiagForm.ImportExportName),4) <> '.s40' then
     DiagForm.ImportExportName := DiagForm.ImportExportName + '.s40';

   F := TFileStream.Create(DiagForm.ImportExportName,fmCreate);
   len := DBVersion;
   F.Write(len,4);
   len := ManStream.Size;
   F.Write(len,4);
   ManStream.Position := 0;
   F.CopyFrom(ManStream,len);

   for I := 0 to TMDIChild(ActiveMDIChild).BigDataLeaf.Count - 1 do
   if TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i]).menindex<>-1 then
   begin
    dl := TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i]);

    manfound := false;
    for j := 0 to DiagForm.ListView2.Items.Count - 1 do
     if dl.men = DiagForm.ListView2.Items[j].Subitems[0] then
      begin
        manfound := true;
        break;
      end;

    if manfound then
         if DiagForm.RadioButton2.Checked or ((dl.mezo >= StrToInt(DiagForm.MezoB.Text)) and
         (dl.mezo <= StrToInt(DiagForm.MezoE.Text)) and (dl.micro >= StrToInt(DiagForm.MicroB.Text)) and
         (dl.micro <= StrToInt(DiagForm.MicroE.Text)) )then
     begin
      M := TMemoryStream.Create;
      TMDIChild(ActiveMDIChild).SaveDL(dl,M);
      M.Position := 0;
      len := M.Size;
      F.Write(len,4);
      F.CopyFrom(M,M.Size);
      M.Free;
     end;

   end;
   F.Free;
   Screen.Cursor := crDefault;
   MessageBOX(Handle, PChar('Экспорт данных завершен.'), Pchar(MainTitle) ,
   MB_ICONINFORMATION);
  except
   Screen.Cursor := crDefault;
   MessageBOX(Handle, PChar('Ошибка записи файла "'+DiagForm.ImportExportName+'"...'), Pchar(MainTitle) ,
   MB_ICONERROR);
  end;
end;

procedure TMainForm.ImportData;
var
  cntimport, curpos, curpos2, mensize, dtcount, len, i,j,jj : integer;
  dt, dt2, dtl, dt2l : TDateTime;
  dl : TDataLeaf;
  execfound2, fnd, datafound, execfound, ofpfound, manfound : boolean;
  volume1, volume2, summ_ofpspeed, summ_ofpvolume, ofpspeed, ofpvolume, summ_speed, summ_volume, speed, volume : real;
  MemHead, MemTail, ManStream, M, M2 : TMemoryStream;
  h : tHandle;
  FirstWeekDay: Integer;
  FirstWeekDate: Integer;
  S1 , S2, S3, S4, S5 : TStringList;
  F, F2 : TFileStream;
  buffer : PAnsiChar;
  pbuffer : PByte;
  s : string;
begin
  try
     cntimport := 0; // количество импортированных тренировок

     Screen.Cursor := crHourGlass;
     S1 := TStringList.Create;
     ManForm.fname := TMDIChild(ActiveMDIChild).fname;
     ManForm.ScanDBToStrings(S1);

     F := TFileStream.Create(DiagForm.ImportExportName,fmOpenRead);
     F.Read(len,4);
     if len<=DBVersion then
     begin
      F.Read(len,4);
      mensize := len;
      while F.Position < mensize do
      begin

      // Размер записи
      F.Read(len,4);
      M := TMemoryStream.Create;
      pBuffer := AllocMem(len-4);
      F.Read(pbuffer^,len-4);
      M.Write(pbuffer^,len-4);
      FreeMem(pbuffer);
      M.Position := 0;
      // Фамилия
      M.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      M.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := StrPas(buffer);
      FreeMem(buffer);
      // Имя
      M.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      M.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := s + ' ' + StrPas(buffer);
      FreeMem(buffer);
      // Отчество
      M.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      M.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := s + ' ' + StrPas(buffer);
      FreeMem(buffer);

      datafound := false;
      for I := 0 to DiagForm.ListView2.Items.Count - 1 do
      if DiagForm.ListView2.Items[i].SubItems[0] = s then
         begin
          datafound := true;
          break;
         end;

      if datafound then
      begin

      fnd := false;
      for I := 0 to S1.Count - 1 do
      if S1[i] = s then
         begin
          fnd := true;
          break;
         end;

      if not fnd then
      begin
        M.Position := 0;
        ManForm.AddMem(M);
      end;

      M.Free;

      end;

      end;

      F.Free;
     end;

   F := TFileStream.Create(DiagForm.ImportExportName,fmOpenRead);
   F.Read(len,4);
   if len > DBVersion then
    MessageBOX(Handle, PChar('Версия файла "'+DiagForm.ImportExportName+'" не поддерживается...'), PChar(MainTitle) ,
    MB_ICONWARNING)
   else
   begin

    ProgressForm.Caption := 'Импорт базы данных...';
    ProgressForm.pb.Position := 0;
    ProgressForm.Show;

    F.Read(len,4);
    F.Seek(len,soFromCurrent);
    while F.Position < F.Size do
    begin

     ProgressForm.pb.Position := Trunc(F.Position/F.Size*100);
//     ProgressForm.pb.Repaint;
     Application.ProcessMessages;

     curpos := F.Position;
     F.Read(len,4);
     M := TMemoryStream.Create;
     M.CopyFrom(F,len);

     M.Position := 0;
     M.Read(dt,8);
     M.Read(dtl,8);

     M.Read(len,4);
     M.Read(len,4);
     buffer := AllocMem(len+1);
     buffer[len] := Chr(0);
     M.Read(buffer^,len);
     buffer[len]:=Chr(0);
     s := StrPas(buffer);
     FreeMem(buffer);

     datafound := false;
     for I := 0 to DiagForm.ListView2.Items.Count - 1 do
     if DiagForm.ListView2.Items[i].SubItems[0] = s then
         begin
          datafound := true;
          break;
         end;

     if datafound then
     begin
         if DiagForm.RadioButton2.Checked or ((dl.mezo >= StrToInt(DiagForm.MezoB.Text)) and
         (dl.mezo <= StrToInt(DiagForm.MezoE.Text)) and (dl.micro >= StrToInt(DiagForm.MicroB.Text)) and
         (dl.micro <= StrToInt(DiagForm.MicroE.Text)) )then
      begin
       fnd := false;
       // Совпадение параметров - Запишем данные
       F2 := TFileStream.Create(TMDIChild(ActiveMDIChild).fname,fmOpenReadWrite);

       MemHead := TMemoryStream.Create;
       MemTail := TMemoryStream.Create;
       M2 := TMemoryStream.Create;

       F2.Read(len,4);
       F2.Read(len,4);
       F2.Seek(len,soFromCurrent);
       while F2.Position < F2.Size do
       begin
        fnd := false;
        curpos2 := F2.Position;
        F2.Read(len,4);
        M2.Clear;
        M2.Position := 0;
        M2.CopyFrom(F2,len);

        M2.Position := 0;
        M2.Read(dt2,8);
        M2.Read(dt2l,8);

        M2.Read(len,4);
        M2.Read(len,4);
        buffer := AllocMem(len+1);
        buffer[len] := Chr(0);
        M2.Read(buffer^,len);
        buffer[len]:=Chr(0);
        s := StrPas(buffer);
        FreeMem(buffer);

        datafound := false;
        for I := 0 to DiagForm.ListView2.Items.Count - 1 do
        if DiagForm.ListView2.Items[i].SubItems[0] = s then
         begin
          datafound := true;
          break;
         end;

        if datafound then
         if (dt2 = dt) and (dt2l = dtl) then
          begin
            MemTail.Clear;
            MemHead.Clear;
            fnd := true;
            MemTail.CopyFrom(F2,F2.Size-F2.Position);
            F2.Position := 0;
            MemHead.CopyFrom(F2,curpos2);
            break;
          end;

       end;

       if not fnd then
         inc(cntimport);

       if fnd then
       begin

         // Удалим старые данные
         F2.Free;
         F2 := TFileStream.Create(TMDIChild(ActiveMDIChild).fname,fmCreate);
         MemHead.Position := 0;
         MemTail.Position := 0;
         F2.CopyFrom(MemHead,MemHead.Size);
         F2.CopyFrom(MemTail,MemTail.Size);
       end;

       M2.Free;
       MemHead.Free;
       MemTail.Free;

       F2.Seek(0,soFromEnd);
       M.Position := 0;
       len := M.Size;
       F2.Write(len,4);
       F2.CopyFrom(M,M.Size);
       F2.Free;

      end;

     end;

     M.Free;
    end;
   end;


   F.Free;


   ProgressForm.Close;
   Screen.Cursor := crDefault;
   MessageBOX(Handle, PChar('Импорт данных завершен. Импортировано тренировок - ' + IntToStr(cntimport) + '.'), Pchar(MainTitle) ,
   MB_ICONINFORMATION);

   // Обновление информации
   UpdateDataBase;

  except
   ProgressForm.Close;
   Screen.Cursor := crDefault;
   MessageBOX(Handle, PChar('Ошибка записи файла "'+DiagForm.ImportExportName+'"...'), Pchar(MainTitle) ,
   MB_ICONERROR);
  end;
end;

procedure TMainForm.UpdateDataBase;
begin
   TMDIChild(ActiveMDIChild).ClearPanel;
   TMDIChild(ActiveMDIChild).ExecList.Clear;
   while TMDIChild(ActiveMDIChild).BigDataLeaf.Count>0 do
    TMDIChild(ActiveMDIChild).BigDataLeaf.Delete(TMDIChild(ActiveMDIChild).BigDataLeaf.Count-1);
   TMDIChild(ActiveMDIChild).BigDataLeaf.Clear;
   TMDIChild(ActiveMDIChild).ManStream.Clear;
   TMDIChild(ActiveMDIChild).MenList.Clear;
   TMDIChild(ActiveMDIChild).OpenFile;
end;

procedure TMainForm.ToolButton14Click(Sender: TObject);
begin
 if not IfRegOK then
  Exit;

 if Copy(ActiveMDIChild.Caption,1,9) = 'Диаграмма' then
 begin
  SelectForm.ComboBox1.Items.Clear;
  if TOutputForm(ActiveMDIChild).Panel1.Visible then
   SelectForm.ComboBox1.Items.Add('Текст');
  if TOutputForm(ActiveMDIChild).PageControl1.Visible then
   SelectForm.ComboBox1.Items.Add('Диаграмма');
  if TOutputForm(ActiveMDIChild).StringGrid1.Visible then
   SelectForm.ComboBox1.Items.Add('Таблица');
  if SelectForm.ComboBox1.Items.Count>0 then
   SelectForm.ComboBox1.ItemIndex := 0;
  SelectForm.Caption := 'Экспорт в Microsoft Excel...';
  SelectForm.GroupBox1.Caption := 'Укажите какую информацию следует экспортировать';

  if SelectForm.ComboBox1.Items.Count=1 then
  begin
   if SelectForm.ComboBox1.Text = 'Текст' then
   TOutputForm(ActiveMDIChild).ActivateExcel_TopTable
  else
   if SelectForm.ComboBox1.Text = 'Таблица' then
   TOutputForm(ActiveMDIChild).ActivateExcel_BottomTable
  else
  if SelectForm.ComboBox1.Text = 'Диаграмма' then
   if TOutputForm(ActiveMDIChild).PageControl1.ActivePageIndex = 0 then
    TOutputForm(ActiveMDIChild).ActivateExcel
   else
    MessageBOX(Handle, PChar('Для эксорта данных в Microsoft Excel следует выбрать первую диаграмму.'), PChar(MainTitle) , MB_OK or MB_ICONWARNING);
  end
  else
  if SelectForm.ShowModal = mrOk then
  begin
   if SelectForm.ComboBox1.Text = 'Текст' then
   TOutputForm(ActiveMDIChild).ActivateExcel_TopTable
  else
   if SelectForm.ComboBox1.Text = 'Таблица' then
   TOutputForm(ActiveMDIChild).ActivateExcel_BottomTable
  else
  if SelectForm.ComboBox1.Text = 'Диаграмма' then
   if TOutputForm(ActiveMDIChild).PageControl1.ActivePageIndex = 0 then
    TOutputForm(ActiveMDIChild).ActivateExcel
   else
    MessageBOX(Handle, PChar('Для эксорта данных в Microsoft Excel следует выбрать первую диаграмму.'), PChar(MainTitle) , MB_OK or MB_ICONWARNING);
  end;

 end
 else
 begin
  DiagForm.fname :=  TMDIChild(ActiveMDIChild).fname;
  DiagForm.MasterType := 3; // Экспорт в Excel
  if DiagForm.ShowModal = mrOk then
   ExportExcel;
 end;
end;

procedure TMainForm.ToolButton4Click(Sender: TObject);
begin
 N9Click(Sender);
end;

procedure TMainForm.ToolButton8Click(Sender: TObject);
begin
 DiagForm.fname :=  TMDIChild(ActiveMDIChild).fname;
 DiagForm.MasterType := 0;
 if DiagForm.ShowModal = mrOk then
  GenerateData;
end;

procedure TMainForm.ToolButton9Click(Sender: TObject);
begin
 N12Click(Sender);
end;

procedure TMainForm.IdFTP1Work(ASender: TObject; AWorkMode: TWorkMode;
  AWorkCount: Int64);
begin
 ProgressForm.pb.StepBy(AWorkCount);
 ProgressForm.pb.Repaint;
 Application.ProcessMessages;
end;

procedure TMainForm.IdFTP1WorkEnd(ASender: TObject; AWorkMode: TWorkMode);
begin
 ProgressForm.Close;

 if not ImportMode then
  ImportFile.Free;

 if ImportMode then
  MessageBOX(Handle, PChar('Импорт локальной базы данных завершен.'), PChar(MainTitle), MB_ICONINFORMATION)
 else
  MessageBOX(Handle, PChar('Экспорт локальной базы данных завершен.'), PChar(MainTitle), MB_ICONINFORMATION);

 if ImportMode then
 try
  ImportMode := false;
  ImportFileMem.SaveToFile(ExtractFilePath(paramstr(0)) + 'local.mtx');
  ImportFileMem.Free;
  TMDIChild(ActiveMDIChild).Modified := false;
  TMDIChild(ActiveMDIChild).Close;
  Timer1.Enabled := true;

 except;
 end;

 try
  IdFTP1.Disconnect;
 except
 end;

end;

procedure TMainForm.IdFTP3AfterGet(ASender: TObject; AStream: TStream);
begin
 try
  IdFTP3.Disconnect;
 except
 end;
end;

procedure TMainForm.IdFTP3Work(ASender: TObject; AWorkMode: TWorkMode;
  AWorkCount: Int64);
begin
 if ImportMode then
 begin
  ProgressForm.pb.StepBy(AWorkCount);
  ProgressForm.pb.Repaint;
  Application.ProcessMessages;
 end;
end;

procedure TMainForm.IdFTP3WorkEnd(ASender: TObject; AWorkMode: TWorkMode);
begin
 if ImportMode then
 try

  ImportMode := false;
  ProgressForm.Close;

  MessageBOX(Handle, PChar('Импорт базы данных завершен.'), PChar(MainTitle), MB_ICONINFORMATION);

  ImportFileMem.SaveToFile(ExtractFilePath(paramstr(0)) + ImportfileName);
  ImportFileMem.Free;


  Timer2.Enabled := true;

 except;
 end;

end;

function TMainForm.IfRegOK : boolean;
begin
 if not RegDone then
   begin
    MessageBOX(Handle, PChar(demostr), PChar(MainTitle) ,
        MB_ICONWARNING);
   end;
 Result := RegDone;
end;


procedure TMainForm.FileSave1Execute(Sender: TObject);
var
s : string;
begin
 if not IfRegOK then
  Exit;

 // Сохранение справочника Видов спорта и упражнений
 if ActiveMDIChild.Caption = ExecTitle then
  TFormSpr(ActiveMDIChild).SaveExecData
 else
 if ActiveMDIChild.Caption = 'Ввод данных' then
  TMDIChild(ActiveMDIChild).SaveDataBase
 else
 if Copy(ActiveMDIChild.Caption,1,9) = 'Диаграмма' then
 begin

  SelectForm.ComboBox1.Items.Clear;
  if TOutputForm(ActiveMDIChild).Panel1.Visible then
   SelectForm.ComboBox1.Items.Add('Текст');
  if TOutputForm(ActiveMDIChild).PageControl1.Visible then
   SelectForm.ComboBox1.Items.Add('Диаграмма');
  if TOutputForm(ActiveMDIChild).StringGrid1.Visible then
   SelectForm.ComboBox1.Items.Add('Таблица');
  if SelectForm.ComboBox1.Items.Count>0 then
   SelectForm.ComboBox1.ItemIndex := 0;
  SelectForm.Caption := 'Сохранить...';
  SelectForm.GroupBox1.Caption := 'Укажите какую информацию следует сохранить';

  if SelectForm.ComboBox1.Items.Count=1 then
  begin
   if SelectForm.ComboBox1.Text = 'Текст' then
   begin
   SaveDialog1.FileName := '';
   SaveDialog1.Filter := 'Вэб-страница (*.html)|*.html';
   SaveDialog1.Title := 'Сохранить текст';
   if SaveDialog1.Execute then
   begin
      if ExtractFileExt(SaveDialog1.FileName) = '.html' then
       s := SaveDialog1.FileName
      else
       s := SaveDialog1.FileName + '.html';
      if FileExists(s) then
      begin
       if MessageBOX(Handle, PChar('Файл '+s+' уже существует. Заменить?'), PChar(MainTitle) ,
        MB_YESNO or MB_ICONQUESTION) = IDYes then
         ListViewConfHTML(TOutputForm(ActiveMDIChild).lv,s,true);
      end
      else
        ListViewConfHTML(TOutputForm(ActiveMDIChild).lv,s,true);
   end;
   end
   else
   if SelectForm.ComboBox1.Text = 'Таблица' then
   begin
   SaveDialog1.FileName := '';
   SaveDialog1.Filter := 'Вэб-страница (*.html)|*.html';
   SaveDialog1.Title := 'Сохранить таблицу';
   if SaveDialog1.Execute then
   begin
      if ExtractFileExt(SaveDialog1.FileName) = '.html' then
       s := SaveDialog1.FileName
      else
       s := SaveDialog1.FileName + '.html';
      if FileExists(s) then
      begin
       if MessageBOX(Handle, PChar('Файл '+s+' уже существует. Заменить?'), PChar(MainTitle) ,
        MB_YESNO or MB_ICONQUESTION) = IDYes then
         StringGridConfHTML(TOutputForm(ActiveMDIChild).StringGrid1,s, true);
      end
      else
       StringGridConfHTML(TOutputForm(ActiveMDIChild).StringGrid1,s, true);
   end;
   end
   else
   if SelectForm.ComboBox1.Text = 'Диаграмма' then
   begin
   SaveDialog1.FileName := '';
   SaveDialog1.Filter := 'Растровый рисунок  (*.bmp)|*.bmp';
   SaveDialog1.Title := 'Сохранить диаграмму';
   case TOutputForm(ActiveMDIChild).PageControl1.ActivePageIndex of
    0 : SaveDialog1.FileName := 'Динамика';
    1 : SaveDialog1.FileName := 'Величины';
    2 : SaveDialog1.FileName := 'Соотношение упражнений';
    3 : SaveDialog1.FileName := 'Соотношение СФП и ОФП';
    4 : SaveDialog1.FileName := 'Восстановление';
    5 : SaveDialog1.FileName := 'Контроль';
   end;
   if SaveDialog1.Execute then
     begin
      if ExtractFileExt(SaveDialog1.FileName) = '.bmp' then
       s := SaveDialog1.FileName
      else
       s := SaveDialog1.FileName + '.bmp';
       if MessageBOX(Handle, PChar('Файл '+s+' уже существует. Заменить?'), PChar(MainTitle) ,
        MB_YESNO or MB_ICONQUESTION) = IDYes then
    begin
     case TOutputForm(ActiveMDIChild).PageControl1.ActivePageIndex of
      0: TOutputForm(ActiveMDIChild).Chart1.SaveToBitmapFile(s);
      1: TOutputForm(ActiveMDIChild).Chart2.SaveToBitmapFile(s);
      2: TOutputForm(ActiveMDIChild).Chart3.SaveToBitmapFile(s);
      3: TOutputForm(ActiveMDIChild).Chart4.SaveToBitmapFile(s);
      4: TOutputForm(ActiveMDIChild).Chart5.SaveToBitmapFile(s);
      5: TOutputForm(ActiveMDIChild).Chart6.SaveToBitmapFile(s);
     end;
    end
    else
     case TOutputForm(ActiveMDIChild).PageControl1.ActivePageIndex of
      0: TOutputForm(ActiveMDIChild).Chart1.SaveToBitmapFile(s);
      1: TOutputForm(ActiveMDIChild).Chart2.SaveToBitmapFile(s);
      2: TOutputForm(ActiveMDIChild).Chart3.SaveToBitmapFile(s);
      3: TOutputForm(ActiveMDIChild).Chart4.SaveToBitmapFile(s);
      4: TOutputForm(ActiveMDIChild).Chart5.SaveToBitmapFile(s);
      5: TOutputForm(ActiveMDIChild).Chart6.SaveToBitmapFile(s);
     end;
   end;
   end;


  end
  else

  if SelectForm.ShowModal = mrOk then
  begin
   if SelectForm.ComboBox1.Text = 'Текст' then
   begin
   SaveDialog1.FileName := '';
   SaveDialog1.Filter := 'Вэб-страница (*.html)|*.html';
   SaveDialog1.Title := 'Сохранить текст';
   if SaveDialog1.Execute then
   begin
       if ExtractFileExt(SaveDialog1.FileName) = '.html' then
       s := SaveDialog1.FileName
      else
       s := SaveDialog1.FileName + '.html';
      if FileExists(s) then
      begin
       if MessageBOX(Handle, PChar('Файл '+s+' уже существует. Заменить?'), PChar(MainTitle) ,
        MB_YESNO or MB_ICONQUESTION) = IDYes then
         ListViewConfHTML(TOutputForm(ActiveMDIChild).lv,s,true);
      end
      else
        ListViewConfHTML(TOutputForm(ActiveMDIChild).lv,s,true);
   end;
   end
   else
   if SelectForm.ComboBox1.Text = 'Таблица' then
   begin
   SaveDialog1.FileName := '';
   SaveDialog1.Filter := 'Вэб-страница (*.html)|*.html';
   SaveDialog1.Title := 'Сохранить таблицу';
   if SaveDialog1.Execute then
   begin
      if ExtractFileExt(SaveDialog1.FileName) = '.html' then
       s := SaveDialog1.FileName
      else
       s := SaveDialog1.FileName + '.html';
      if FileExists(s) then
      begin
       if MessageBOX(Handle, PChar('Файл '+s+' уже существует. Заменить?'), PChar(MainTitle) ,
        MB_YESNO or MB_ICONQUESTION) = IDYes then
         StringGridConfHTML(TOutputForm(ActiveMDIChild).StringGrid1,s,true);
      end
      else
       StringGridConfHTML(TOutputForm(ActiveMDIChild).StringGrid1,s,true);
   end;
   end
   else
   if SelectForm.ComboBox1.Text = 'Диаграмма' then
   begin
   SaveDialog1.FileName := '';
   SaveDialog1.Filter := 'Растровый рисунок  (*.bmp)|*.bmp';
   SaveDialog1.Title := 'Сохранить диаграмму';
   case TOutputForm(ActiveMDIChild).PageControl1.ActivePageIndex of
    0 : SaveDialog1.FileName := 'Динамика';
    1 : SaveDialog1.FileName := 'Величины';
    2 : SaveDialog1.FileName := 'Соотношение упражнений';
    3 : SaveDialog1.FileName := 'Соотношение СФП и ОФП';
    4 : SaveDialog1.FileName := 'Восстановление';
    5 : SaveDialog1.FileName := 'Контроль';
   end;
   if SaveDialog1.Execute then
    begin
      if ExtractFileExt(SaveDialog1.FileName) = '.bmp' then
       s := SaveDialog1.FileName
      else
       s := SaveDialog1.FileName + '.bmp';
       if MessageBOX(Handle, PChar('Файл '+s+' уже существует. Заменить?'), PChar(MainTitle) ,
        MB_YESNO or MB_ICONQUESTION) = IDYes then
    begin
     case TOutputForm(ActiveMDIChild).PageControl1.ActivePageIndex of
      0: TOutputForm(ActiveMDIChild).Chart1.SaveToBitmapFile(s);
      1: TOutputForm(ActiveMDIChild).Chart2.SaveToBitmapFile(s);
      2: TOutputForm(ActiveMDIChild).Chart3.SaveToBitmapFile(s);
      3: TOutputForm(ActiveMDIChild).Chart4.SaveToBitmapFile(s);
      4: TOutputForm(ActiveMDIChild).Chart5.SaveToBitmapFile(s);
      5: TOutputForm(ActiveMDIChild).Chart6.SaveToBitmapFile(s);
     end;
    end
    else
     case TOutputForm(ActiveMDIChild).PageControl1.ActivePageIndex of
      0: TOutputForm(ActiveMDIChild).Chart1.SaveToBitmapFile(s);
      1: TOutputForm(ActiveMDIChild).Chart2.SaveToBitmapFile(s);
      2: TOutputForm(ActiveMDIChild).Chart3.SaveToBitmapFile(s);
      3: TOutputForm(ActiveMDIChild).Chart4.SaveToBitmapFile(s);
      4: TOutputForm(ActiveMDIChild).Chart5.SaveToBitmapFile(s);
      5: TOutputForm(ActiveMDIChild).Chart6.SaveToBitmapFile(s);
     end;
    end;
   end;
  end;
 end
 else
  TMDIChild(ActiveMDIChild).SaveDataBase;
end;

procedure TMainForm.RegComplete;
var
 F : TMemoryStream;
 success : boolean;
 s : string;
 st : TStringList;
 i : integer;
 Registry : TRegistry;
begin
  h12 := $0;
  h34 := $0;
  RegDone := false;
  s := '222:747';
  try
    st := TStringList.Create;
    F := TMemoryStream.Create;
    success := DownloadURL_NOCache_mem('http://'+s+'@startedu.ru/sport4/users.txt',F);
  finally
    F.Position := 0;
    st.LoadFromStream(F);
    F.Position := 0;
    UsersList.Clear;
    UsersList.LoadFromStream(F);
    F.Free;
  end;
  if success then
  begin
   for i:= 0 to st.Count - 1 do
   begin
    s := copy(st[i],2,pos(']',st[i])-2);
    if s = RegData then
    begin
     RegDone := true;
     RegFIO := copy(st[i],pos(']',st[i])+2,length(st[i])-pos(']',st[i])-2);
     Registry := TRegistry.Create;
     try
      Registry.RootKey := HKEY_CURRENT_USER;
      if Registry.OpenKey('\Software\Siberia-Soft\Sport4', True)
      then begin
       Registry.WriteString('RegFIO',RegFIO);
      end;
     finally
      Registry.CloseKey;
      Registry.Free;
     end;
     break;
    end;
   end;
   st.Free;
  end
  else
   RegDone :=  Length(RegData)>0;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
 Registry : TRegistry;
begin
 AboutBox2 := TAboutBox.Create(self);
 AboutBox2.BorderStyle := bsNone;
 AboutBox2.Show;
 Application.ProcessMessages;

 UsersList := TStringList.Create;

 RegistryS30(ExtractFilePath(paramstr(0)));

 Registry := TRegistry.Create;
  try
    Registry.RootKey := HKEY_CURRENT_USER;
    if Registry.OpenKeyReadOnly('\Software\Siberia-Soft\Sport4') then
    begin
     if Registry.ValueExists('Registry') then
     begin
      RegData := Registry.ReadString('Registry');
      RegFIO := Registry.ReadString('RegFIO');
     end
     else
     begin
      RegData := '';
      RegFIO := '';
     end;
    end;
  finally
    Registry.CloseKey;
    Registry.Free;
  end;

 if IsConnectedToInternet then
  RegComplete
 else
  RegDone := length(RegData)>0;

 N7.Enabled := IsConnectedToInternet;

 Caption := MainTitle;
 Application.Title := MainTitle;
end;

function TMainForm.IsUpdateRequared(fname : string):boolean;
var
 F, F2 : TFileStream;
 buffer : PAnsiChar;
 us, i, len : integer;
 s : string;
 Registry : TRegistry;
begin
  i := 0;
  Registry := TRegistry.Create;
  try
    Registry.RootKey := HKEY_CURRENT_USER;
    if Registry.OpenKey('\Software\Siberia-Soft\Sport3', True) then
     if Registry.ValueExists('UpdateSize') then
      i := Registry.ReadInteger('UpdateSize');
  finally
    Registry.CloseKey;
    Registry.Free;
  end;
  try
     F := TFileStream.Create(fname,fmOpenRead);
     us := F.Size;
  finally
     F.Free;
  end;
  Result := us <> i;
end;

procedure TMainForm.FormShow(Sender: TObject);
begin
 if fileExists(ExtractFilePath(paramstr(0)) + 'download.file') then
   DeleteFile(ExtractFilePath(paramstr(0)) + 'download.file');

//AdminMode := paramstr(2)='A';
 AdminMode := true;

 N6.Visible := AdminMode;
 N10.Visible := AdminMode;

 if AdminMode then
 begin
  Caption := Caption + ' - Администратор';
 end
 else
 begin
  Caption := Caption + ' - Клиент';
 end;

 if paramstr(3)='' then
  CreateMDIChild('Ввод данных',ExtractFilePath(paramstr(0)) + 'local.mtx')
 else
  CreateMDIChild(paramstr(3), paramstr(3));

 AboutBox2.Close;
 AboutBox2.Free;
end;

procedure TMainForm.HelpAbout1Execute(Sender: TObject);
begin
  if RegDone then
   AboutBox.RegLabel.Caption := regstr + ' ('+RegFIO+')'
  else
   AboutBox.RegLabel.Caption := demostr;
  AboutBox.ShowModal;
end;

procedure TMainForm.N10Click(Sender: TObject);
var
 Registry : TRegistry;
begin
 Registry := TRegistry.Create;
  try
    Registry.RootKey := HKEY_CURRENT_USER;
    if Registry.OpenKeyReadOnly('\Software\Siberia-Soft\Sport4') then
    begin
     if Registry.ValueExists('URL') then
     begin
      OptionForm.Edit1.Text := Registry.ReadString('URL');
     end
     else
      OptionForm.Edit1.Text := 'http://';
    end;
  finally
    Registry.CloseKey;
    Registry.Free;
  end;

 if OptionForm.ShowModal = mrOk then
  begin
   Registry := TRegistry.Create;
   try
    Registry.RootKey := HKEY_CURRENT_USER;
    if Registry.OpenKey('\Software\Siberia-Soft\Sport4', True)
    then
     Registry.WriteString('URL',OptionForm.Edit1.Text);
   finally
    Registry.CloseKey;
    Registry.Free;
   end;
  end;
end;

procedure TMainForm.N11Click(Sender: TObject);
begin
 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'ofp.mtx';
 RazrDlg.stitle := 'ОФП';
 RazrDlg.title := 'Общая физическая подготовка';
 RazrDlg.ShowModal;
end;

procedure TMainForm.N12Click(Sender: TObject);
var
  i : integer;
  fnd : boolean;
begin
  fnd := false;
  for i := 0 to MDIChildCount - 1 do
   begin
     if  MDIChildren[i].Caption = 'Ввод данных' then
       begin
        fnd := true;
        Break;
       end;
   end;
  if not fnd then
    CreateMDIChild('Ввод данных',ExtractFilePath(paramstr(0)) + 'local.mtx')
  else
   begin
    MDIChildren[i].Show;
   end;
end;

procedure TMainForm.N14Click(Sender: TObject);
begin
 ToolButton8Click(Sender);
end;

procedure TMainForm.N16Click(Sender: TObject);
begin
 ToolButton12Click(Sender);
end;

procedure TMainForm.N17Click(Sender: TObject);
begin
 ToolButton11Click(Sender);
end;

procedure TMainForm.N18Click(Sender: TObject);
var
 s : string;
 F : TFileStream;
 ManStream : TMemoryStream;
 len : integer;
begin
   if not IfRegOK then exit;
//   if not AdminMode then exit;

   SaveDialog1.Filter := 'База данных Спорт 4.0  (*.s40)|*.s40';
   SaveDialog1.Title := 'Сохранить базу данных';
   SaveDialog1.FileName := TMDIChild(ActiveMDIChild).ComboBox1.Text;
   if SaveDialog1.Execute then
    begin
     ManStream := TMemoryStream.Create;
     s := TMDIChild(ActiveMDIChild).fname;
     TMDIChild(ActiveMDIChild).ReadManStream(s,ManStream);
     try
      if copy(SaveDialog1.FileName,Pos('.',SaveDialog1.FileName),4) = '.s40' then
       TMDIChild(ActiveMDIChild).fname := SaveDialog1.FileName
      else
       TMDIChild(ActiveMDIChild).fname := SaveDialog1.FileName + '.s40';
      if FileExists(TMDIChild(ActiveMDIChild).fname) then
      begin
       if MessageBOX(Handle, PChar('Файл '+TMDIChild(ActiveMDIChild).fname+' уже существует. Заменить?'), PChar(MainTitle) ,
        MB_YESNO or MB_ICONQUESTION) = IDYes then
        try

         F := TFileStream.Create(TMDIChild(ActiveMDIChild).fname,fmCreate);
         len := DBVersion;
         F.Write(len,4);
         len := ManStream.Size;
         F.Write(len,4);
         ManStream.Position := 0;
         F.CopyFrom(ManStream,len);
        finally
         F.Free;
        end;
        TMDIChild(ActiveMDIChild).SaveDataBase;
      end
      else
      begin
        try
         F := TFileStream.Create(TMDIChild(ActiveMDIChild).fname,fmCreate);
         len := DBVersion;
         F.Write(len,4);
         len := ManStream.Size;
         F.Write(len,4);
         ManStream.Position := 0;
         F.CopyFrom(ManStream,len);
        finally
         F.Free;
        end;
        TMDIChild(ActiveMDIChild).SaveDataBase;
      end;

      TMDIChild(ActiveMDIChild).fname := s;
     except
      TMDIChild(ActiveMDIChild).fname := s;
     end;
     ManStream.Free;
    end;
end;

procedure TMainForm.N19Click(Sender: TObject);
var
 Registry : TRegistry;
begin
if IsConnectedToInternet then
 begin
 RegForm.LabeledEdit2.Text := RegData;
 RegForm.LabeledEdit2.Enabled := not RegDone;
 if RegForm.ShowModal = mrOk then
  begin
   RegData := RegForm.LabeledEdit2.Text;

   Registry := TRegistry.Create;
   try
    Registry.RootKey := HKEY_CURRENT_USER;
    if Registry.OpenKey('\Software\Siberia-Soft\Sport4', True)
    then begin
     Registry.WriteString('Registry',RegData);
    end;
   finally
    Registry.CloseKey;
    Registry.Free;
   end;

   RegComplete;

   if RegDone then
    MessageBOX(Handle, thstr, MainTitle , MB_ICONINFORMATION);
  end;
 end
 else
  MessageBOX(Handle, 'Отсутствует активное подключение к сети Интернет.', MainTitle , MB_ICONWARNING);
end;

procedure TMainForm.N24Click(Sender: TObject);
begin
 ToolButton14Click(Sender);
end;

procedure TMainForm.N26Click(Sender: TObject);
begin
 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'control.mtx';
 RazrDlg.stitle := 'контроль';
 RazrDlg.title := 'Контроль';
 RazrDlg.ShowModal;
end;

procedure TMainForm.N27Click(Sender: TObject);
var
  i : integer;
  fnd : boolean;
begin
  if Copy(ActiveMDIChild.Caption,1,5) = 'Отчет' then
   THTMLForm(ActiveMDIChild).wb.ExecWB(OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER);
{  fnd := false;
  for i := 0 to MDIChildCount - 1 do
   begin
     if  MDIChildren[i].Caption = 'Предварительный просмотр - ' + ActiveMDIChild.Caption then
       begin
        fnd := true;
        Break;
       end;
   end;
  if not fnd then
   CreatePreviewChild(ActiveMDIChild.Caption)
  else
   begin
    MDIChildren[i].Show;
   end;  }
end;

procedure TMainForm.N30Click(Sender: TObject);
begin
 TMDIChild(ActiveMDIChild).AddMezoCicleEmpty(0,true);
end;

procedure TMainForm.N3Click(Sender: TObject);
var
  FormSpr: TFormSpr;
  i : integer;
  fnd : boolean;
begin
  fnd := false;
  for i := 0 to MDIChildCount - 1 do
   begin
     if  MDIChildren[i].Caption = ExecTitle then
       begin
        fnd := true;
        Break;
       end;
   end;
  if not fnd then
   begin
    FormSpr := TFormSpr.Create(Application);
    DockTabSet1.Visible := true;
    DockTabSet1.Tabs.Add(ExecTitle + '          ');
    DockTabSet1.TabIndex := DockTabSet1.Tabs.Count-1;

    FormSpr.Caption := ExecTitle;
   end
  else
   begin
    MDIChildren[i].Show;
   end;
end;

procedure TMainForm.N4Click(Sender: TObject);
begin
 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'razr.mtx';
 RazrDlg.stitle := 'разряд';
 RazrDlg.title := 'Спортивные разряды';
 RazrDlg.ShowModal;
end;

procedure TMainForm.ProcessInternetUpdate;
var
 s : string;
 F, F2 : TFileStream;
 Registry : TRegistry;
 success : boolean;
begin
 Registry := TRegistry.Create;
  try
    Registry.RootKey := HKEY_CURRENT_USER;
    if Registry.OpenKeyReadOnly('\Software\Siberia-Soft\Sport4') then
    begin
     if Registry.ValueExists('URL') then
     begin
      s := Registry.ReadString('URL');
     end
     else
      s := '';
    end;
  finally
    Registry.CloseKey;
    Registry.Free;
  end;
  if Length(s) > 0 then
  begin
    try
      F := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'download.file',fmCreate);
      success := DownloadURL_NOCache(s,F);
    finally
      F.Free;
    end;
    if not success then
       MessageBOX(0, 'Ошибка при загрузке обновления.', MainTitle , MB_ICONWARNING)
 //      sb.Panels[1].Text:= 'Ошибка при загрузке обновления.'
    else
    begin
       if IsUpdateRequared(ExtractFilePath(paramstr(0)) + 'download.file') then
       begin
        MessageBOX(0, 'Обновление успешно загружено. Для продолжения необходимо выйти из программы.', MainTitle , MB_ICONINFORMATION);
        //sb.Panels[1].Text:= 'Обновление успешно загружено. Для продолжения необходимо выйти из программы.';
        Close;
       end
       else
        deletefile(ExtractFilePath(paramstr(0)) + 'download.file');
    end;
  end;
end;

procedure TMainForm.N7Click(Sender: TObject);
var
 s : string;
 F, F2 : TFileStream;
 Registry : TRegistry;
 success : boolean;
begin
if IsConnectedToInternet then
begin
 Timer1.Enabled := false;
 ProcessInternetUpdate;
end
else
begin
 OpenDialog1.Filter := 'Обновление программы Спорт 4 (*.data4)|*.data4';
 OpenDialog1.FileName := 'sportupdate.data4';
 OpenDialog1.Title := 'Обновление программы';
 if OpenDialog1.Execute then
  try
      F := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'download.file',fmCreate);
      F2  := TFileStream.Create(OpenDialog1.FileName,fmOpenRead);
      F.CopyFrom(F2,F2.Size);
      F2.Free;
      F.Free;
      if IsUpdateRequared(ExtractFilePath(paramstr(0)) + 'download.file') then
      begin
       MessageBOX(Handle, 'Обновление успешно загружено. Для продолжения необходимо выйти из программы.', MainTitle , MB_TOPMOST or MB_ICONINFORMATION);
       Close;
      end
      else
       deletefile(ExtractFilePath(paramstr(0)) + 'download.file');
  except
      MessageBOX(Handle, 'Ошибка при загрузке обновления.', MainTitle , MB_TOPMOST or MB_ICONWARNING);
  end;
end;
end;

procedure TMainForm.N8Click(Sender: TObject);
begin
 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'relax.mtx';
 RazrDlg.stitle := 'восстановление';
 RazrDlg.title := 'Виды восстановлений';
 RazrDlg.ShowModal;
end;

procedure TMainForm.N9Click(Sender: TObject);
begin


  if Copy(ActiveMDIChild.Caption,1,5) = 'Отчет' then
   THTMLForm(ActiveMDIChild).wb.ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER);

{ if pd.Execute then
  if Copy(ActiveMDIChild.Caption,1,9) = 'Диаграмма' then
  begin

  SelectForm.ComboBox1.Items.Clear;
  if TOutputForm(ActiveMDIChild).Panel1.Visible then
   SelectForm.ComboBox1.Items.Add('Текст');
  if TOutputForm(ActiveMDIChild).PageControl1.Visible then
   SelectForm.ComboBox1.Items.Add('Диаграмма');
  if TOutputForm(ActiveMDIChild).StringGrid1.Visible then
   SelectForm.ComboBox1.Items.Add('Таблица');
  if SelectForm.ComboBox1.Items.Count>0 then
   SelectForm.ComboBox1.ItemIndex := 0;
  SelectForm.Caption := 'Печать...';
  SelectForm.GroupBox1.Caption := 'Укажите какую информацию следует распечатать';

  if SelectForm.ComboBox1.Items.Count=1 then
  begin
   if SelectForm.ComboBox1.Text = 'Текст' then
    PrintListView(TOutputForm(ActiveMDIChild).lv, 'Текст', poPortrait)
  else
   if SelectForm.ComboBox1.Text = 'Таблица' then
    PrintStringGrid(TOutputForm(ActiveMDIChild).StringGrid1, 'Таблица', poPortrait)
  else
  if SelectForm.ComboBox1.Text = 'Диаграмма' then
     case TOutputForm(ActiveMDIChild).PageControl1.ActivePageIndex of
      0: TOutputForm(ActiveMDIChild).Chart1.Print;
      1: TOutputForm(ActiveMDIChild).Chart2.Print;
      2: TOutputForm(ActiveMDIChild).Chart3.Print;
      3: TOutputForm(ActiveMDIChild).Chart4.Print;
      4: TOutputForm(ActiveMDIChild).Chart5.Print;
     end;
  end
  else
  if SelectForm.ShowModal = mrOk then
  begin
   if SelectForm.ComboBox1.Text = 'Текст' then
    PrintListView(TOutputForm(ActiveMDIChild).lv, 'Текст', poPortrait)
  else
   if SelectForm.ComboBox1.Text = 'Таблица' then
    PrintStringGrid(TOutputForm(ActiveMDIChild).StringGrid1, 'Таблица', poPortrait)
  else
  if SelectForm.ComboBox1.Text = 'Диаграмма' then
     case TOutputForm(ActiveMDIChild).PageControl1.ActivePageIndex of
      0: TOutputForm(ActiveMDIChild).Chart1.Print;
      1: TOutputForm(ActiveMDIChild).Chart2.Print;
      2: TOutputForm(ActiveMDIChild).Chart3.Print;
      3: TOutputForm(ActiveMDIChild).Chart4.Print;
      4: TOutputForm(ActiveMDIChild).Chart5.Print;
     end;
  end;
  end; }
end;

procedure TMainForm.FileExit1Execute(Sender: TObject);
begin
  Close;
end;


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

function SortMicro(p1,p2:pointer):integer;
begin
   Result:=Sign(PMonth(p1).Micro-PMonth(p2).Micro);
end;

function SortMicroList(p1,p2:pointer):integer;
begin
   Result:=Sign(PMicro(p1).M-PMicro(p2).M);
end;

function SortYear(p1,p2:pointer):integer;
begin
   Result:=Sign(PMonth(p1).Year-PMonth(p2).Year);
end;

function SortMonth(p1,p2:pointer):integer;
begin
   Result:=Sign(PMonth(p1).Name-PMonth(p2).Name);
end;


procedure TMainForm.GenerateData;
var
  allrelax, relaxvolume, dtcount, len, i,j,jj : integer;
  dt : TDateTime;
  dl : TDataLeaf;
  myfound, execfound2, fnd, datafound, execfound, ofpfound, manfound : boolean;
  control, volume1, volume2, summ_ofpspeed, summ_ofpvolume, ofpspeed, ofpvolume, summ_speed, summ_volume, speed, volume : real;
//  M : TMemoryStream;
  h : tHandle;
  ActivateExcel : function : Integer;
  FirstWeekDay: Integer;
  FirstWeekDate: Integer;
  OutputForm : TOutputForm;
  HTMLForm : THtmlForm;
  S2P, S3P, SUpr1, SUpr2, SRC, STControl, S0, S1 , S2, S3, S4, S5, S6, SR, STRelax, SRA : TStringList;
  CurMonth, SWeek, SMonth, SYear, SDay : Word;
  M0 : TMemoryStream;
  SportMens, ss : string;
  LI : TListItem;
  ColUpr1, ColUpr2, ColUprOFP : Integer;
  AllSR, AllSRA, speed2, summ_speed2, summ_volume2 : real;
  VPP, IPP, VSP, ISP : real;
  st : TStringList;
  MicroList, SubMonthList, MonthList : TList;

procedure GridRemoveColumn(StrGrid: TStringGrid; DelColumn: Integer);
var
  Column: Integer;
begin
  if DelColumn <= StrGrid.ColCount then
  begin
    for Column := DelColumn to StrGrid.ColCount - 1 do
      StrGrid.Cols[Column - 1].Assign(StrGrid.Cols[Column]);
    StrGrid.ColCount := StrGrid.ColCount - 1;
  end;
end;

procedure GridAddColumn(StrGrid: TStringGrid; NewColumn: Integer; S : string);
var
  Column: Integer;
begin
  StrGrid.ColCount := StrGrid.ColCount + 1;
  for Column := StrGrid.ColCount - 1 downto NewColumn do
    StrGrid.Cols[Column].Assign(StrGrid.Cols[Column - 1]);
  StrGrid.Cols[NewColumn - 1].Text := S;
end;

procedure DateToWeek(ADate: TDateTime; var ADay, AWeek, AMonth, AYear: Word);
begin
   ADate := ADate - ((DayOfWeek(ADate) - FirstWeekDay + 7) mod 7) + 7 - FirstWeekDate;
   DecodeDate(ADate, AYear, AMonth, ADay);
   AWeek := (Trunc(ADate - EncodeDate(AYear, 1, 1)) div 7) + 1;
end;

begin

  ColUpr1 :=0;
  ColUpr2 :=0;
  ColUprOFP :=0;

  FirstWeekDay:=2;
  FirstWeekDate:=4;

  S0 := TStringList.Create;
  S1 := TStringList.Create;
  S2 := TStringList.Create;
  S3 := TStringList.Create;
  S4 := TStringList.Create;
  S5 := TStringList.Create;
  S6 := TStringList.Create;
  SR := TStringList.Create;
  SRA := TStringList.Create;
  SRC := TStringList.Create;
  SUpr1 := TStringList.Create;
  SUpr2 := TStringList.Create;

  S2P := TStringList.Create;
  S3P := TStringList.Create;
  MonthList := TList.Create;

  RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'relax.mtx';
  STRelax := TStringList.Create;
  RazrDlg.GetRazr(STRelax);

  RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'control.mtx';
  STControl := TStringList.Create;
  RazrDlg.GetRazr(STControl);

  S4.Add('Соревновательные');
  S4.Add('Вспомогательные');
  S4.Add('Общеразвивающие');
  S5.Add('0');
  S5.Add('0');
  S5.Add('0');

  for I := 0 to DiagForm.Memo2.Lines.Count - 1 do
   SUpr1.Add('0');
  for I := 0 to DiagForm.Memo3.Lines.Count - 1 do
   SUpr2.Add('0');

  len := 0;

  {  for I := 0 to TMDIChild(ActiveMDIChild).BigDataLeaf.Count - 1 do
   if TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i]).menindex<>-1 then
    begin
     dl := TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i]);
     DateToWeek(dl.dt,SWeek, SMonth, SYear);

     manfound := false;
     for j := 0 to DiagForm.ListView2.Items.Count - 1 do
     if dl.men = DiagForm.ListView2.Items[j].Subitems[0] then
      begin
        manfound := true;
        break;
      end;

     if DiagForm.RadioButton5.Checked then // по дням
     begin
      if manfound then
       if DiagForm.RadioButton6.Checked then // Весь период
        inc(len)
       else
       if (dl.dt >= DiagForm.DateTimePicker1.Date-1) and
       (dl.dt <= DiagForm.DateTimePicker2.Date) then
        inc(len);
     end
     else

     if DiagForm.RadioButton3.Checked then // по месяцам
     begin
      if manfound then
       if DiagForm.RadioButton6.Checked then // Весь период
        if i>0 then
          begin
           CurMonth := SMonth;
           DateToWeek(TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i-1]).dt, SWeek, SMonth, SYear);
           if CurMonth <> SMonth then
            inc(len);
          end
       else
       if (dl.dt >= DiagForm.DateTimePicker1.Date-1) and
       (dl.dt <= DiagForm.DateTimePicker2.Date) then
        if i>0 then
          begin
           CurMonth := SMonth;
           DateToWeek(TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i-1]).dt, SWeek, SMonth, SYear);
           if CurMonth <> SMonth then
            inc(len);
          end;
     end;


    end; }


  summ_speed := 0;
  summ_volume := 0;

  summ_volume2 := 0;
  summ_speed2 := 0;

  summ_ofpspeed := 0;
  summ_ofpvolume := 0;
  dtcount := 0;


  for I := 0 to TMDIChild(ActiveMDIChild).BigDataLeaf.Count - 1 do
   if TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i]).menindex<>-1 then
   begin

    speed := 0;
    volume := 0;
    volume1 := 0;
    volume2 := 0;
    speed2 := 0;
    ofpspeed := 0;
    ofpvolume := 0;
    relaxvolume := 0;
    allrelax := 0;
    control := 0;

    dl := TDataLeaf(TMDIChild(ActiveMDIChild).BigDataLeaf.Items[i]);

    manfound := false;
    for j := 0 to DiagForm.ListView2.Items.Count - 1 do
     if dl.men = DiagForm.ListView2.Items[j].Subitems[0] then
      begin
        manfound := true;
        break;
      end;

    // Восстановление +++++++++++++++++++++++++++++++++++++++++++
//    if DiagForm.CheckBox6.Checked then
//    begin
 //    if manfound then
{     if DiagForm.RadioButton6.Checked or ((dl.dt >= DiagForm.DateTimePicker1.Date-1) and
     (dl.dt <= DiagForm.DateTimePicker2.Date))then
     begin
      allrelax := STRelax.Count;
      for jj := 0 to STRelax.Count - 1 do
       for j := 0 to dl.RelaxList.Count - 1 do
       begin
        if dl.RelaxList[j] = StRelax[jj] then
         inc(relaxvolume);
       end;
     end; }
//    end;
    // Восстановление ------------------------------------------------

    // Контроль  +++++++++++++++++++++++++++++++++++++++++++
{    if DiagForm.CheckBox7.Checked then
    begin
     if manfound then
     if DiagForm.RadioButton6.Checked or ((dl.dt >= DiagForm.DateTimePicker1.Date-1) and
     (dl.dt <= DiagForm.DateTimePicker2.Date))then
       try
        if (dl.ControlList.Count > 0) and (DiagForm.ComboBox1.ItemIndex <> -1) then
         control := control + StrToFloat(dl.ControlList[DiagForm.ComboBox1.ItemIndex]);
       except

       end;
    end; }
    // Контроль ------------------------------------------------


    if manfound then
     if DiagForm.RadioButton6.Checked or (dl.mezo = DiagForm.MezoBox.ItemIndex+1) then
     begin
      DecodeDate(dl.dt, SYear, SMonth, SDay);
      myfound := false;
      for jj := 0 to MonthList.Count - 1 do
      begin
       tMonth := MonthList.Items[jj];
       if (tMonth^.Year = SYear) and (tMonth^.Name = SMonth) and (tMonth^.Micro = dl.micro) then
       begin
        myfound := true;
        break;
       end;
      end;
      if not myfound then
      begin
       New(tMonth);
       MonthList.Add(tMonth);
       tMonth^.Name := SMonth;
       tMonth^.Year := SYear;
       tMonth^.Micro := dl.micro;
       tMonth^.Days := TStringList.Create;
      end;
      if tMonth<>nil then
      begin
       myfound := false;
       for jj := 0 to tMonth^.Days.Count - 1 do
       begin
        if StrToInt(tMonth^.Days.Strings[jj]) = SDay then
        begin
         myfound := true;
         break;
        end;
       end;
       if not myfound then
        tMonth^.Days.Add(IntToStr(SDay));
      end;
      for j := 0 to dl.BigDataGrid.Count - 1 do
      if TDataContent(dl.BigDataGrid[j]).dt<>0 then
      begin

      DecodeDate(TDataContent(dl.BigDataGrid[j]).dt, SYear, SMonth, SDay);
      myfound := false;
      for jj := 0 to MonthList.Count - 1 do
      begin
       tMonth := MonthList.Items[jj];
       if (tMonth^.Year = SYear) and (tMonth^.Name = SMonth) and (tMonth^.Micro = dl.micro) then
       begin
        myfound := true;
        break;
       end;
      end;
      if not myfound then
      begin
       New(tMonth);
       MonthList.Add(tMonth);
       tMonth^.Name := SMonth;
       tMonth^.Year := SYear;
       tMonth^.Micro := dl.micro;
       tMonth^.Days := TStringList.Create;
      end;
      if tMonth<>nil then
      begin
       myfound := false;
       for jj := 0 to tMonth^.Days.Count - 1 do
       begin
        if StrToInt(tMonth^.Days.Strings[jj]) = SDay then
        begin
         myfound := true;
         break;
        end;
       end;
       if not myfound then
        tMonth^.Days.Add(IntToStr(SDay));
      end;
      end;

     end;

    if manfound then
    if DiagForm.RadioButton6.Checked or (dl.mezo = DiagForm.MezoBox.ItemIndex+1) then
    for j := 0 to dl.BigDataGrid.Count - 1 do
     begin

      execfound := false;
      execfound2 := false;

      for jj := 0 to DiagForm.ListView4.Items.Count - 1 do
      if TDataContent(dl.BigDataGrid[j]).IsOFP = 0 then
       if TDataContent(dl.BigDataGrid[j]).Exec = DiagForm.ListView4.Items[jj].Subitems[0] then
       begin
        SUpr1[jj] := IntToStr(StrToInt(SUpr1[jj]) + 1);
        Inc(ColUpr1);
        execfound := true;
        break;
       end;

      for jj := 0 to DiagForm.ListView5.Items.Count - 1 do
      if TDataContent(dl.BigDataGrid[j]).IsOFP = 0 then
       if TDataContent(dl.BigDataGrid[j]).Exec = DiagForm.ListView5.Items[jj].Subitems[0] then
       begin
        SUpr2[jj] := IntToStr(StrToInt(SUpr2[jj]) + 1);
        Inc(ColUpr2);
        execfound2 := true;
        break;
       end;

      ofpfound := false;

      for jj := 0 to DiagForm.ListView6.Items.Count - 1 do
      if TDataContent(dl.BigDataGrid[j]).IsOFP = 1 then
       if TDataContent(dl.BigDataGrid[j]).OFP = DiagForm.ListView6.Items[jj].Subitems[0] then
       begin
        Inc(ColUprOFP);
        ofpfound := true;
        break;
       end;

      if execfound or execfound2 then
      begin
       for jj := 0 to TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Count - 1 do
         try
          if execfound then
           volume1 := volume1 +
           StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res2)*
           StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res3);

          if execfound2 then
          begin
           volume2 := volume2 +
           StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res2)*
           StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res3);

           speed2 := speed2 +
           StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res1)*
           StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res2)*
           StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res3);
          end;

          volume := volume +
          StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res2)*
          StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res3);

          speed := speed +
          StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res1)*
          StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res2)*
          StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res3);

         except
         end;
//       if volume>0 then
//       speed := speed / volume;
      end;

      if ofpfound then
      begin
       for jj := 0 to TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Count - 1 do
         try
          ofpvolume := ofpvolume +
          StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res2);

          ofpspeed := ofpspeed +
          StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Items[jj]).Res1);
         except
         end;
//       if TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Count>0 then
//        ofpspeed := ofpspeed / TDataContent(dl.BigDataGrid.Items[j]).DataGrid.Count;
      end;

     end;

    if manfound then
    if DiagForm.RadioButton6.Checked or (dl.mezo = DiagForm.MezoBox.ItemIndex+1) then
    begin


     S5[0] := FloatToStr(StrToFloat(S5[0]) + volume1);
     S5[1] := FloatToStr(StrToFloat(S5[1]) + volume2);
     S5[2] := FloatToStr(StrToFloat(S5[2]) + ofpvolume);

     summ_speed := summ_speed + speed;
     summ_volume := summ_volume + volume;
     summ_volume2 := summ_volume2 + volume2;
     summ_speed2 := summ_speed2 + speed2;

     summ_ofpspeed := summ_ofpspeed + ofpspeed;
     summ_ofpvolume := summ_ofpvolume + ofpvolume;

     inc(dtcount);

     datafound := false;

//    if DiagForm.RadioButton5.Checked then // по дням ***********************************************
//     begin

      // Распределение мо микроциклам

      for jj := 0 to S1.Count - 1 do
       if StrToInt(S1[jj]) = dl.micro then
       begin
        datafound := true;
        break;
       end;

      if datafound then
      begin
       S2[jj] := FloatToStr(StrToFloat(S2[jj]) + volume);
       S6[jj] := FloatToStr(StrToFloat(S6[jj]) + ofpvolume);
       SR[jj] := FloatToStr(StrToFloat(SR[jj]) + relaxvolume);
       SRA[jj] := FloatToStr(StrToFloat(SRA[jj]) + allrelax);
       SRC[jj] := FloatToStr(StrToFloat(SRC[jj]) + control);
       //if (volume <> 0) then
       S3[jj] := FloatToStr(StrToFloat(S3[jj]) + speed);
      end
      else
      begin
//      if (volume <> 0) or (speed<>0)  then
//       begin
        S1.Add(IntToStr(dl.micro));
        S2.Add(FloatToStr(volume));
        S6.Add(FloatToStr(ofpvolume));
        SR.Add(FloatToStr(relaxvolume));
        SRA.Add(FloatToStr(allrelax));
        SRC.Add(FloatToStr(control));
       // if (volume <> 0) then
         S3.Add(FloatToStr(speed));
       // else
       //  S3.Add('0');
//       end;
      end;


    end;
   end;

  VPP := 0;
  IPP := 0;
  VSP := 0;
  ISP := 0;
  // Вычислим проценты

  for jj := 0 to S1.Count - 1 do
  begin
    if StrToInt(S1[jj]) <= 4 then
    // Подготовительный период
    begin
     VPP := VPP + StrToFloat(S2[jj]);
     if StrToFloat(S2[jj]) <> 0 then
      IPP := IPP + StrToFloat(S3[jj])/StrToFloat(S2[jj]);
    end
    else
    // Соревновательный период
    if StrToInt(S1[jj]) > 4 then
    begin
     VSP := VSP + StrToFloat(S2[jj]);
     if StrToFloat(S2[jj]) <> 0 then
      ISP := ISP + StrToFloat(S3[jj])/StrToFloat(S2[jj]);
    end
  end;

  for jj := 0 to S1.Count - 1 do
  begin
    if StrToInt(S1[jj]) <= 4 then
     S2P.Add(FloatToStr(StrToFloat(S2[jj])/VPP*100))
    else
     S2P.Add(FloatToStr(StrToFloat(S2[jj])/VSP*100));
    if StrToInt(S1[jj]) <= 4 then
     S3P.Add(FloatToStr((StrToFloat(S3[jj])/StrToFloat(S2[jj]))/IPP*100))
    else
     S3P.Add(FloatToStr((StrToFloat(S3[jj])/StrToFloat(S2[jj]))/ISP*100));
  end;

  MicroList := TList.Create;
  for jj := 0 to S1.Count - 1 do
  begin
    New(tMicro);
    MicroList.Add(tMicro);
    tMicro^.M := StrToInt(S1[jj]);
    tMicro^.V := StrToFloat(S2[jj]);
    tMicro^.I := StrToFloat(S3[jj])/StrToFloat(S2[jj]);
    tMicro^.VP := StrToFloat(S2P[jj]);
    tMicro^.IP := StrToFloat(S3P[jj]);
  end;
  MicroList.Sort(SortMicroList);

  SportMens := '';
  for I := 0 to DiagForm.Memo1.Lines.Count - 1 do
   if i = DiagForm.Memo1.Lines.Count - 1 then
    SportMens := SportMens + DiagForm.Memo1.Lines[i]
   else
    SportMens := SportMens + DiagForm.Memo1.Lines[i] + ', ';

//  OutputForm := TOutputForm.Create(Application);
 HTMLForm := THTMLForm.Create(Application);
 st := TStringList.Create;
 st.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">');
 st.Add('<html>');
 st.Add('<head>');
 st.Add('<title></title>');
 st.Add('<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">');
 st.Add('<title>Отчет '+IntToStr(DiagForm.MezoBox.ItemIndex+1)+' мезоцикл</title>');
 st.Add('<link rel="stylesheet" href="js/style.css" type="text/css" media="screen" />');
 st.Add('<script type="text/javascript" src="js/jquery.js"></script>');
 st.Add('<script type="text/javascript" src="js/jquery.calendar-widget.js"></script>');
 st.Add('<!--[if lt IE 9]><script language="javascript" type="text/javascript" src="js/excanvas.js"></script><![endif]-->');
 st.Add('<script type="text/javascript" src="js/jquery.jqplot.js"></script>');
 st.Add('<script type="text/javascript" src="js/plugins/jqplot.logAxisRenderer.min.js"></script>');
 st.Add('<script type="text/javascript" src="js/plugins/jqplot.canvasTextRenderer.min.js"></script>');
 st.Add('<script type="text/javascript" src="js/plugins/jqplot.canvasAxisLabelRenderer.min.js"></script>');
 st.Add('<script type="text/javascript" src="js/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>');
 st.Add('<script type="text/javascript" src="js/plugins/jqplot.categoryAxisRenderer.min.js"></script>');
 st.Add('<link rel="stylesheet" type="text/css" href="js/jquery.jqplot.css" />');

 st.Add('<script type="text/javascript">');
 st.Add('$(function()');
 st.Add('{');

 MonthList.Sort(SortMicro);

 for j := 0 to MonthList.Count - 1 do
  begin
   tMonth := MonthList.Items[j];
   ss := '';
   for jj := 0 to tMonth^.Days.Count - 1 do
    ss := ss + tMonth^.Days.Strings[jj]+',';
   st.Add('	$("#month'+IntToStr(j+1)+'").calendarWidget({month: '+IntToStr(tMonth^.Name-1)+
   ',year: '+IntToStr(tMonth^.Year)+', micro: '+IntToStr(tMonth^.Micro)+', sportdays:['+ss+']});');
  end;

 st.Add('});');
 st.Add('</script>');

 st.Add('<script type="text/javascript">');
 st.Add('$(document).ready(function(){');

 st.Add('$.jqplot('+chr($27)+'chartdiv'+chr($27)+
 ',  [[['+chr($27)+'Подготовительный'+chr($27)+', '+FormatFloat('0.0',(VPP/(VSP+VPP))*100)+
 '],['+chr($27)+'Соревновательный'+chr($27)+','+FormatFloat('0.0',(VSP/(VSP+VPP))*100)+
 ']],[['+chr($27)+'Подготовительный'+chr($27)+', '+FormatFloat('0.0',(IPP/(ISP+IPP))*100)+
 '],['+chr($27)+'Соревновательный'+chr($27)+','+FormatFloat('0.0',(ISP/(ISP+IPP))*100)+']]],');
 st.Add('{ title:'+chr($27)+'Соотношение нагрузки по периодам'+chr($27)+',');
 st.Add(' seriesDefaults:{pointLabels: {show: true}}, series:[ {label:'+
 chr($27)+'Экстенсивность (объем) (%)'+chr($27)+'}, {label:'+chr($27)+'Интенсивность (%)'+chr($27)+
 '}, {yaxis:'+chr($27)+'y2axis'+chr($27)+'},{color:'+chr($27)+'#5FAB78'+chr($27)+'}], axes:{xaxis:{renderer: $.jqplot.CategoryAxisRenderer, label: '+
 chr($27)+'Периоды'+chr($27)+
 ', labelRenderer: $.jqplot.CanvasAxisLabelRenderer, tickRenderer: $.jqplot.CanvasAxisTickRenderer}, '+
 'yaxis:{ autoscale:true, label: '+chr($27)+'Объем'+chr($27)+', labelRenderer: $.jqplot.CanvasAxisLabelRenderer, tickRenderer: $.jqplot.CanvasAxisTickRenderer, min:0, max:100},'+
 'y2axis:{ autoscale:true, label: '+chr($27)+'Интенсивность'+chr($27)+', labelRenderer: $.jqplot.CanvasAxisLabelRenderer, tickRenderer: $.jqplot.CanvasAxisTickRenderer, min:0, max:100}},');
 st.Add(' legend: {show: true, location: '+chr($27)+'e'+chr($27)+', placement: '+chr($27)+'outside'+chr($27)+'}');

 st.Add('});');
 st.Add('});');
 st.Add('</script>');

 st.Add('<script type="text/javascript">');
 st.Add('$(document).ready(function(){');

 ss:= ',  [[';
 for j := 0 to MicroList.Count - 2 do
  begin
   tMicro := MicroList.Items[j];
   ss := ss + '[' + chr($27)+ IntToStr(tMicro^.M)+' микроцикл'+chr($27)+', '+FormatFloat('0.0',tMicro^.VP) + '],';
  end;
 tMicro := MicroList.Items[MicroList.Count - 1];
 ss := ss + '[' + chr($27)+ IntToStr(tMicro^.M)+' микроцикл'+chr($27)+', '+FormatFloat('0.0',tMicro^.VP) + ']';
 ss := ss + '],[';
 for j := 0 to MicroList.Count - 2 do
  begin
   tMicro := MicroList.Items[j];
   ss := ss + '[' + chr($27)+ IntToStr(tMicro^.M)+' микроцикл'+chr($27)+', '+FormatFloat('0.0',tMicro^.IP) + '],';
  end;
 tMicro := MicroList.Items[MicroList.Count - 1];
 ss := ss + '[' + chr($27)+ IntToStr(tMicro^.M)+' микроцикл'+chr($27)+', '+FormatFloat('0.0',tMicro^.IP) + ']';

 st.Add('$.jqplot('+chr($27)+'chartdiv2'+chr($27)+ss+']],');
 st.Add('{ title:'+chr($27)+'Объем и интенсивность за Мезоцикл'+chr($27)+',');
 st.Add(' seriesDefaults:{pointLabels: {show: true}}, series:[ {label:'+
 chr($27)+'Экстенсивность (объем) (%)'+chr($27)+'}, {label:'+chr($27)+'Интенсивность (%)'+chr($27)+
 '}, {yaxis:'+chr($27)+'y2axis'+chr($27)+'},{color:'+chr($27)+'#5FAB78'+chr($27)+'}], axes:{xaxis:{renderer: $.jqplot.CategoryAxisRenderer, label: '+
 chr($27)+'Периоды'+chr($27)+
 ', labelRenderer: $.jqplot.CanvasAxisLabelRenderer, tickRenderer: $.jqplot.CanvasAxisTickRenderer}, '+
 'yaxis:{ autoscale:true, label: '+chr($27)+'Объем'+chr($27)+', labelRenderer: $.jqplot.CanvasAxisLabelRenderer, tickRenderer: $.jqplot.CanvasAxisTickRenderer, min:0, max:100},'+
 'y2axis:{ autoscale:true, label: '+chr($27)+'Интенсивность'+chr($27)+', labelRenderer: $.jqplot.CanvasAxisLabelRenderer, tickRenderer: $.jqplot.CanvasAxisTickRenderer, min:0, max:100}},');
 st.Add(' legend: {show: true, location: '+chr($27)+'e'+chr($27)+', placement: '+chr($27)+'outside'+chr($27)+'}');

 st.Add('});');
 st.Add('});');
 st.Add('</script>');

 st.Add('</head>');

 st.Add('<body>');
 st.Add('<p>'+SportMens+'</p>');
 st.Add('<p>1. График тренировок ('+IntToStr(DiagForm.MezoBox.ItemIndex+1)+' мезоцикл)</p>');

 st.Add('<p><table>');
 for j := 1 to 8 do
   st.Add('<tr><td class="sportday'+IntTostr(j)+'"></td><td>'+IntTostr(j)+' микроцикл</td></tr>');
 st.Add('</table></p>');

 st.Add('<table><tr>');
 for j := 0 to MonthList.Count - 1 do
  begin
   if j mod 4 = 0 then
    st.Add('</tr><tr>');
   st.Add('<td><div id="month'+IntToStr(j+1)+'">');
   st.Add('<p>Включите Javascript для просмотра календаря.</p>');
   st.Add('</div></td>');
  end;
 st.Add('</tr></table>');

 st.Add('<p>2. Величины нагрузки</p>');

 st.Add('<table><tr><td>');
 st.Add('<table>');
 st.Add('<tr><td><table><tr><td>Соотношение</td></tr></table></td></tr>');
 st.Add('<tr><td><table border="1" cellspacing="3" cellpadding="3"><tr><td align="center">V</td><td align="center">'+FormatFloat('0.0',VSP+VPP)+'</td></tr>');
 st.Add('<tr><td align="center">I</td><td align="center">'+FormatFloat('0.0',ISP+IPP)+'</td></tr></table></td></tr>');
 st.Add('<tr><td><table border="1" cellspacing="3" cellpadding="3"><tr><td align="center">V (кпш)</td><td align="center">'+FormatFloat('0.0',VPP)+
 '</td><td align="center">'+FormatFloat('0.0',VSP)+'</td></tr>');
 st.Add('<tr><td align="center">I (кг)</td><td align="center">'+FormatFloat('0.0',IPP)+
 '</td><td align="center">'+FormatFloat('0.0',ISP)+'</td></tr>');
 st.Add('<tr><td align="center">V (%)</td><td align="center">'+FormatFloat('0.0',(VPP/(VSP+VPP))*100)+
 '</td><td align="center">'+FormatFloat('0.0',(VSP/(VSP+VPP))*100)+'</td></tr>');
 st.Add('<tr><td align="center">I (%)</td><td align="center">'+FormatFloat('0.0',(IPP/(ISP+IPP))*100)+
 '</td><td align="center">'+FormatFloat('0.0',(ISP/(ISP+IPP))*100)+'</td></tr></table></td>');
 st.Add('</tr></table></td>');

 st.Add('<td><p><div id='+chr($27)+'chartdiv'+chr($27)+' style='+chr($27)+'height:300px;width:600px; '+chr($27)+
 '></div></p></td></tr></table>');

 st.Add('<p>3. Варианты распределения нагрузки</p>');

 st.Add('Варианты распределения нагрузки (кпш, кг)');
 st.Add('<table border="1" cellspacing="3" cellpadding="3"><tr>');
 st.Add('<td align="center"></td>');
 for j := 0 to MicroList.Count - 1 do
  begin
   tMicro := MicroList.Items[j];
   st.Add('<td align="center">'+IntToStr(tMicro^.M)+' микроцикл</td>');
  end;
 st.Add('</tr><tr><td align="center">V</td>');
 for j := 0 to MicroList.Count - 1 do
  begin
   tMicro := MicroList.Items[j];
   st.Add('<td align="center">'+FormatFloat('0.0',tMicro^.V)+'</td>');
  end;
 st.Add('</tr><tr><td align="center">I</td>');
 for j := 0 to MicroList.Count - 1 do
  begin
   tMicro := MicroList.Items[j];
   st.Add('<td align="center">'+FormatFloat('0.0',tMicro^.I)+'</td>');
  end;

 st.Add('</tr></table>');
 st.Add('Варианты распределения нагрузки (%)');
 st.Add('<table border="1" cellspacing="3" cellpadding="3"><tr>');

 st.Add('<td align="center"></td>');
 for j := 0 to MicroList.Count - 1 do
  begin
   tMicro := MicroList.Items[j];
   st.Add('<td align="center">'+IntToStr(tMicro^.M)+' микроцикл</td>');
  end;
 st.Add('</tr><tr><td align="center">V</td>');
 for j := 0 to MicroList.Count - 1 do
  begin
   tMicro := MicroList.Items[j];
   st.Add('<td align="center">'+FormatFloat('0.0',tMicro^.VP)+'</td>');
  end;
 st.Add('</tr><tr><td align="center">I</td>');
 for j := 0 to MicroList.Count - 1 do
  begin
   tMicro := MicroList.Items[j];
   st.Add('<td align="center">'+FormatFloat('0.0',tMicro^.IP)+'</td>');
  end;
 st.Add('</tr></table>');

 st.Add('<p><div id='+chr($27)+'chartdiv2'+chr($27)+' style='+chr($27)+'height:300px;width:700px; '+chr($27)+
 '></div></p>');

 {  fnd := false;
  len := 0;
  for I := DockTabSet1.Tabs.Count - 1 downto 0 do
   if copy(DockTabSet1.Tabs[i],1,9) = 'Диаграмма' then
    begin
     fnd := true;
     if StrToInt(copy(DockTabSet1.Tabs[i],11,Pos(':',DockTabSet1.Tabs[i])-11))>len then
       len := StrToInt(copy(DockTabSet1.Tabs[i],11,Pos(':',DockTabSet1.Tabs[i])-11));
    end;

  if not fnd then
  begin
   if DiagForm.RadioButton6.Checked then // Весь период
    OutputForm.Caption := 'Диаграмма 1: Весь период'
   else
    OutputForm.Caption := 'Диаграмма 1: Период ' + DiagForm.MezoBox.Text;
  end
  else
  begin
   if DiagForm.RadioButton6.Checked then // Весь период
    OutputForm.Caption := 'Диаграмма '+IntToStr(len+1)+': Весь период'
   else
    OutputForm.Caption := 'Диаграмма '+IntToStr(len+1)+': Период ' + DiagForm.MezoBox.Text;
  end; }

//  OutputForm.PageControl1.ActivePageIndex := 0;

//  OutputForm.LV.Items.Clear;

//  LI := OutputForm.LV.Items.Add;
//  LI.Caption := 'Спортсмены';
{  LI.SubItems.Add(ss);

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Упраженения';
  ss := '';
  for I := 0 to DiagForm.Memo2.Lines.Count - 1 do
   if i = DiagForm.Memo2.Lines.Count - 1 then
    ss := ss + DiagForm.Memo2.Lines[i]
   else
    ss := ss + DiagForm.Memo2.Lines[i] + ', ';
  for I := 0 to DiagForm.Memo3.Lines.Count - 1 do
   if i = DiagForm.Memo3.Lines.Count - 1 then
    ss := ss + DiagForm.Memo3.Lines[i]
   else
    ss := ss + DiagForm.Memo3.Lines[i] + ', ';
  LI.SubItems.Add(ss);

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Упраженения ОФП';
  ss := '';
  for I := 0 to DiagForm.Memo4.Lines.Count - 1 do
   if i = DiagForm.Memo4.Lines.Count - 1 then
    ss := ss + DiagForm.Memo4.Lines[i]
   else
    ss := ss + DiagForm.Memo4.Lines[i] + ', ';
  LI.SubItems.Add(ss);

  if DiagForm.RadioButton6.Checked then // Весь период
   begin
    LI := OutputForm.LV.Items.Add;
    LI.Caption := 'Период анализа';
    LI.SubItems.Add('Весь период');
   end
  else
   begin
    LI := OutputForm.LV.Items.Add;
    LI.Caption := 'Период анализа';
    LI.SubItems.Add('Период ' + DiagForm.MezoBox.Text);
   end;

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Объем всех упражнений СФП (Вариант 1)';
  LI.SubItems.Add(FormatFloat('0',summ_volume));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Объем всех упражнений СФП (Вариант 2)';
  LI.SubItems.Add(FormatFloat('0',summ_speed));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Интенсивность всех упражнений (в единицах)';
  if summ_volume<>0 then LI.SubItems.Add(FormatFloat('0.0',summ_speed/summ_volume));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Объем ОФП';
  LI.SubItems.Add(FormatFloat('0.0',summ_ofpvolume));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Интенсивность ОФП';
  LI.SubItems.Add(FormatFloat('0.0',summ_ofpspeed));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Частота тренировок по СФП';
  LI.SubItems.Add(FormatFloat('0',ColUpr1 + ColUpr2));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Частота тренировок по ОФП';
  LI.SubItems.Add(FormatFloat('0',ColUprOFP));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Частота тренировок по СФП и ОФП в сумме';
  LI.SubItems.Add(FormatFloat('0',ColUpr1 + ColUpr2 + ColUprOFP));

  for I := 0 to DiagForm.Memo2.Lines.Count - 1 do
  begin
   LI := OutputForm.LV.Items.Add;
   LI.Caption := 'Частота упражнения "' + DiagForm.Memo2.Lines[i] + '"';
   LI.SubItems.Add(Supr1[i]);
  end;

  for I := 0 to DiagForm.Memo3.Lines.Count - 1 do
  begin
   LI := OutputForm.LV.Items.Add;
   LI.Caption := 'Частота упражнения "' + DiagForm.Memo3.Lines[i] + '"';
   LI.SubItems.Add(Supr2[i]);
  end;

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Объем вспомогательных упражнений (Вариант 1)';
  LI.SubItems.Add(FormatFloat('0',summ_volume2));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Объем вспомогательных упражнений (Вариант 2)';
  LI.SubItems.Add(FormatFloat('0',summ_speed2));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Интенсивность вспомогательных упражнений (в единицах)';
  if summ_volume2<>0 then LI.SubItems.Add(FormatFloat('0.0',summ_speed2/summ_volume2));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Интенсивность вспомогательных упражнений (в %)';
  LI.SubItems.Add(FormatFloat('0',0));

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Соотношение СФП и ОФП (в %)';
  LI.SubItems.Add(FormatFloat('0.0',summ_volume / (summ_volume + summ_ofpvolume) * 100) + '% к ' +
  FormatFloat('0.0',summ_ofpvolume / (summ_volume + summ_ofpvolume) * 100) + '%');

  LI := OutputForm.LV.Items.Add;
  LI.Caption := 'Соотношение соревновательных и  вспомогательных упражнений (в %)';
  LI.SubItems.Add(FormatFloat('0.0',summ_volume / (summ_volume + summ_volume2) * 100) + '% к ' +
  FormatFloat('0.0',summ_volume2 / (summ_volume + summ_volume2) * 100) + '%');   }

  DockTabSet1.Visible := true;
  HTMLForm.Caption := 'Отчет ' + SportMens + ' ' + IntToStr(DiagForm.MezoBox.ItemIndex+1) + ' мезоцикл';
  DockTabSet1.Tabs.Add(HTMLForm.Caption+'       ');

  DockTabSet1.TabIndex := DockTabSet1.Tabs.Count-1;

//  OutputForm.Series1.Clear;
//  OutputForm.Series2.Clear;
//  OutputForm.Series3.Clear;
//  OutputForm.LineSeries1.Clear;
//  OutputForm.LineSeries2.Clear;
//  OutputForm.LineSeries3.Clear;

{  with OutputForm.StringGrid1 do
   begin
     for i := 0 to ColCount - 1 do
       for j := 0 to RowCount - 1 do
         Cells[i, j] := '';
   end;

  OutputForm.StringGrid1.Cols[0].Add('Микроцикл');
  OutputForm.StringGrid1.Cols[1].Add('Объем V');
  OutputForm.StringGrid1.Cols[2].Add('Интенсивность I');
  OutputForm.StringGrid1.Cols[3].Add('Объем ОФП');
  OutputForm.StringGrid1.Cols[4].Add('Объем V (%)');
  OutputForm.StringGrid1.Cols[5].Add('Интенсивность I (%)'); }


  // Заполняем первую диаграмму

{  OutputForm.Series1.AddXY(0,(VPP/(VSP+VPP))*100,'Подготовительный');
  OutputForm.Series2.AddXY(0,(IPP/(ISP+IPP))*100,'Подготовительный');
  OutputForm.Series1.AddXY(1,(VSP/(VSP+VPP))*100,'Соревновательный');
  OutputForm.Series2.AddXY(1,(ISP/(ISP+IPP))*100,'Соревновательный');

  LI := OutputForm.LV2.Items.Add;
  LI.Caption := 'Объем V';
  LI.SubItems.Add(FormatFloat('0.0',VSP+VPP));
  LI := OutputForm.LV2.Items.Add;
  LI.Caption := 'Интенсивность I';
  LI.SubItems.Add(FormatFloat('0.0',ISP+IPP));

  LI := OutputForm.LV2.Items.Add;
  LI.Caption := 'Объем в подготовительный период';
  LI.SubItems.Add(FormatFloat('0.0',VPP));
  LI := OutputForm.LV2.Items.Add;
  LI.Caption := 'Интенсивность в подготовительный период';
  LI.SubItems.Add(FormatFloat('0.0',IPP));
  LI := OutputForm.LV2.Items.Add;
  LI.Caption := 'Объем (%) в подготовительный период';
  LI.SubItems.Add(FormatFloat('0.0',(VPP/(VSP+VPP))*100));
  LI := OutputForm.LV2.Items.Add;
  LI.Caption := 'Интенсивность (%) в подготовительный период';
  LI.SubItems.Add(FormatFloat('0.0',(IPP/(ISP+IPP))*100));

  LI := OutputForm.LV2.Items.Add;
  LI.Caption := 'Объем в соревновательный период';
  LI.SubItems.Add(FormatFloat('0.0',VSP));
  LI := OutputForm.LV2.Items.Add;
  LI.Caption := 'Интенсивность в соревновательный период';
  LI.SubItems.Add(FormatFloat('0.0',ISP));
  LI := OutputForm.LV2.Items.Add;
  LI.Caption := 'Объем (%) в соревновательный период';
  LI.SubItems.Add(FormatFloat('0.0',(VSP/(VSP+VPP))*100));
  LI := OutputForm.LV2.Items.Add;
  LI.Caption := 'Интенсивность (%) в соревновательный период';
  LI.SubItems.Add(FormatFloat('0.0',(ISP/(ISP+IPP))*100));

  OutputForm.StringGrid1.RowCount := S1.Count + 1;

  for jj := 0 to S1.Count - 1 do
  begin
    // Заполняем вторую диаграмму
    OutputForm.LineSeries1.AddXY(StrToInt(S1[jj]),StrToFloat(S2P[jj]),S1[jj]);
    OutputForm.LineSeries2.AddXY(StrToInt(S1[jj]),StrToFloat(S3P[jj]),S1[jj]);

    // Остальное
    OutputForm.LineSeries4.AddXY(StrToInt(S1[jj]),StrToFloat(S2[jj]),S1[jj]);
    OutputForm.Series3.AddXY(StrToInt(S1[jj]),StrToFloat(S6[jj]),S1[jj]);
    OutputForm.BarSeries1.AddXY(StrToInt(S1[jj]),StrToFloat(SRC[jj]),S1[jj]);

      //      OutputForm.AreaSeries1.AddXY(StrToDate(S1[jj]),StrToFloat(SR[jj])/StrToFloat(SRA[jj])*100,S1[jj]);

    // Заполним таблицу
    OutputForm.StringGrid1.Cells[0,jj+1] := S1[jj];
    OutputForm.StringGrid1.Cells[1,jj+1] := FormatFloat('0',StrToFloat(S2[jj]));
    if StrToFloat(S2[jj])<>0 then
     OutputForm.StringGrid1.Cells[2,jj+1] := FormatFloat('0.0',StrToFloat(S3[jj])/StrToFloat(S2[jj]));
    OutputForm.StringGrid1.Cells[3,jj+1] := FormatFloat('0.0',OutputForm.Series3.YValues[jj]);

    OutputForm.StringGrid1.Cells[4,jj+1] := FormatFloat('0.0',StrToFloat(S2P[jj]));
    OutputForm.StringGrid1.Cells[5,jj+1] := FormatFloat('0.0',StrToFloat(S3P[jj]));
  end; }

{  With OutputForm.Series1 do
  begin
   XValues.Sort ;
   XValues.FillSequence ;
   Repaint;
  end;
  With OutputForm.Series2 do
  begin
   XValues.Sort ;
   XValues.FillSequence ;
   Repaint;
  end;
  With OutputForm.Series3 do
  begin
   XValues.Sort ;
   XValues.FillSequence ;
   Repaint;
  end;
  With OutputForm.LineSeries1 do
  begin
   XValues.Sort ;
   XValues.FillSequence ;
   Repaint;
  end;
  With OutputForm.LineSeries2 do
  begin
   XValues.Sort ;
   XValues.FillSequence ;
   Repaint;
  end;
  With OutputForm.LineSeries3 do
  begin
   XValues.Sort ;
   XValues.FillSequence ;
   Repaint;
  end;
  With OutputForm.LineSeries4 do
  begin
   XValues.Sort ;
   XValues.FillSequence ;
   Repaint;
  end;
  With OutputForm.AreaSeries1 do
  begin
   XValues.Sort ;
   XValues.FillSequence ;
   Repaint;
  end;
  With OutputForm.BarSeries1 do
  begin
   XValues.Sort ;
   XValues.FillSequence ;
   Repaint;
  end;

  for jj := 0 to S4.Count - 1 do
   case jj of
    0: OutputForm.LineSeries3.AddPie(StrToFloat(S5[jj]),S4[jj],clBlue);
    1: OutputForm.LineSeries3.AddPie(StrToFloat(S5[jj]),S4[jj],clRed);
    2: OutputForm.LineSeries3.AddPie(StrToFloat(S5[jj]),S4[jj],clGreen);
   end;  }

  S0.Free;
  S1.Free;
  S2.Free;
  S3.Free;

  S2P.Free;
  S3P.Free;

  S4.Free;
  S5.Free;
  S6.Free;
  SR.Free;
  SRA.Free;
  STRElax.Free;
  SRC.Free;
  STControl.Free;
  Supr1.Free;
  Supr2.Free;

  for jj := 0 to MonthList.Count - 1 do
   begin
    tMonth := MonthList.Items[jj];
    tMonth^.Days.Free;
    Dispose(tMonth);
   end;
  MonthList.Free;
  for jj := 0 to MicroList.Count - 1 do
   begin
    tMicro := MicroList.Items[jj];
    Dispose(tMicro);
   end;
  MicroList.Free;

{  OutputForm.TabSheet5.TabVisible := true;
  OutputForm.TabSheet6.TabVisible := true;

  OutputForm.Panel1.Visible := true;
  OutputForm.StringGrid1.Visible := true;
  OutputForm.PageControl1.Visible := true;

  if not OutputForm.PageControl1.Visible then
  begin
   if OutputForm.StringGrid1.Visible and OutputForm.Panel1.Visible then
   begin
    OutputForm.Panel1.Height := Trunc(OutputForm.ClientHeight / 2);
    OutputForm.StringGrid1.Height := Trunc(OutputForm.ClientHeight / 2);
   end
   else
   if not OutputForm.StringGrid1.Visible and OutputForm.Panel1.Visible then
   begin
    OutputForm.Panel1.Height := OutputForm.ClientHeight;
   end
   else
   if OutputForm.StringGrid1.Visible and not OutputForm.Panel1.Visible then
   begin
    OutputForm.StringGrid1.Height := OutputForm.ClientHeight;
   end;
  end;  }

 st.Add('</body>');
 st.Add('</html>');

 st.SaveToFile(ExtractFilePath(paramstr(0))+'report.html');
 st.Free;

 HTMLForm.wb.Navigate(ExtractFilePath(paramstr(0))+'report.html');
 with HTMLForm.wb do
    if Document <> nil then
     with Application as IOleobject do
       DoVerb(OLEIVERB_UIACTIVATE, nil, HTMLForm.wb, 0, Handle, GetClientRect);

{  h := LoadLibrary(PChar(ExtractFilePath(paramstr(0))+'acexcel.dll'));
  if h <= BadDllLoad then
     MessageBox(Application.Handle,PChar('Невозможно загрузить библиотеку acexcel.dll.'), PCHAR(MainTitle), MB_OK + MB_ICONWARNING)
  else
  begin
    @ActivateExcel := GetProcAddress(h,'ActivateExcel');
    if @ActivateExcel <> nil then
     ActivateExcel;
  end;
  FreeLibrary(h);   }

end;


end.
