unit CHILDWIN;

interface

uses Windows, Classes, Graphics, Forms, Controls, StdCtrls, Messages,
     SysUtils, ComCtrls, ExtCtrls, Contnrs, Buttons, ImgList, Mask, ActnList,
  Menus, Grids, ValEdit, ToolWin, ActnMan, ActnColorMaps, Tabs, DockTabSet,
  ButtonGroup;

type

  ManRec = record
    SportList : TStringList;
  end;
  pManRec = ^ManRec;

  TDataPoint = class
    Panel : TPanel;
    X,Y,Top : Integer;
    Edit1, Edit2, Edit3, Edit31 : TMemo;
    Label1 : TLabel;
    Res1, Res2, Res3 : string;
    constructor Create;
    procedure InitVisual(PosX : integer);
    destructor Destroy; override;
  end;

  TDataLeaf = class
    men, sport : string;
    mezo, micro, trcount, menindex, sportindex : integer;
    BigDataGrid: TObjectList;           // ñïèñîê
    RelaxList : TStringList;            // Восстановление
    ControlList : TStringList;          // Контроль
    saved : boolean;
    dt,dtl : TDateTime;
    constructor Create;
    destructor Destroy; override;
  end;

  TDataContent = class
    trainkind : integer;
    dt : TDateTime;
    ScrollBox : TScrollBox;
    Panel : TPanel;
    TrainDate : TDateTimePicker;
    TrainBox : TComboBox;
    TrainLabel, Label1 : TLabel;
    ComboBox1 : TComboBox;
    Exec, OFP, OFPEdit1, OFPEdit2 : string;
    trindex, execindex, ofpindex : integer;
    BitBtnNext, BitBtnPrev : TBitBtn;
    X,Y,Top, IsOFP : Integer;
    DataGrid: TObjectList;           // ñïèñîê
    CheckBox1 : TCheckBox;
    constructor Create;
    procedure  InitVisual(PosX : integer);
    destructor Destroy; override;
  end;

  TMDIChild = class(TForm)
    ImageList1: TImageList;
    ActionList1: TActionList;
    AddExec: TAction;
    DelExec: TAction;
    AddZone: TAction;
    DelZone: TAction;
    ImageList2: TImageList;
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    ComboBox1: TComboBox;
    Label2: TLabel;
    ComboBox2: TComboBox;
    Bevel1: TBevel;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    Label3: TLabel;
    ImageList3: TImageList;
    XPColorMap1: TXPColorMap;
    TreeView1: TTreeView;
    Splitter1: TSplitter;
    N2: TMenuItem;
    N3: TMenuItem;
    Panel4: TPanel;
    DockTabSet1: TDockTabSet;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    ScrollBox1: TScrollBox;
    Label5: TLabel;
    BitBtnUp: TBitBtn;
    BitBtnDown: TBitBtn;
    DateTimePicker1: TDateTimePicker;
    ComboBox3: TComboBox;
    TabSheet2: TTabSheet;
    ScrollBox2: TScrollBox;
    ListView1: TListView;
    TabSheet3: TTabSheet;
    Panel3: TPanel;
    vle: TValueListEditor;
    PopupMenu2: TPopupMenu;
    N4: TMenuItem;
    N5: TMenuItem;
    NewTraining: TAction;
    DelTrainning: TAction;
    Button2: TControlBar;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    Label4: TLabel;
    Label6: TLabel;
    procedure ToolButton3Click(Sender: TObject);
    procedure DelTrainningExecute(Sender: TObject);
    procedure NewTrainingExecute(Sender: TObject);
    procedure DockTabSet1Change(Sender: TObject; NewTab: Integer;
      var AllowChange: Boolean);
    procedure N4Click(Sender: TObject);
    procedure ScrollBox1MouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure ScrollBox1MouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure N2Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure TreeView1Expanded(Sender: TObject; Node: TTreeNode);
    procedure TreeView1Collapsed(Sender: TObject; Node: TTreeNode);
    procedure vleClick(Sender: TObject);
    procedure TreeView1Change(Sender: TObject; Node: TTreeNode);
    procedure ListView1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure N1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure DateTimePicker1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DateTimePicker1CloseUp(Sender: TObject);
    procedure DelZoneExecute(Sender: TObject);
    procedure AddZoneExecute(Sender: TObject);
    procedure DelExecExecute(Sender: TObject);
    procedure AddExecExecute(Sender: TObject);
    procedure BitBtnUpClick(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);

    procedure ComboBox1Change(Sender: TObject);

    procedure BitBtnDownClick(Sender: TObject);

    procedure FormCreate(Sender: TObject);

    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    GlobalX, GlobalY : integer;
    errfound : boolean;
    DeleteItem, DockChanged : boolean;
    procedure ClickNext(Sender: TObject);
    procedure ClickPrev(Sender: TObject);
    procedure CheckClick(Sender: TObject);
    procedure ReadExecsFromList;
    procedure ReadOFPFromList;
    procedure SaveNewLeaf(fname : string; dl : TDataLeaf);
    procedure SaveAll(fname : string);
    procedure SportReadFile(fname:string);
    procedure ShowLeaf(dl : TDataLeaf; clearpanel1 : boolean; trcount : integer);
    function  AddTreeData(dl : TDataLeaf):TDataLeaf;
    function  AddMicroCicle(T : TTreeNode; s : string): TTreeNode;
    procedure RecountLabels;
    function  GetCurrentTreeNode(dl : TdataLeaf):TTreeNode;
    procedure ShowTreeNode(Node : TTreeNode);
    function  GetTrCount(Node : TTreeNode) : integer;
    function  GetCounter(Node : TTreeNode):integer;
    function  GetLeafInfo(dl : TDataLeaf; var firstd, lastd : TDateTime):string;
    procedure NewTrain;
    procedure SuperClearPanel;
  public
    { Public declarations }
    Modified : boolean;
    MenList : TList;                    // Список спортсменов
    OFPList, ExecList : TStringList;
    ManStream : TMemoryStream;
    fname : string;
    CurrentLeaf : TDataLeaf;
    BigDataLeaf: TObjectList;           // Сïèñîê базы данных
    procedure UpdateDataLeaf(dl : TDataLeaf);
    procedure OpenFile;
    procedure ReadManStream(fname : string; var ManStream : TMemoryStream);
    procedure ScanDB(fname : string);
    procedure ClearPanel;
    procedure SaveDL(dl : TDataLeaf; var M : TMemoryStream);
    function  SaveDataBase : boolean;
    function  AddMezoCicle(M : integer; selected : boolean) : TTreeNode;
    function  AddMezoCicleEmpty(M : integer; selected : boolean) : TTreeNode;
    procedure ReadExecs(selectitem : boolean);
    procedure InitFirst;
  end;

var
  FormTop : integer;
  tManRec : pManRec;

implementation

uses Main, RazrForm, ProgressUnit, DateUnit;

{$R *.dfm}


function TMDIChild.AddMezoCicleEmpty(M : integer; selected : boolean) : TTreeNode;
var
 foundmzc : boolean;
 i, mzc : integer;
 TN, TN1 : TTreeNode;
begin
  foundmzc := false;
  mzc:=1;

  for i := 0 to TreeView1.Items.Count - 1 do
  if TreeView1.Items[i].Level=0 then
    begin
     mzc := StrToInt(copy(TreeView1.Items[i].Text,1,pos(' ',TreeView1.Items[i].Text)-1));
     if M = mzc then
      TN := TreeView1.Items[i];
     foundmzc := true;
    end;

  if foundmzc then
  begin
   if M <> mzc then
    TN := TreeView1.Items.Add(nil,IntToStr(mzc+1)+' Мезоцикл');
   TN.ImageIndex := 0;
   TN.SelectedIndex := 0;
   TN.Selected := selected;
  end
  else
  begin
   if M <> mzc then
    TN := TreeView1.Items.Add(nil,IntToStr(M+1) + ' Мезоцикл');
   TN.ImageIndex := 0;
   TN.SelectedIndex := 0;
   TN.Selected := selected;
  end;

  Result := TN;
end;


function TMDIChild.AddMezoCicle(M : integer; selected : boolean) : TTreeNode;
var
 foundmzc : boolean;
 i, mzc : integer;
 TN, TN1 : TTreeNode;
begin
  foundmzc := false;
  mzc:=1;

  if TreeView1.Items.Count=0 then
    TreeView1.Items.AddFirst(nil,'1 Мезоцикл');

  for i := 0 to TreeView1.Items.Count - 1 do
  if TreeView1.Items[i].Level=0 then
    begin
     mzc := StrToInt(copy(TreeView1.Items[i].Text,1,pos(' ',TreeView1.Items[i].Text)-1));
     if M = mzc then
     begin
      TN := TreeView1.Items[i];
      foundmzc := true;
      break;
     end;
    end;

  if foundmzc then
  begin
//   if M <> mzc then
//    TN := TreeView1.Items.Add(nil,IntToStr(mzc+1)+' Мезоцикл');
//   TN.ImageIndex := 0;
//   TN.SelectedIndex := 0;
   TN.Selected := selected;
  end
  else
  begin
   if M <> mzc then
    TN := TreeView1.Items.Add(nil,IntToStr(M) + ' Мезоцикл');
   TN.ImageIndex := 0;
   TN.SelectedIndex := 0;
   TN.Selected := selected;
  end;

  if selected then
   TreeView1.FullExpand;

  Result := TN;
end;

function MonthNumber(m : string):integer;
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

procedure TMDIChild.ScanDB(fname : string);
var
 F : TFileStream;
 mansize, cnt, id, i,j,len : integer;
 buffer : PAnsiChar;
 s : string;
 LI : TListItem;
begin
 ComboBox1.Items.Clear;
 ComboBox2.Items.Clear;
 id := 0;
 if FileExists(fname) then
    try
     F := TFileStream.Create(fname,fmOpenRead);
     F.Read(len,4);
     if len<=DBVersion then
     begin
      F.Read(mansize,4);
      while F.Position < mansize do
      begin
      New(tManRec);
      MenList.Add(tManRec);
      tManRec^.SportList := TStringList.Create;
      inc(id);
      // Размер записи
      F.Read(len,4);
      // Фамилия
      F.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      F.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := StrPas(buffer);
      FreeMem(buffer);
      // Имя
      F.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      F.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := s + ' ' + StrPas(buffer);
      FreeMem(buffer);
      // Отчество
      F.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      F.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := s + ' ' + StrPas(buffer);
      FreeMem(buffer);
      ComboBox1.Items.Add(s);

      // Разряд
      F.Read(len,4);
      F.Read(len,4);
      F.Seek(len,soFromCurrent);

      F.Read(cnt,4);
      s := '';
      for I := 0 to cnt - 1 do
      begin
        F.Read(len,4);
        buffer := AllocMem(len+1);
        buffer[len] := Chr(0);
        F.Read(buffer^,len);
        buffer[len]:=Chr(0);
        s := StrPas(buffer);
        FreeMem(buffer);
        tManRec^.SportList.Add(s);
        if id=1 then
         ComboBox2.Items.Add(s);
       end;
      end;
     end
     else
      MessageBOX(Handle, 'Версия справочника "'+ManTitle+'" устарела...', MainTitle ,
      MB_ICONWarning);

     F.Free;

    except
     MessageBOX(Handle, 'Ошибка при чтении справочника "'+ManTitle+'"...', MainTitle ,
     MB_ICONERROR);
    end;
end;

procedure TMDIChild.ScrollBox1MouseWheelDown(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
 ScrollBox1.ScrollInView(Button2);
end;

procedure TMDIChild.ScrollBox1MouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
 ScrollBox1.ScrollInView(Label5);
end;

procedure TDataPoint.InitVisual(PosX : integer);
begin
  Edit1 := TMemo.Create(Panel);
  Edit1.Parent := Panel;
  Edit1.Width := 48;
  Edit1.Height := 21;
  Edit1.Top := 14;
  Edit1.Left := 2;
  Edit1.Alignment := taCenter;
  Edit1.Visible := true;
  Edit2 := TMemo.Create(Panel);
  Edit2.Parent := Panel;
  Edit2.Width := 48;
  Edit2.Height := 21;
  Edit2.Top := 36;
  Edit2.Left := 2;
  Edit2.Alignment := taCenter;
  Edit2.Visible := true;

  Edit31 := TMemo.Create(Panel);
  Edit31.Parent := Panel;
  Edit31.Width := 32;
  Edit31.Height := 43;
  Edit31.Top := 14;
  Edit31.Left := 51;
  Edit31.ReadOnly := True;;
  Edit31.Enabled := false;
  Edit31.Alignment := taCenter;
  Edit31.Visible := true;

  Edit3 := TMemo.Create(Panel);
  Edit3.Parent := Panel;
  Edit3.Width := 28;
  Edit3.Height := 21;
  Edit3.Top := 27;
  Edit3.Left := 53;
  Edit3.BorderStyle := bsNone;
  Edit3.Alignment := taCenter;
  Edit3.Enabled := true;
  Edit3.Visible := true;

  Label1 := TLabel.Create(Panel);
  Label1.Parent := Panel;
  Label1.Width := 33;
  Label1.Height := 14;
  Label1.Top := 0;
  Label1.Left := 26;
  Label1.Visible := true;
  Label1.Transparent := true;
  Label1.Caption := 'Зона ' + IntToStr(PosX);
end;

constructor TDataPoint.Create;
begin
  inherited Create;
end;

destructor TDataPoint.Destroy;
begin
{  if Edit1<>nil then
   Edit1.Free;
  if Edit2<>nil then
  Edit2.Free;
  if Edit3<>nil then
  Edit3.Free;
  if Label1<>nil then
  Label1.Free;
  if Panel<>nil then
  Panel.Free; }
  inherited Destroy;
end;

procedure TDataContent.InitVisual(PosX : integer);
begin
  Label1 := TLabel.Create(Panel);
  Label1.Parent := Panel;
  Label1.Width := 252;
  Label1.Height := 21;
  Label1.Top := 3;
  Label1.Left := 7;

  Label1.Caption := IntToStr(Y) + ' упражнение:';
  Label1.Transparent := True;
  Label1.Font.Style := [fsBold];
  Label1.Visible := true;

  CheckBox1 := TCheckBox.Create(Panel);
  CheckBox1.Parent := Panel;
  CheckBox1.Width := 252;
  CheckBox1.Height := 21;
  CheckBox1.Top := 37;
  CheckBox1.Left := 7;
  CheckBox1.Caption := 'ОФП';
  CheckBox1.Visible := true;
  CheckBox1.Enabled := true;

  ComboBox1 := TComboBox.Create(Panel);
  ComboBox1.Parent := Panel;
  ComboBox1.Width := 252;
  ComboBox1.Height := 21;
  ComboBox1.Top := 18;
  ComboBox1.Left := 5;
  ComboBox1.Visible := true;
  ComboBox1.Style := csDropDownList;

  BitBtnNext := TBitBtn.Create(ScrollBox);
  BitBtnNext.Parent := ScrollBox;
  BitBtnNext.Hint := 'Добавить зону (Ctrl-W)';
  BitBtnNext.ShowHint := true;
  BitBtnNext.Width := 42;
  BitBtnNext.Height := 17;
  BitBtnNext.Top := PosX + 65;
  BitBtnNext.Left := 295 + 43;
  BitBtnNext.Visible := true;

  BitBtnPrev := TBitBtn.Create(ScrollBox);
  BitBtnPrev.Hint := 'Удалить зону (Ctrl-Q)';
  BitBtnPrev.ShowHint := true;
  BitBtnPrev.Parent := ScrollBox;
  BitBtnPrev.Width := 42;
  BitBtnPrev.Height := 17;
  BitBtnPrev.Top := PosX + 65;
  BitBtnPrev.Left := 295;
  BitBtnPrev.Visible := true;
end;

constructor TDataContent.Create;
begin
  DataGrid := TObjectList.Create;
  inherited Create;
end;

destructor TDataContent.Destroy;
begin
{  if Label1<>nil then
   Label1.Free;
  if ComboBox1<>nil then
   ComboBox1.Free;
  if BitBtnNext<>nil then
   BitBtnNext.Free;
  if BitBtnPrev<>nil then
   BitBtnPrev.Free;
  if CheckBox1<>nil then
   CheckBox1.Free;
  if Panel<>nil then
   Panel.Free;   }
  while DataGrid.Count>0 do
   DataGrid.Delete(DataGrid.Count-1);
  DataGrid.Free;
  inherited Destroy;
end;

constructor TDataLeaf.Create;
begin
  BigDataGrid := TObjectList.Create;
  RelaxList := TStringList.Create;
  ControlList := TStringList.Create;
  inherited Create;
end;

destructor TDataLeaf.Destroy;
begin
  while BigDataGrid.Count>0 do
   BigDataGrid.Delete(BigDataGrid.Count-1);
  BigDataGrid.Free;
  RelaxList.Free;
  ControlList.Free;
  inherited Destroy;
end;

procedure TMDIChild.AddExecExecute(Sender: TObject);
begin
 BitBtnDownClick(Sender);
end;

procedure TMDIChild.AddZoneExecute(Sender: TObject);
begin
 ClickNext(Sender);
end;

procedure TMDIChild.BitBtnDownClick(Sender: TObject);
var
 Panel1, Panel2 : TPanel;
begin
    Modified := true;

    BitBtnUp.Top := BitBtnUp.Top + 93;
    BitBtnDown.Top := BitBtnDown.Top + 93;
    Button2.Top := Button2.Top + 93;

    Panel1 := TPanel.Create(ScrollBox1);
    Panel1.Parent := ScrollBox1;
    Panel1.Width := 262;
    Panel1.Height := 59;
    Panel1.Top := BitBtnUp.Top;
    Panel1.Left := 28;
    Panel1.Visible := false;
    CurrentLeaf.BigDataGrid.Add(TDataContent.Create);
    TDataContent(CurrentLeaf.BigDataGrid.Last).trindex := CurrentLeaf.trcount;
    TDataContent(CurrentLeaf.BigDataGrid.Last).Panel := Panel1;
    TDataContent(CurrentLeaf.BigDataGrid.Last).ScrollBox := ScrollBox1;
    Inc(GlobalY);
    TDataContent(CurrentLeaf.BigDataGrid.Last).X := GlobalX;
    TDataContent(CurrentLeaf.BigDataGrid.Last).Y := GlobalY;
    TDataContent(CurrentLeaf.BigDataGrid.Last).InitVisual(BitBtnUp.Top);
    TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnNext.OnClick := ClickNext;
    TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnPrev.OnClick := ClickPrev;
    TDataContent(CurrentLeaf.BigDataGrid.Last).CheckBox1.OnClick := CheckClick;

    Panel2 := TPanel.Create(ScrollBox1);
    Panel2.Parent := ScrollBox1;
    Panel2.Width := 85;
    Panel2.Height := 59;
    Panel2.Top := BitBtnUp.Top;
    Panel2.Left := 295;
    Panel2.Visible := false;
    TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Add(TDataPoint.Create);
    TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Last).Panel := Panel2;
    TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Last).X := GlobalX;
    TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Last).InitVisual(1);

    ImageList1.GetBitmap(2,TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnPrev.Glyph);
    ImageList1.GetBitmap(3,TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnNext.Glyph);

    Panel1.Visible := true;
    Panel2.Visible := true;
//    ScrollBox1.ScrollInView(TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnPrev);
    ScrollBox1.ScrollInView(Button2);

    ReadExecsFromList;
    RecountLabels;

end;

procedure TMDIChild.FormActivate(Sender: TObject);
var
 i : integer;
begin
 DeleteItem := False;

 for I := 0 to MainForm.DockTabSet1.Tabs.Count - 1 do
  if Caption = Trim(MainForm.DockTabSet1.Tabs[i]) then
  begin
    MainForm.DockTabSet1.TabIndex := i;
    break;
  end;

 MainForm.N9.Enabled := false;
 MainForm.ToolButton4.Enabled := false;
 MainForm.ToolButton2.Enabled := true;
 MainForm.FileSaveItem.Enabled := true;
 MainForm.N18.Enabled := true;
 MainForm.C1.Enabled := true;

 MainForm.ToolButton8.Enabled := true;
 MainForm.N14.Enabled := true;

 MainForm.ToolButton11.Enabled := true;
 MainForm.N16.Enabled := true;
 MainForm.ToolButton12.Enabled := true;
 MainForm.N17.Enabled := true;
 MainForm.ToolButton14.Enabled := true;
 MainForm.N24.Enabled := true;
 MainForm.N27.Enabled := false;

 MainForm.SB.Panels[0].Text := '';
end;

procedure TMDIChild.FormClose(Sender: TObject; var Action: TCloseAction);
var
 i : integer;
begin

  ExecList.Free;
  while BigDataLeaf.Count>0 do
   BigDataLeaf.Delete(BigDataLeaf.Count-1);
  BigDataLeaf.Free;

  MenList.Free;
  OFPList.Free;
  ManStream.Free;

  if MainForm.MDIChildCount = 1 then
   begin
    MainForm.ToolButton8.Enabled := false;
    MainForm.N14.Enabled := false;
    MainForm.ToolButton2.Enabled := false;
    MainForm.FileSaveItem.Enabled := false;
    MainForm.DockTabSet1.Visible := false;
    MainForm.N18.Enabled := false;

    MainForm.ToolButton11.Enabled := false;
    MainForm.N16.Enabled := false;
    MainForm.ToolButton12.Enabled := false;
    MainForm.N17.Enabled := false;
    MainForm.ToolButton14.Enabled := false;
    MainForm.N24.Enabled := false;
    MainForm.N27.Enabled := false;
    MainForm.SB.Panels[0].Text := '';
   end;

  for I := 0 to MainForm.DockTabSet1.Tabs.Count - 1 do
   if Trim(MainForm.DockTabSet1.Tabs[i]) = Caption then
    begin
     MainForm.DockTabSet1.Tabs.Delete(i);
     break;
    end;

  Action := caFree;
end;

procedure TMDIChild.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
 J : Integer;
begin
  if Modified then
  begin
    case MessageBOX(Handle, PChar('Сохранить изменения в базе данных "' +
      Caption + '" ?'), PChar(MainTitle) ,
      MB_YESNOCANCEL or MB_ICONQUESTION) of
      IDYes:
        begin
          SaveDataBase;
          if errfound then
           CanClose := false
          else
           CanClose := true;
        end;
      IDNo:
        begin
          CanClose := true;
        end;
      IDCancel:
          CanClose := false;
     end;
  end
  else
    CanClose := true;
end;

procedure TMDIChild.FormCreate(Sender: TObject);
var
 ST : TStringList;
 i : integer;
 LI : TListItem;
 LabelI : TLabel;
 EditI : TEdit;
begin
 Modified := false;
 ExecList := TStringList.Create;
 BigDataLeaf := TObjectList.Create;
 MenList := TList.Create;
 ManStream := TMemoryStream.Create;
 OFPList := TStringList.Create;

 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'ofp.mtx';
 RazrDlg.GetRazr(OFPList);

 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'relax.mtx';
 ST := TStringList.Create;
 RazrDlg.GetRazr(ST);
 for i := 0 to ST.Count - 1 do
 begin
   LI := ListView1.Items.Add;
   LI.Caption := ST[i];
 end;
 ST.Free;

 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'control.mtx';
 ST := TStringList.Create;
 RazrDlg.GetRazr(ST);
 vle.TitleCaptions[0] := 'Вид контроля';
 vle.TitleCaptions[1] := 'Значение';
 for i := 0 to ST.Count - 1 do
  vle.InsertRow(ST[i],'',true);
 ST.Free;

 PageControl1.ActivePageIndex := 0;
 DateTimePicker1.DateTime := Now;   
 DockTabSet1.Tabs.Add('1 Микроцикл');
 DockTabSet1.Tabs.Add('2 Микроцикл');
 DockTabSet1.Tabs.Add('3 Микроцикл');
 DockTabSet1.Tabs.Add('4 Микроцикл');
 DockTabSet1.Tabs.Add('5 Микроцикл');
 DockTabSet1.Tabs.Add('6 Микроцикл');
 DockTabSet1.Tabs.Add('7 Микроцикл');
 DockTabSet1.Tabs.Add('8 Микроцикл');
 DockTabSet1.TabIndex := 0;
end;


procedure TMDIChild.OpenFile;
begin
 if not FileExists(fname) then
 begin
  ClearPanel;
  InitFirst;
  if ComboBox1.Items.Count>0 then
   ComboBox1.ItemIndex := 0;
  if ComboBox2.Items.Count>0 then
   ComboBox2.ItemIndex := 0;
  Label3.Caption := ComboBox2.Text;
  ReadExecs(true);
  if TreeView1.Items.Count=0 then
    TreeView1.Items.AddFirst(nil,'1 Мезоцикл');
 end
 else
 begin
  ScanDB(fname);
  if ComboBox1.Items.Count>0 then
    ComboBox1.ItemIndex := 0;
  if ComboBox2.Items.Count>0 then
    ComboBox2.ItemIndex := 0;

  SportReadFile(fname);

  if BigDataLeaf.Count=0 then
  begin
//   ClearPanel;
   InitFirst;
   Label3.Caption := ComboBox2.Text;
   ReadExecs(true);
  end
  else
  begin
   CurrentLeaf := TDataLeaf(BigDataLeaf.Last);
   ReadExecs(false);
   ComboBox1.ItemIndex := TDataLeaf(BigDataLeaf.Last).menindex;
   ComboBox1Change(nil);
  end;
 end;
end;

procedure TMDIChild.SportReadFile(fname:string);
var
 F : TFileStream;
 ver, cnt2, cnt, i,j,len : integer;
 buffer : PAnsiChar;
 pbuffer : PByte;
 s : string;
 M : TMemoryStream;
 dt : TDateTime;
 dl : TDataLeaf;
begin
  try
   F := TFileStream.Create(fname,fmOpenRead);
   F.Read(ver,4);
   if ver > DBVersion then
    MessageBOX(Handle, PChar('Версия файла "'+fname+'" не поддерживается...'), PChar(MainTitle) ,
    MB_ICONWARNING)
   else
   begin
    Screen.Cursor := crHourGlass;
    ProgressForm.Caption := 'Открытие базы данных...';
    ProgressForm.pb.Position := 0;
    ProgressForm.Show;

    F.Read(len,4);
    ManStream.CopyFrom(F,len);
    while F.Position < F.Size do
    begin
     ProgressForm.pb.Position := Trunc(F.Position/F.Size*100);
//     ProgressForm.pb.Repaint;
     Application.ProcessMessages;

     F.Read(len,4);

     BigDataLeaf.Add(TDataLeaf.Create);
     TDataLeaf(BigDataLeaf.Last).BigDataGrid := TObjectList.Create;

     F.Read(dt,8);
     TDataLeaf(BigDataLeaf.Last).dt := dt;

     F.Read(dt,8);
     TDataLeaf(BigDataLeaf.Last).dtl := dt;

     TDataLeaf(BigDataLeaf.Last).saved := true;

     F.Read(len,4);
     TDataLeaf(BigDataLeaf.Last).menindex := len;
     F.Read(len,4);
     buffer := AllocMem(len+1);
     buffer[len] := Chr(0);
     F.Read(buffer^,len);
     buffer[len]:=Chr(0);
     s := StrPas(buffer);
     FreeMem(buffer);
     TDataLeaf(BigDataLeaf.Last).men := s;

     F.Read(len,4);
     TDataLeaf(BigDataLeaf.Last).sportindex := len;
     F.Read(len,4);
     buffer := AllocMem(len+1);
     buffer[len] := Chr(0);
     F.Read(buffer^,len);
     buffer[len]:=Chr(0);
     s := StrPas(buffer);
     FreeMem(buffer);
     TDataLeaf(BigDataLeaf.Last).sport := s;

     F.Read(cnt,4);
     TDataLeaf(BigDataLeaf.Last).trcount := cnt;
     F.Read(cnt,4);
     TDataLeaf(BigDataLeaf.Last).mezo := cnt;
     F.Read(cnt,4);
     TDataLeaf(BigDataLeaf.Last).micro := cnt;

     if ComboBox1.ItemIndex = TDataLeaf(BigDataLeaf.Last).menindex then
      AddTreeData(TDataLeaf(BigDataLeaf.Last));

     F.Read(cnt,4);
     for i := 0 to cnt - 1 do
      begin
       TDataLeaf(BigDataLeaf.Last).BigDataGrid.Add(TDataContent.Create);

       F.Read(len,4);
       TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).TrIndex := len;
       F.Read(dt,8);
       TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).dt := dt;

       F.Read(len,4);
       TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).ExecIndex := len;
       F.Read(len,4);
       buffer := AllocMem(len+1);
       buffer[len] := Chr(0);
       F.Read(buffer^,len);
       buffer[len]:=Chr(0);
       s := StrPas(buffer);
       FreeMem(buffer);
       TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).Exec := s;

       F.Read(len,4);
       TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).IsOFP := len;
       F.Read(len,4);
       TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).OFPIndex := len;
       F.Read(len,4);
       buffer := AllocMem(len+1);
       buffer[len] := Chr(0);
       F.Read(buffer^,len);
       buffer[len]:=Chr(0);
       s := StrPas(buffer);
       FreeMem(buffer);
       TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).OFP := s;

       F.Read(cnt2,4);
       for j := 0 to cnt2 - 1 do
        begin
         F.Read(len,4);
         buffer := AllocMem(len+1);
         buffer[len] := Chr(0);
         F.Read(buffer^,len);
         buffer[len]:=Chr(0);
         s := StrPas(buffer);
         FreeMem(buffer);
         TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).DataGrid.Add(TDataPoint.Create);
         TDataPoint(TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).DataGrid.Last).Res1 := s;
         F.Read(len,4);
         buffer := AllocMem(len+1);
         buffer[len] := Chr(0);
         F.Read(buffer^,len);
         buffer[len]:=Chr(0);
         s := StrPas(buffer);
         FreeMem(buffer);
         TDataPoint(TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).DataGrid.Last).Res2 := s;
         F.Read(len,4);
         buffer := AllocMem(len+1);
         buffer[len] := Chr(0);
         F.Read(buffer^,len);
         buffer[len]:=Chr(0);
         s := StrPas(buffer);
         FreeMem(buffer);
         TDataPoint(TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).DataGrid.Last).Res3 := s;
        end;

      end;

      F.Read(cnt,4);
      for i := 0 to cnt - 1 do
       begin
         F.Read(len,4);
         buffer := AllocMem(len+1);
         buffer[len] := Chr(0);
         F.Read(buffer^,len);
         buffer[len]:=Chr(0);
         s := StrPas(buffer);
         FreeMem(buffer);
         TDataLeaf(BigDataLeaf.Last).RelaxList.Add(s);
       end;

      if ver >= 2 then
      begin
       F.Read(cnt,4);
       for i := 0 to cnt - 1 do
       begin
         F.Read(len,4);
         buffer := AllocMem(len+1);
         buffer[len] := Chr(0);
         F.Read(buffer^,len);
         buffer[len]:=Chr(0);
         s := StrPas(buffer);
         FreeMem(buffer);
         TDataLeaf(BigDataLeaf.Last).ControlList.Add(s);
       end;
      end;

    end;

   end;

//   F.Read(len,4);
//   TreeView1.Selected := TreeView1.Items[len];

   F.Free;
   Screen.Cursor := crDefault;
   ProgressForm.Close;
  except
   Screen.Cursor := crDefault;
   ProgressForm.Close;
   MessageBOX(Handle, PChar('Ошибка при чтени файла "'+fname+'"...'), Pchar(MainTitle) ,
   MB_ICONERROR);
  end;

end;

function TMDIChild.GetTrCount(Node : TTreeNode) : integer;
var
 MN : TTreeNode;
 cnt : integer;
begin
   cnt := 0;
   if Node.HasChildren then
   begin
     MN := Node.getFirstChild;
     while MN<>nil do
     begin
      cnt := cnt + GetTrCount(MN);
      if MN.Data <> nil then
       cnt := cnt + TDataLeaf(MN.Data).trcount;
      MN := Node.GetNextChild(MN);
     end;
   end;
   Result := cnt;
end;

function TMDIChild.GetCounter(Node : TTreeNode):integer;
var
 MN : TTreeNode;
 trcount, cnt, i,j, predindex : integer;
 dl : TDataLeaf;
begin
    trcount := 0;
    for i := 0 to TreeView1.Items.Count -1 do
    begin
      MN := TreeView1.Items[i];
      if MN.Level = 2 then
      begin
      if Node <> nil then
       if Node = MN then
        break;
      dl := TDataLeaf(MN.Data);
      predindex := 0;
      for j := 0 to dl.BigDataGrid.Count - 1 do
       begin
        if TDataContent(dl.BigDataGrid.Items[j]).trindex <> predindex then
         inc(trcount);
        predindex := TDataContent(dl.BigDataGrid.Items[j]).trindex;
       end;
      end;
    end;
    Result := trcount;
end;

procedure TMDIChild.ShowTreeNode(Node : TTreeNode);
begin
 N1.Enabled := false;
 if Node <> nil then
  if Node.Level = 2 then
   begin
    N1.Enabled := true;
    if DeleteItem then
     ShowLeaf(TDataLeaf(Node.Data),false,GetCounter(Node)+1)
    else
     ShowLeaf(TDataLeaf(Node.Data),true,GetCounter(Node)+1);
    DockChanged := true;
    DockTabSet1.TabIndex := StrToInt(copy(Node.Parent.Text,1,1))-1;
    DockChanged := false;
    MainForm.SB.Panels[0].Text := Node.Text;
    MainForm.SB.Panels[1].Text := 'Количество тренировок: ' + IntToStr(TDataLeaf(Node.Data).trcount);
   end
  else
   begin
    if Node.Parent <> nil then
    MainForm.SB.Panels[0].Text := Node.Text;
    MainForm.SB.Panels[1].Text := 'Количество тренировок: ' + IntToStr(GetTrCount(Node));
   end;
end;

procedure TMDIChild.ToolButton3Click(Sender: TObject);
begin
 if SaveDataBase then
  ShowTreeNode(TreeView1.Selected);
end;

procedure TMDIChild.TreeView1Change(Sender: TObject; Node: TTreeNode);
begin
 ShowTreeNode(Node);
end;

procedure TMDIChild.TreeView1Collapsed(Sender: TObject; Node: TTreeNode);
begin
 Node.ImageIndex := 0;
end;

procedure TMDIChild.TreeView1Expanded(Sender: TObject; Node: TTreeNode);
begin
  Node.ImageIndex := 1;
end;

procedure TMDIChild.ShowLeaf(dl : TDataLeaf; clearPanel1:boolean; trcount : integer);
var
  Panel1, Panel2 : TPanel;
  i,j, ofpcount, predindex : integer;
  TrainLabel : TLabel;
  TrainBox : TComboBox;
  TrainDate : TDateTimePicker;
begin


  PageControl1.Visible := false;

  if clearpanel1 then
   ClearPanel;


  BitBtnUp.Top := 25;
  BitBtnDown.Top := 55;
  Button2.Top := 140;

//  ComboBox1.ItemIndex := dl.menindex;

  ComboBox2.Items.Clear;
  tManRec := MenList.Items[ComboBox1.ItemIndex];
  ComboBox2.Items.AddStrings(tManRec^.SportList);

  ComboBox2.ItemIndex := dl.sportindex;
  Label3.Caption := ComboBox2.Text;
  Label5.Caption := IntToStr(trcount);

  DateTimePicker1.DateTime := dl.dt;
  ReadExecs(false);


  for i := 0 to ListView1.Items.Count - 1 do
    ListView1.Items[i].Checked := false;

  for i := 0 to ListView1.Items.Count - 1 do
   for j := 0 to dl.RelaxList.Count - 1 do
    begin
     ListView1.Items[i].Checked := dl.RelaxList[j] = ListView1.Items[i].Caption;
     if ListView1.Items[i].Checked then break;
    end;

  for i := 1 to vle.RowCount-1 do
   vle.Values[vle.Keys[i]] := '';

  for i := 0 to dl.ControlList.Count-1 do
   vle.Values[vle.Keys[i+1]] := dl.ControlList[i];

  Panel1 := TPanel.Create(ScrollBox1);
  Panel1.Parent := ScrollBox1;
  Panel1.Width := 262;
  Panel1.Height := 59;
  Panel1.Top := 26;
  Panel1.Left := 28;
  Panel1.Visible := true;

  TDataContent(dl.BigDataGrid.Items[0]).Panel := Panel1;
  predindex := TDataContent(dl.BigDataGrid.Items[0]).trindex;
  TDataContent(dl.BigDataGrid.Items[0]).ScrollBox := ScrollBox1;
  TDataContent(dl.BigDataGrid.Items[0]).InitVisual(26);
  TDataContent(dl.BigDataGrid.Items[0]).BitBtnNext.OnClick := ClickNext;
  TDataContent(dl.BigDataGrid.Items[0]).BitBtnPrev.OnClick := ClickPrev;
  if TDataContent(dl.BigDataGrid.Items[0]).isOFP = 1 then
   TDataContent(dl.BigDataGrid.Items[0]).CheckBox1.State := cbChecked;
  TDataContent(dl.BigDataGrid.Items[0]).CheckBox1.OnClick := CheckClick;

  ofpcount := 0;
  if TDataContent(dl.BigDataGrid.Items[0]).isOFP=1 then
  begin
   TDataContent(dl.BigDataGrid.Items[0]).ComboBox1.Items.AddStrings(OFPList);
   TDataContent(dl.BigDataGrid.Items[0]).ComboBox1.ItemIndex := TDataContent(dl.BigDataGrid.Items[0]).OFPIndex;
   inc(ofpcount);
   TDataContent(dl.BigDataGrid.Items[0]).Label1.Font.Color := clRed;
   TDataContent(dl.BigDataGrid.Items[0]).Label1.Caption := IntToStr(ofpcount) + ' ОФП:';
   TDataContent(dl.BigDataGrid.Items[0]).BitBtnPrev.Visible := False;
   TDataContent(dl.BigDataGrid.Items[0]).BitBtnNext.Visible := False;
  end
  else
  begin
   TDataContent(dl.BigDataGrid.Items[0]).ComboBox1.Items.AddStrings(ExecList);
   TDataContent(dl.BigDataGrid.Items[0]).ComboBox1.ItemIndex := TDataContent(dl.BigDataGrid.Items[0]).ExecIndex;
   TDataContent(dl.BigDataGrid.Items[0]).Label1.Font.Color := clBlack;
   TDataContent(dl.BigDataGrid.Items[0]).Label1.Caption := '1 упражнение:';
  end;

  Panel2 := TPanel.Create(ScrollBox1);
  Panel2.Parent := ScrollBox1;
  Panel2.Width := 85;
  Panel2.Height := 59;
  Panel2.Top := 26;
  Panel2.Left := 295;
  Panel2.Visible := true;

  TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Panel := Panel2;
  TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).InitVisual(1);

  TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Edit1.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Res1;
  TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Edit2.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Res2;
  TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Edit3.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Res3;

  if TDataContent(dl.BigDataGrid.Items[0]).isOFP=1 then
  begin
   TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Label1.Visible := false;
   TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Edit1.Width := 48+33;
   TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Edit2.Width := 48+33;
   TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Edit3.Visible := false;
   TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[0]).Edit31.Visible := false;
  end;

  ImageList1.GetBitmap(2,TDataContent(dl.BigDataGrid.Items[0]).BitBtnPrev.Glyph);
  ImageList1.GetBitmap(3,TDataContent(dl.BigDataGrid.Items[0]).BitBtnNext.Glyph);

  for i := 1 to TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Count - 1 do
   begin
    Panel1 := TPanel.Create(ScrollBox1);
    Panel1.Parent := ScrollBox1;
    Panel1.Width := 85;
    Panel1.Height := 59;
    Panel1.Top := TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[i-1]).Panel.Top;
    Panel1.Left := TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[i-1]).Panel.Left + 90;
    Panel1.Visible := true;
    TDataContent(dl.BigDataGrid.Items[0]).BitBtnPrev.Left := TDataContent(dl.BigDataGrid.Items[0]).BitBtnPrev.Left + 90;
    TDataContent(dl.BigDataGrid.Items[0]).BitBtnNext.Left := TDataContent(dl.BigDataGrid.Items[0]).BitBtnNext.Left + 90;

    TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[i]).Panel := Panel1;
    TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[i]).InitVisual(i+1);

    TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[i]).Edit1.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[i]).Res1;
    TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[i]).Edit2.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[i]).Res2;
    TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[i]).Edit3.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[0]).DataGrid.Items[i]).Res3;
    ScrollBox1.ScrollInView(Panel1);
   end;

  for i := 1 to dl.BigDataGrid.Count - 1 do
   begin
    if TDataContent(dl.BigDataGrid.Items[i]).trindex <> predindex then
    begin
     BitBtnUp.Top := BitBtnUp.Top + 113;
     BitBtnDown.Top := BitBtnDown.Top + 113;
     Button2.Top := Button2.Top + 113;

     TrainLabel := TLabel.Create(ScrollBox1);
     TrainLabel.Parent := ScrollBox1;
     TrainLabel.Height := 13;
     TrainLabel.Top := BitBtnUp.Top-20;
     TrainLabel.Left := 8;
     TrainLabel.Font.Color:= clMaroon;
     TrainLabel.Font.Style := [fsBold];
     TrainLabel.Caption := intToStr(TDataContent(dl.BigDataGrid.Items[i]).trindex + trcount-1);
     TrainLabel.Visible := true;
     TDataContent(dl.BigDataGrid.Items[i]).TrainLabel := TrainLabel;

     TrainBox := TComboBox.Create(ScrollBox1);
     TrainBox.Parent := ScrollBox1;
     TrainBox.Visible := false;
     TrainBox.ItemHeight := 13;
     TrainBox.Top := BitBtnUp.Top-24;
     TrainBox.Width := 98;
     TrainBox.Left := 28;
     TrainBox.Style := csDropDownList;
     TrainBox.Items.Add('Тренировка');
     TrainBox.Items.Add('Соревнование');
     TrainBox.Items.Add('Сбор');
     TrainBox.ItemIndex := TDataContent(dl.BigDataGrid.Items[i]).trainkind;
     TrainBox.Visible := true;
     TDataContent(dl.BigDataGrid.Items[i]).TrainBox := TrainBox;

     TrainDate := TDateTimePicker.Create(ScrollBox1);
     TrainDate.Parent := ScrollBox1;
     TrainDate.Height := 21;
     TrainDate.Top := BitBtnUp.Top-24;
     TrainDate.Width := 91;
     TrainDate.Left := 132;
     TrainDate.Visible := true;
//     TrainDate.Enabled := false;
     if TDataContent(dl.BigDataGrid.Items[i]).dt = 0 then
      TrainDate.Date := Now
     else
      TrainDate.Date := TDataContent(dl.BigDataGrid.Items[i]).dt;
     TDataContent(dl.BigDataGrid.Items[i]).TrainDate := TrainDate;

    end
    else
    begin
     BitBtnUp.Top := BitBtnUp.Top + 93;
     BitBtnDown.Top := BitBtnDown.Top + 93;
     Button2.Top := Button2.Top + 93;
    end;
    predindex := TDataContent(dl.BigDataGrid.Items[i]).trindex;

    Panel1 := TPanel.Create(ScrollBox1);
    Panel1.Parent := ScrollBox1;
    Panel1.Width := 262;
    Panel1.Height := 59;
    Panel1.Top := BitBtnUp.Top;
    Panel1.Left := 28;
    Panel1.Visible := true;
    TDataContent(dl.BigDataGrid.Items[i]).Panel := Panel1;
    TDataContent(dl.BigDataGrid.Items[i]).ScrollBox := ScrollBox1;
    TDataContent(dl.BigDataGrid.Items[i]).InitVisual(BitBtnUp.Top);
    TDataContent(dl.BigDataGrid.Items[i]).BitBtnNext.OnClick := ClickNext;
    TDataContent(dl.BigDataGrid.Items[i]).BitBtnPrev.OnClick := ClickPrev;

    if TDataContent(dl.BigDataGrid.Items[i]).isOFP = 1 then
     TDataContent(dl.BigDataGrid.Items[i]).CheckBox1.State := cbChecked;
    TDataContent(dl.BigDataGrid.Items[i]).CheckBox1.OnClick := CheckClick;



    if TDataContent(dl.BigDataGrid.Items[i]).isOFP=1 then
    begin
     TDataContent(dl.BigDataGrid.Items[i]).ComboBox1.Items.AddStrings(OFPList);
     TDataContent(dl.BigDataGrid.Items[i]).ComboBox1.ItemIndex := TDataContent(dl.BigDataGrid.Items[i]).OFPIndex;
     inc(ofpcount);
     TDataContent(dl.BigDataGrid.Items[i]).Label1.Font.Color := clRed;
     TDataContent(dl.BigDataGrid.Items[i]).Label1.Caption := IntToStr(ofpcount) + ' ОФП:';
     TDataContent(dl.BigDataGrid.Items[i]).BitBtnPrev.Visible := False;
     TDataContent(dl.BigDataGrid.Items[i]).BitBtnNext.Visible := False;
    end
    else
    begin
     TDataContent(dl.BigDataGrid.Items[i]).ComboBox1.Items.AddStrings(ExecList);
     TDataContent(dl.BigDataGrid.Items[i]).ComboBox1.ItemIndex := TDataContent(dl.BigDataGrid.Items[i]).ExecIndex;
     TDataContent(dl.BigDataGrid.Items[i]).Label1.Font.Color := clBlack;
     TDataContent(dl.BigDataGrid.Items[i]).Label1.Caption := IntToStr(i-ofpcount+1) + ' упражнение:';
    end;

    Panel2 := TPanel.Create(ScrollBox1);
    Panel2.Parent := ScrollBox1;
    Panel2.Width := 85;
    Panel2.Height := 59;
    Panel2.Top := BitBtnUp.Top;
    Panel2.Left := 295;
    Panel2.Visible := true;

    TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Panel := Panel2;
    TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).InitVisual(1);
    TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit1.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Res1;
    TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit2.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Res2;
    TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit3.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Res3;

    if TDataContent(dl.BigDataGrid.Items[i]).isOFP=1 then
    begin
     TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Label1.Visible := false;
     TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit1.Width := 48+33;
     TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit2.Width := 48+33;
     TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit3.Visible := false;
     TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit31.Visible := false;
    end;

    ImageList1.GetBitmap(2,TDataContent(dl.BigDataGrid.Items[i]).BitBtnPrev.Glyph);
    ImageList1.GetBitmap(3,TDataContent(dl.BigDataGrid.Items[i]).BitBtnNext.Glyph);

//    ScrollBox1.ScrollInView(TDataContent(dl.BigDataGrid.Items[i]).BitBtnPrev);

    for j := 1 to TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Count - 1 do
    begin
     Panel1 := TPanel.Create(ScrollBox1);
     Panel1.Parent := ScrollBox1;
     Panel1.Width := 85;
     Panel1.Height := 59;
     Panel1.Top := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j-1]).Panel.Top;
     Panel1.Left := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j-1]).Panel.Left + 90;
     Panel1.Visible := true;
     TDataContent(dl.BigDataGrid.Items[i]).BitBtnPrev.Left := TDataContent(dl.BigDataGrid.Items[i]).BitBtnPrev.Left + 90;
     TDataContent(dl.BigDataGrid.Items[i]).BitBtnNext.Left := TDataContent(dl.BigDataGrid.Items[i]).BitBtnNext.Left + 90;

     TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Panel := Panel1;
     TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).InitVisual(j+1);

     TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Res1;
     TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Res2;
     TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit3.Text := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Res3;
     ScrollBox1.ScrollInView(Panel1);
    end;

   end;
 CurrentLeaf := dl;
 RecountLabels;
 PageControl1.Visible := true;

end;

procedure TMDIChild.ClearPanel;
var
 i,j : integer;
begin
 if CurrentLeaf<>nil then
  if CurrentLeaf.BigDataGrid<>nil then
  for i := 0 to CurrentLeaf.BigDataGrid.Count - 1 do
   begin
    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).Label1<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).Label1.Free;
    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).TrainLabel<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).TrainLabel.Free;
    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).TrainBox<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).TrainBox.Free;
    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).TrainDate<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).TrainDate.Free;
    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.Free;
    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).CheckBox1<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).CheckBox1.Free;
    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).Panel<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).Panel.Free;
    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Free;
    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Free;
    for j := 0 to TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Count - 1 do
    begin
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Label1<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Label1.Free;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1.Free;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2.Free;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit3<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit3.Free;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit31<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit31.Free;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Panel<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Panel.Free;
    end;
   end;
 BitBtnUp.Top := 25;
 BitBtnDown.Top := 55;
 Button2.Top := 140;
 ScrollBox1.EnableAutoRange;
 ScrollBox1.Repaint;
end;

procedure TMDIChild.ReadExecs(selectitem : boolean);
var
 F : TFileStream;
 len,j,i,cnt : integer;
 buffer : PAnsiChar;
 s : string;
 ST : TStringList;
 found : boolean;
begin
 st := TStringList.Create;
 ExecList.Clear;
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
     found := s = ComboBox2.Text;

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
        if found then ST.Add(s);
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

 if Selectitem then
 for i := 0 to CurrentLeaf.BigDataGrid.Count - 1 do
  if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).IsOFP = 0 then
  begin
    j := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.ItemIndex;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).execindex := j;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.Items.Clear;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.Items.AddStrings(st);
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.ItemIndex := j;
  end;

 ExecList.AddStrings(st);
 st.Free;

end;

procedure TMDIChild.ReadExecsFromList;
var
 i,j : integer;
begin
 for i := 0 to CurrentLeaf.BigDataGrid.Count - 1 do
  if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).IsOFP = 0 then
  begin
    j := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.ItemIndex;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).execindex := j;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.Items.Clear;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.Items.AddStrings(ExecList);
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.ItemIndex := j;
  end;
end;

procedure TMDIChild.ReadOFPFromList;
var
 i,j : integer;
begin
 for i := 0 to CurrentLeaf.BigDataGrid.Count - 1 do
  if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).IsOFP = 1 then
  begin
    j := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.ItemIndex;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ofpindex := j;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.Items.Clear;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.Items.AddStrings(OFPList);
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.ItemIndex := j;
  end;
end;

procedure TMDIChild.BitBtnUpClick(Sender: TObject);
var
 i,j : integer;
begin
    if CurrentLeaf.BigDataGrid.Count>1 then
    begin
     for I := TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Count - 1 downto 0 do
     begin
      j := i;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Label1<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Label1.Destroy;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Edit1<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Edit1.Destroy;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Edit2<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Edit2.Destroy;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Edit3<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Edit3.Destroy;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Edit31<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Edit31.Destroy;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Panel<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Items[j]).Panel.Destroy;
      TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Delete(i);
     end;

    if TDataContent(CurrentLeaf.BigDataGrid.Last).Label1<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Last).Label1.Destroy;
    if TDataContent(CurrentLeaf.BigDataGrid.Last).ComboBox1<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Last).ComboBox1.Destroy;
    if TDataContent(CurrentLeaf.BigDataGrid.Last).CheckBox1<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Last).CheckBox1.Destroy;
    if TDataContent(CurrentLeaf.BigDataGrid.Last).Panel<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Last).Panel.Destroy;
    if TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnNext<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnNext.Destroy;
    if TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnPrev<>nil then
     TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnPrev.Destroy;
    if TDataContent(CurrentLeaf.BigDataGrid.Last).TrainLabel<>nil then
    begin
     TDataContent(CurrentLeaf.BigDataGrid.Last).TrainLabel.Destroy;
     TDataContent(CurrentLeaf.BigDataGrid.Last).TrainBox.Destroy;
     TDataContent(CurrentLeaf.BigDataGrid.Last).TrainDate.Destroy;
     BitBtnUp.Top := BitBtnUp.Top - 113;
     BitBtnDown.Top := BitBtnDown.Top - 113;
     Button2.Top := Button2.Top - 113;
     Dec(CurrentLeaf.trcount);
    end
    else
    begin
     BitBtnUp.Top := BitBtnUp.Top - 93;
     BitBtnDown.Top := BitBtnDown.Top - 93;
     Button2.Top := Button2.Top - 93;
    end;

     CurrentLeaf.BigDataGrid.Delete(CurrentLeaf.BigDataGrid.Count-1);
     Dec(GlobalY);

     ScrollBox1.ScrollInView(BitBtnUp);
     Modified := true;
    end
    else
     N1Click(Sender);
end;

procedure TMDIChild.Button2Click(Sender: TObject);
begin
 NewTrain;
end;

procedure TMDIChild.NewTrain;
var
 Panel1, Panel2 : TPanel;
 TrainLabel : TLabel;
 TrainBox : TComboBox;
 TrainDate : TDateTimePicker;
begin
  if DateForm.ShowModal = mrOk then
  begin
    Modified := true;

    BitBtnUp.Top := BitBtnUp.Top + 113;
    BitBtnDown.Top := BitBtnDown.Top + 113;
    Button2.Top := Button2.Top + 113;

    TrainLabel := TLabel.Create(ScrollBox1);
    TrainLabel.Parent := ScrollBox1;
    TrainLabel.Height := 13;
    TrainLabel.Top := BitBtnUp.Top-20;
    TrainLabel.Left := 5;
    TrainLabel.Font.Color:= clMaroon;
    TrainLabel.Font.Style := [fsBold];
    Inc(CurrentLeaf.trcount);
    GlobalY := 0;
//    TrainLabel.Caption := intToStr(CurrentLeaf.trcount);
    TrainLabel.Caption := '';
    TrainLabel.Visible := true;

    TrainBox := TComboBox.Create(ScrollBox1);
    TrainBox.Parent := ScrollBox1;
    TrainBox.Visible := false;
    TrainBox.ItemHeight := 13;
    TrainBox.Top := BitBtnUp.Top-24;
    TrainBox.Width := 98;
    TrainBox.Left := 28;
    TrainBox.Style := csDropDownList;
    TrainBox.Items.Add('Тренировка');
    TrainBox.Items.Add('Соревнование');
    TrainBox.Items.Add('Сбор');
    TrainBox.ItemIndex := 0;
    TrainBox.Visible := true;

    TrainDate := TDateTimePicker.Create(ScrollBox1);
    TrainDate.Parent := ScrollBox1;
    TrainDate.Height := 21;
    TrainDate.Top := BitBtnUp.Top-24;
    TrainDate.Width := 91;
    TrainDate.Left := 132;
    TrainDate.Visible := true;
    TrainDate.Date := DateForm.dt.Date;
//    TrainDate.Enabled := false;

    Panel1 := TPanel.Create(ScrollBox1);
    Panel1.Parent := ScrollBox1;
    Panel1.Width := 262;
    Panel1.Height := 59;
    Panel1.Top := BitBtnUp.Top;
    Panel1.Left := 28;
    Panel1.Visible := false;


    CurrentLeaf.BigDataGrid.Add(TDataContent.Create);
    TDataContent(CurrentLeaf.BigDataGrid.Last).trindex := CurrentLeaf.trcount;
    TDataContent(CurrentLeaf.BigDataGrid.Last).TrainLabel := TrainLabel;
    TDataContent(CurrentLeaf.BigDataGrid.Last).TrainBox := TrainBox;
    TDataContent(CurrentLeaf.BigDataGrid.Last).TrainDate := TrainDate;
    TDataContent(CurrentLeaf.BigDataGrid.Last).Panel := Panel1;
    TDataContent(CurrentLeaf.BigDataGrid.Last).ScrollBox := ScrollBox1;
    Inc(GlobalY);
    TDataContent(CurrentLeaf.BigDataGrid.Last).X := GlobalX;
    TDataContent(CurrentLeaf.BigDataGrid.Last).Y := GlobalY;
    TDataContent(CurrentLeaf.BigDataGrid.Last).InitVisual(BitBtnUp.Top);
    TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnNext.OnClick := ClickNext;
    TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnPrev.OnClick := ClickPrev;
    TDataContent(CurrentLeaf.BigDataGrid.Last).CheckBox1.OnClick := CheckClick;

    Panel2 := TPanel.Create(ScrollBox1);
    Panel2.Parent := ScrollBox1;
    Panel2.Width := 85;
    Panel2.Height := 59;
    Panel2.Top := BitBtnUp.Top;
    Panel2.Left := 295;
    Panel2.Visible := false;
    TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Add(TDataPoint.Create);
    TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Last).Panel := Panel2;
    TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Last).X := GlobalX;
    TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Last).DataGrid.Last).InitVisual(1);

    ImageList1.GetBitmap(2,TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnPrev.Glyph);
    ImageList1.GetBitmap(3,TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnNext.Glyph);

    Panel1.Visible := true;
    Panel2.Visible := true;
//    ScrollBox1.ScrollInView(TDataContent(CurrentLeaf.BigDataGrid.Last).BitBtnPrev);
    ScrollBox1.ScrollInView(Button2);

    ReadExecsFromList;
    RecountLabels;
 end;
end;

procedure TMDIChild.NewTrainingExecute(Sender: TObject);
begin
 NewTrain;
end;

function TMDIChild.GetCurrentTreeNode(dl : TdataLeaf):TTReeNode;
var
 j : integer;
 TN : TTreeNode;
begin
 TN := nil;
 for j := 0 to TreeView1.Items.Count - 1 do
  if TreeView1.Items[j].Level=2 then
    if TdataLeaf(TreeView1.Items[j].Data) = dl then
      begin
       TN := TreeView1.Items[j];
       break;
      end;
 Result := TN;
end;

function TMDIChild.GetLeafInfo(dl : TDataLeaf; var firstd, lastd : TDateTime):string;
var
 year, month, day : word;
 yearl, monthl, dayl : word;
 yeart, montht, dayt : word;
 cnt : integer;
 mdate : TDateTime;
begin

 DecodeDate(firstd,year,month,day);
 firstd := EncodeDate(year,month,day);
 DecodeDate(lastd,yearl,monthl,dayl);
 lastd := EncodeDate(yearl,monthl,dayl);

 if firstd > lastd then
 begin
   mdate := firstd;
   firstd := lastd;
   lastd := mdate;
 end;

 for cnt := 0 to dl.BigDataGrid.Count - 1 do
 begin
  if dl.saved then
  begin
   if TDataContent(dl.BigDataGrid.Items[cnt]).dt <> 0 then
    mdate := TDataContent(dl.BigDataGrid.Items[cnt]).dt;
  end
  else
  begin
  if cnt=0 then
   mdate := DateTimePicker1.Date
  else
  if (TDataContent(dl.BigDataGrid.Items[cnt]).TrainLabel<>nil) then
   mdate := TDataContent(dl.BigDataGrid.Items[cnt]).TrainDate.Date;
  end;

  DecodeDate(mdate,yeart,montht,dayt);
  mdate := EncodeDate(yeart,montht,dayt);

  if mdate > lastd then
   lastd := mdate;
  if mdate <= firstd then
   firstd := mdate;
 end;

 DecodeDate(firstd,year,month,day);
 firstd := EncodeDate(year,month,day);
 DecodeDate(lastd,yearl,monthl,dayl);
 lastd := EncodeDate(yearl,monthl,dayl);

 if (year = yearl) and (month=monthl) and (day=dayl) then
  Result := DateToStr(firstd) + ' (' + IntToStr(dl.trcount) + ')'
 else
  Result := DateToStr(firstd) + '-' +  DateToStr(lastd)  + ' (' + IntToStr(dl.trcount) + ')';
end;


function TMDIChild.AddMicroCicle(T : TTreeNode; s : string): TTreeNode;
var
 MN, RN : TTreeNode;
 lastcicle, firstcicle,k : integer;
 found : boolean;
begin
  k := StrToInt(copy(s,1,1));
  if T.HasChildren then
  begin
   MN := T.getFirstChild;
   firstcicle := StrToInt(copy(MN.Text,1,1));
   lastcicle := StrToInt(copy(MN.Text,1,1));
   found := false;
   if s = MN.Text then
    found := true
   else
   while MN<>nil do
    begin
     lastcicle := StrToInt(copy(MN.Text,1,1));
     if s = MN.Text then
     begin
      found := true;
      break;
     end;
//     if k>firstcicle then
//      firstcicle := k;
     if lastcicle>k then
      break;
     RN := MN;
     MN := T.GetNextChild(MN);
    end;

   if not found then
   begin
    if k<firstcicle then
     MN := TreeView1.Items.AddChildFirst(T,s)
    else
    if k>lastcicle then
     MN := TreeView1.Items.AddChild(T,s)
    else
    if (k>firstcicle) and (k<lastcicle) then
     MN := TreeView1.Items.Insert(MN,s)
    else
     MN := TreeView1.Items.AddChild(T,s);
   end;
  end
  else
   MN := TreeView1.Items.AddChild(T,s);

  MN.ImageIndex := 0;
  MN.SelectedIndex := 0;
  Result := MN;
end;

function TMDIChild.AddTreeData(dl : TDataLeaf):TDataLeaf;
var
 firstd, lastd : TDateTime;
 year, month, day : word;
 yearl, monthl, dayl : word;
 yeart, montht, dayt : word;
 smonth,dds,ddi : string;
 YN, DN, MN, TN : TTreeNode;
 found, foundyear,foundmonth,foundday : boolean;
 firstcicle, lastcicle, cnt,i,j,ii,fm,fy,fd : integer;
 dl1 : TdataLeaf;

begin
 if not dl.saved then
 begin
  firstd := DateTimePicker1.Date;
  lastd := DateTimePicker1.Date;
  dds := GetLeafInfo(dl,firstd,lastd);
  dl.dt := firstd;
  dl.dtl := lastd;
  ddi := DockTabSet1.Tabs[DockTabSet1.TabIndex];
 end
 else
 begin
  firstd := dl.dt;
  lastd := dl.dtl;
  dds := GetLeafInfo(dl,firstd,lastd);
  ddi := IntToStr(dl.micro) + ' Микроцикл';
 end;

 YN := nil;
 MN := nil;
 DN := nil;

 if not dl.saved then
 case TreeView1.Selected.Level of
   0 : begin
        YN := TreeView1.Selected;
        MN := AddMicroCicle(YN,ddi);
   end;
   1: begin
       YN := TreeView1.Selected.Parent;
       MN := AddMicroCicle(YN,ddi);
   end;
   2: begin
       YN := TreeView1.Selected.Parent.Parent;
       MN := AddMicroCicle(YN,ddi);
   end;
 end
 else
  begin
    YN := AddMezoCicle(dl.mezo,false);
    MN := AddMicroCicle(YN,ddi);
  end;

 foundday := false;

 for i := 0 to TreeView1.Items.Count - 1 do
    begin
    if TreeView1.Items[i].Level = 2 then
     if (TDataLeaf(TreeView1.Items[i].data).dt = firstd) and
     (TDataLeaf(TreeView1.Items[i].data).dtl = lastd) then
      begin
       DN := TreeView1.Items[i];
       foundday := true;
       break;
      end;
    end;

 if FoundDay then
 begin

{    case MessageBOX(Handle, PChar('Найдены пересечения временных интервалов в базе тренировок. ' +
      ' Сохранить некорректные пересечения? (Внимание! Могут быть потеряны данные.)'), PChar(MainTitle) ,
      MB_YESNO or MB_ICONQUESTION) of
      IDYes:
        begin  }
         TN := MN;
         TN.ImageIndex := 0;
         TN.SelectedIndex := 0;
         foundday := false;

{         dl1 := TDataLeaf(DN.data);
         TreeView1.Selected := DN;
         if dl1.dt = dl1.dtl then
          TreeView1.Selected.Text := DateToStr(dl1.dt) + ' (' + IntToStr(dl1.trcount) + ')'
         else
          TreeView1.Selected.Text := DateToStr(dl1.dt) + '-' + DateToStr(dl1.dtl) + ' (' + IntToStr(dl1.trcount) + ')';

        end;
      IDNo:
        begin
         exit;
        end;
     end; }

 end
 else
 begin
  TN := MN;
  TN.ImageIndex := 0;
  TN.SelectedIndex := 0;
 end;

 if not foundday then
 begin
      DN := TN.getLastChild;
      if DN<>nil then
      begin
       if TDataLeaf(DN.data).dtl < firstd then
        TN := TreeView1.Items.AddChild(TN,dds)
       else
       begin
       DN := TN.getFirstChild;
       if DN=nil then
        TN := TreeView1.Items.AddChildFirst(TN,dds)
       else
       begin
        while TDataLeaf(DN.data).dtl-1 < firstd do
         begin
         if TN.GetNextChild(DN)=nil then
            break;
          DN := TN.GetNextChild(DN);
         end;
        TN := TreeView1.Items.Insert(DN,dds);
       end;
       end;
      end
      else
        TN := TreeView1.Items.AddChildFirst(TN,dds);

  TN.ImageIndex := 2;
  TN.SelectedIndex := 2;
  TN.Data := dl;
  dl1 := dl;
 end;

 Result := dl1;
end;


function TMDIChild.SaveDataBase : boolean;
var
 i,j : integer;
 s : string;
 k : real;
 dl : TDataLeaf;
 TN : TTreeNode;
 dt1, dt2 : TDateTime;
begin
 if not MainForm.IfRegOK then Exit;

 errfound := false;
 if TreeView1.Selected = nil then
  begin
   errfound := true;
   MessageBOX(Handle, PChar('Укажите мезоцикл...'), PChar(Caption) ,
   MB_ICONWARNING)
  end
{ else
 if (TreeView1.Selected <> nil) and (TreeView1.Selected.Level>0) then
  begin
   errfound := true;
   MessageBOX(Handle, PChar('Укажите мезоцикл...'), PChar(Caption) ,
   MB_ICONWARNING)
  end }
 else
 if ComboBox1.ItemIndex = -1 then
  begin
   errfound := true;
   MessageBOX(Handle, PChar('Укажите фамилию спортсмена...'), PChar(Caption) ,
   MB_ICONWARNING)
  end
 else
 if ComboBox1.ItemIndex = -1 then
  begin
   errfound := true;
   MessageBOX(Handle, PChar('Укажите вид спорта...'), PChar(Caption) ,
   MB_ICONWARNING)
  end
 else
 for i := 0 to CurrentLeaf.BigDataGrid.Count - 1 do
 begin
  if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.ItemIndex = -1 then
   begin
    errfound := true;
    MessageBOX(Handle, PChar('Укажите упражнение...'), PChar(Caption) ,
    MB_ICONWARNING);
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).ComboBox1.SetFocus;
    break;
   end;
 end;

 if not errfound then
 begin
  Modified := false;
  if not CurrentLeaf.saved then
   dl := AddTreeData(CurrentLeaf)
  else
  begin
   dl := CurrentLeaf;
   TN := GetCurrentTreeNode(dl);
   dt1 := DateTimePicker1.Date;
   dt2 := dt1;
   dl.saved := false;
   TN.Text := GetLeafInfo(dl, dt1, dt2);
   dl.dt := dt1;
   dl.dtl := dt2;
  end;

  UpdateDataLeaf(dl);

  TN := GetCurrentTreeNode(CurrentLeaf);
  if TN <> nil then
  begin
   try
    // Определим мезо и микро циклы для тренировок
    CurrentLeaf.mezo := StrToInt(copy(TN.Parent.Parent.Text,1,1));
    CurrentLeaf.micro := StrToInt(copy(TN.Parent.Text,1,1));
   except
   end;
   TN.MakeVisible;
   TreeView1.Selected := TN;
  end;

  // Сохраним все
  SaveAll(fname);
  result := true;
 end
 else
  result := false;
end;

procedure TMDIChild.SaveNewLeaf(fname : string; dl : TDataLeaf);
var
 F : TFileStream;
 i,j,len : integer;
 buffer : PAnsiChar;
 pbuffer : PByte;
 s : string;
 M : TMemoryStream;
 dt : TDateTime;
begin
  try
   if not FileExists(fname) then
   begin
    F := TFileStream.Create(fname,fmCreate);
    len := DBVersion;
    F.Write(len,4);
   end
   else
   begin
    F := TFileStream.Create(fname,fmOPenReadWrite);
    F.Seek(0,soFromEnd);
   end;
   M := TMemoryStream.Create;

   dt := dl.dt;
   M.Write(dt,8);

   s := dl.men;
   len := Length(s);
   M.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   M.Write(buffer^,len);
   FreeMem(buffer);

   s := dl.sport;
   len := Length(s);
   M.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   M.Write(buffer^,len);
   FreeMem(buffer);

   len := dl.BigDataGrid.Count;
   M.Write(len,4);
   for i := 0 to dl.BigDataGrid.Count - 1 do
   begin
    s := TDataContent(dl.BigDataGrid.Items[i]).ComboBox1.Text;
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);
    len := TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Count;
    M.Write(len,4);
    for j := 0 to TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Count - 1 do
    begin
     s := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1.Text;
     len := Length(s);
     M.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     M.Write(buffer^,len);
     FreeMem(buffer);
     s := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2.Text;
     len := Length(s);
     M.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     M.Write(buffer^,len);
     FreeMem(buffer);
     s := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit3.Text;
     len := Length(s);
     M.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     M.Write(buffer^,len);
     FreeMem(buffer);
    end;
   end;

   M.Position := 0;
   len := M.Size;
   F.Write(len,4);
   F.CopyFrom(M,M.Size);

   F.Free;
   M.Free;
  except
   MessageBOX(Handle, PChar('Ошибка записи файла "'+fname+'"...'), Pchar(MainTitle) ,
   MB_ICONERROR);
  end;
end;

procedure TMDIChild.ReadManStream(fname : string; var ManStream : TMemoryStream);
var
 F : TFileStream;
 len : integer;
begin
  try
   F := TFileStream.Create(fname,fmOpenReadWrite);
   F.Read(len,4);
   F.Read(len,4);
   ManStream.Clear;
   ManStream.CopyFrom(F,len);
   F.Free;
  except
   MessageBOX(Handle, PChar('Ошибка при чтении файла "'+fname+'"...'), Pchar(MainTitle) ,
   MB_ICONERROR);
  end;
end;

procedure TMDIChild.SaveAll(fname : string);
var
 F : TFileStream;
 ii,i,j,len : integer;
 buffer : PAnsiChar;
 pbuffer : PByte;
 s : string;
 M : TMemoryStream;
 dt : TDateTime;
 dl : TDataLeaf;
begin
  try
   Screen.Cursor := crHourGlass;

   F := TFileStream.Create(fname,fmOpenReadWrite);
   F.Read(len,4);
   F.Read(len,4);
   ManStream.Clear;
   ManStream.CopyFrom(F,len);
   F.Free;


   F := TFileStream.Create(fname,fmCreate);
   len := DBVersion;
   F.Write(len,4);
   len := ManStream.Size;
   F.Write(len,4);
   ManStream.Position := 0;
   F.CopyFrom(ManStream,len);

//   for II := 0 to TreeView1.Items.Count - 1 do
//   if TreeView1.Items[ii].Level=2 then
   for II := 0 to BigDataLeaf.Count - 1 do
   if TDataLeaf(BigDataLeaf.Items[ii]).menindex<>-1 then
   begin

//   dl := TDataLeaf(TreeView1.Items[ii].Data);
   dl := TDataLeaf(BigDataLeaf.Items[ii]);

   dl.saved := true;
   M := TMemoryStream.Create;
   SaveDL(dl,M);
   M.Position := 0;
   len := M.Size;
   F.Write(len,4);
   F.CopyFrom(M,M.Size);
   M.Free;

   end;


   F.Free;
   Screen.Cursor := crDefault;
  except
   Screen.Cursor := crDefault;
   MessageBOX(Handle, PChar('Ошибка записи файла "'+fname+'"...'), Pchar(MainTitle) ,
   MB_ICONERROR);
  end;
end;

procedure TMDIChild.SaveDL(dl : TDataLeaf; var M : TMemoryStream);
var
 ii,i,j,len : integer;
 buffer : PAnsiChar;
 pbuffer : PByte;
 s : string;
 dt : TDateTime;
begin

   dt := dl.dt;
   M.Write(dt,8);

   // !!!!!!!!!!!!!!!!!!!!!!!!!!
   dt := dl.dtl;
   M.Write(dt,8);
   // !!!!!!!!!!!!!!!!!!!!!!!!!!

   M.Write(dl.menindex,4);
   s := dl.men;
   len := Length(s);
   M.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   M.Write(buffer^,len);
   FreeMem(buffer);

   M.Write(dl.sportindex,4);
   s := dl.sport;
   len := Length(s);
   M.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   M.Write(buffer^,len);
   FreeMem(buffer);

   // !!!!!!!!!!!!!!!!!!!!!!!!!!
   M.Write(dl.trcount,4);
   M.Write(dl.mezo,4);
   M.Write(dl.micro,4);
   // !!!!!!!!!!!!!!!!!!!!!!!!!!

   len := dl.BigDataGrid.Count;
   M.Write(len,4);
   for i := 0 to dl.BigDataGrid.Count - 1 do
   begin
    // !!!!!!!!!!!!!!!!!!!!!!!!!!
    M.Write(TDataContent(dl.BigDataGrid.Items[i]).TrIndex,4);
    dt := TDataContent(dl.BigDataGrid.Items[i]).dt;
    M.Write(dt,8);
    // !!!!!!!!!!!!!!!!!!!!!!!!!!

    M.Write(TDataContent(dl.BigDataGrid.Items[i]).ExecIndex,4);
    s := TDataContent(dl.BigDataGrid.Items[i]).Exec;
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);

    M.Write(TDataContent(dl.BigDataGrid.Items[i]).IsOfp,4);
    M.Write(TDataContent(dl.BigDataGrid.Items[i]).OfpIndex,4);
    s := TDataContent(dl.BigDataGrid.Items[i]).OFP;
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);

    len := TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Count;
    M.Write(len,4);
    for j := 0 to TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Count - 1 do
    begin
     s := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Res1;
     len := Length(s);
     M.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     M.Write(buffer^,len);
     FreeMem(buffer);
     s := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Res2;
     len := Length(s);
     M.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     M.Write(buffer^,len);
     FreeMem(buffer);
     s := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Res3;
     len := Length(s);
     M.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     M.Write(buffer^,len);
     FreeMem(buffer);
    end;
   end;

   // Восстановление
   len := dl.RelaxList.Count;
   M.Write(len,4);
   for i := 0 to dl.RelaxList.Count - 1 do
   begin
     s := dl.RelaxList[i];
     len := Length(s);
     M.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     M.Write(buffer^,len);
     FreeMem(buffer);
   end;

   // Контроль
   len := dl.ControlList.Count;
   M.Write(len,4);
   for i := 0 to dl.ControlList.Count - 1 do
   begin
     s := dl.ControlList[i];
     len := Length(s);
     M.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     M.Write(buffer^,len);
     FreeMem(buffer);
   end;

end;

procedure TMDIChild.UpdateDataLeaf(dl : TDataLeaf);
var
 i,j : integer;
 k : real;
 year,month,day : word;
begin
 dl.saved := true;
 dl.men := ComboBox1.Text;
 dl.menindex := ComboBox1.ItemIndex;
 dl.sport := ComboBox2.Text;
 dl.sportindex := ComboBox2.ItemIndex;

 dl.ControlList.Clear;
 for i := 1 to vle.RowCount - 1 do
  dl.ControlList.Add(vle.Cells[2,i]);

 dl.RelaxList.Clear;
 for i := 0 to ListView1.Items.Count - 1 do
  if ListView1.Items[i].Checked then
   dl.RelaxList.Add(ListView1.Items[i].Caption);

 for i := 0 to dl.BigDataGrid.Count - 1 do
 begin

  if i=0 then
   begin
    TDataContent(dl.BigDataGrid.Items[i]).dt := DateTimePicker1.Date;
    DecodeDate(TDataContent(dl.BigDataGrid.Items[i]).dt,year,month,day);
    TDataContent(dl.BigDataGrid.Items[i]).dt := EncodeDate(year,month,day);
   end
  else
  if TDataContent(dl.BigDataGrid.Items[i]).TrainDate <> nil then
   begin
    TDataContent(dl.BigDataGrid.Items[i]).dt := TDataContent(dl.BigDataGrid.Items[i]).TrainDate.Date;
    DecodeDate(TDataContent(dl.BigDataGrid.Items[i]).dt,year,month,day);
    TDataContent(dl.BigDataGrid.Items[i]).dt := EncodeDate(year,month,day);
   end;

  if i=0 then
   TDataContent(dl.BigDataGrid.Items[i]).trainkind := ComboBox3.ItemIndex
  else
   if TDataContent(dl.BigDataGrid.Items[i]).TrainBox <> nil then
    TDataContent(dl.BigDataGrid.Items[i]).trainkind := TDataContent(dl.BigDataGrid.Items[i]).TrainBox.ItemIndex;

  if TDataContent(dl.BigDataGrid.Items[i]).IsOFP = 1 then
  begin
   TDataContent(dl.BigDataGrid.Items[i]).OFP := TDataContent(dl.BigDataGrid.Items[i]).ComboBox1.Text;
   TDataContent(dl.BigDataGrid.Items[i]).OFPIndex := TDataContent(dl.BigDataGrid.Items[i]).ComboBox1.ItemIndex;
  end
  else
  begin
   TDataContent(dl.BigDataGrid.Items[i]).Exec := TDataContent(dl.BigDataGrid.Items[i]).ComboBox1.Text;
   TDataContent(dl.BigDataGrid.Items[i]).ExecIndex := TDataContent(dl.BigDataGrid.Items[i]).ComboBox1.ItemIndex;
  end;

  for j := 0 to TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Count - 1 do
   begin
    TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Res1 := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1.Text;
    TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Res2 := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2.Text;
    TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Res3 := TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit3.Text;

     try
      k := StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1.Text);
     except
      MessageBOX(Handle, PChar('Невозможно преобразовать выражение в числовой формат...'), PChar(Caption) ,
      MB_ICONWARNING);
      TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1.SetFocus;
     end;

     try
      k := StrToFloat(TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2.Text);
     except
      MessageBOX(Handle, PChar('Невозможно преобразовать выражение в числовой формат...'), PChar(Caption) ,
      MB_ICONWARNING);
      TDataPoint(TDataContent(dl.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2.SetFocus;
     end;

   end;
 end;

end;

procedure TMDIChild.vleClick(Sender: TObject);
begin
 Modified := true;
end;

procedure TMDIChild.CheckClick(Sender: TObject);
var
 Panel1 : TPanel;
 i,j : integer;
 found : boolean;
begin
    found := false;
    j := 0;
    for i := 0 to CurrentLeaf.BigDataGrid.Count - 1 do
    begin
     if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).CheckBox1.Parent.Top = TCheckBox(Sender).Parent.Top then
       begin
        found := true;
        break;
       end;
    end;
    if not found then i := CurrentLeaf.BigDataGrid.Count - 1;

    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).CheckBox1.Checked then
    begin
     while TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Count>1 do
     begin
      Dec(GlobalX);
      j := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Count-1;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Label1<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Label1.Destroy;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1.Destroy;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2.Destroy;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit3<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit3.Destroy;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit31<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit31.Destroy;
      if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Panel<>nil then
       TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Panel.Destroy;

      TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Delete(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Count-1);
      TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Left := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Left - 90;
      TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Left := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Left - 90;
      ScrollBox1.ScrollInView(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev);
     end;
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Visible := False;
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Visible := False;
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).IsOFP := 1;
     TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[0]).Label1.Visible := false;
     TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit1.Width := 48+33;
     TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit2.Width := 48+33;
     TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit3.Visible := false;
     TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit31.Visible := false;
     ReadOFPFromList;

    end
    else
    begin
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Visible := True;
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Visible := True;
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).IsOFP := 0;
     TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[0]).Label1.Visible := true;
     TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit1.Width := 48;
     TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit2.Width := 48;
     TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit3.Visible := true;
     TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[0]).Edit31.Visible := true;
     ReadExecsFromList;

    end;
    RecountLabels;
    Modified := true;
end;

procedure TMDIChild.RecountLabels;
var
 i, j, k, jj, kold: integer;
begin
    j:=0;
    jj := 0;
    kold := 1;
    k := 1;

    k := TDataContent(CurrentLeaf.BigDataGrid.Items[0]).trindex;
    if TDataContent(CurrentLeaf.BigDataGrid.Items[0]).IsOFP = 1 then
     begin
      inc(j);
      TDataContent(CurrentLeaf.BigDataGrid.Items[0]).Label1.Font.Color := clRed;
      TDataContent(CurrentLeaf.BigDataGrid.Items[0]).Label1.Caption := IntToStr(j) + ' ОФП:';
     end
    else
     begin
      inc(jj);
      TDataContent(CurrentLeaf.BigDataGrid.Items[0]).Label1.Font.Color := clBlack;
      TDataContent(CurrentLeaf.BigDataGrid.Items[0]).Label1.Caption := IntToStr(jj) + ' упражнение:';
     end;
    kold := k;

    for i := 1 to CurrentLeaf.BigDataGrid.Count - 1 do
    begin
     k := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).trindex;
     if kold <> k then
     begin
       j:=0;
       jj:=0;
     end;
     if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).IsOFP = 1 then
     begin
      inc(j);
      TDataContent(CurrentLeaf.BigDataGrid.Items[i]).Label1.Font.Color := clRed;
      TDataContent(CurrentLeaf.BigDataGrid.Items[i]).Label1.Caption := IntToStr(j) + ' ОФП:';
     end
     else
     begin
      inc(jj);
      TDataContent(CurrentLeaf.BigDataGrid.Items[i]).Label1.Font.Color := clBlack;
      TDataContent(CurrentLeaf.BigDataGrid.Items[i]).Label1.Caption := IntToStr(jj) + ' упражнение:';
     end;
     kold := k;
    end;
end;

procedure TMDIChild.ClickNext(Sender: TObject);
var
 Panel1 : TPanel;
 i,j : integer;
 found : boolean;
begin
    Modified := true;
    found := false;
    for i := 0 to CurrentLeaf.BigDataGrid.Count - 1 do
    begin
     if (TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Top = TBitBtn(Sender).Top)
     and (TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Left = TBitBtn(Sender).Left) then
       begin
        found := true;
        break;
       end;
    end;
    if not found then i := CurrentLeaf.BigDataGrid.Count - 1;


    Panel1 := TPanel.Create(ScrollBox1);
    Panel1.Parent := ScrollBox1;
    Panel1.Width := 85;
    Panel1.Height := 59;
    Panel1.Top := TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Last).Panel.Top;
    Panel1.Left := TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Last).Panel.Left + 90;
    Panel1.Visible := true;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Left := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Left + 90;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Visible := true;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Left := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Left + 90;
    TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Add(TDataPoint.Create);
    TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Last).Panel := Panel1;
    Inc(GlobalX);
    TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Last).X := GlobalX;
    TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Last).InitVisual(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Count);
    ScrollBox1.ScrollInView(Panel1);
end;

procedure TMDIChild.ClickPrev(Sender: TObject);
var
 i,j : integer;
 found : boolean;
begin
    Modified := true;
    found := false;
    for i := 0 to CurrentLeaf.BigDataGrid.Count - 1 do
    begin
     if (TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Top = TBitBtn(Sender).Top)
     and (TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Left = TBitBtn(Sender).Left) then
       begin
        found := true;
        break;
       end;
    end;
    if not found then i := CurrentLeaf.BigDataGrid.Count - 1;

    if TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Count>1 then
    begin
     Dec(GlobalX);
     j := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Count-1;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Label1<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Label1.Destroy;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit1.Destroy;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit2.Destroy;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit3<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit3.Destroy;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit31<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Edit31.Destroy;
     if TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Panel<>nil then
      TDataPoint(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Items[j]).Panel.Destroy;
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Delete(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).DataGrid.Count-1);
     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Left := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev.Left - 90;

     TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Left := TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnNext.Left - 90;
     ScrollBox1.ScrollInView(TDataContent(CurrentLeaf.BigDataGrid.Items[i]).BitBtnPrev);
    end;
end;

procedure TMDIChild.ComboBox1Change(Sender: TObject);
var
 i : integer;
 found : boolean;
 TN : TTreeNode;
 CL : TDataLeaf;
begin
// if CurrentLeaf<>nil then
//  ClearPanel;
 Screen.Cursor := crHourGlass;

 found := false;
 TreeView1.Items.Clear;
// AddMezoCicle;

 for I := 0 to BigDataLeaf.Count - 1 do
  if ComboBox1.Text = TDataLeaf(BigDataLeaf.Items[i]).men then
   if TDataLeaf(BigDataLeaf.Items[i]).saved then
    begin
     AddTreeData(TDataLeaf(BigDataLeaf.Items[i]));
     CL := TDataLeaf(BigDataLeaf.Items[i]);
     found := true;
    end;

// TreeView1.Items.AlphaSort(false);

 if found then
 begin
  TN := GetCurrentTreeNode(CL);
  if TN <> nil then
   begin
    TN.MakeVisible;
    TreeView1.Selected := TN;
   end;
//  ShowLeaf(CurrentLeaf,false)
 end
 else
 begin
   ClearPanel;
   InitFirst;
   ComboBox2.Items.Clear;
   tManRec := MenList.Items[ComboBox1.ItemIndex];
   ComboBox2.Items.AddStrings(tManRec^.SportList);
   if ComboBox2.Items.Count>0 then
    ComboBox2.ItemIndex := 0;
   Label3.Caption := ComboBox2.Text;
   ReadExecs(true);
 end;
 Screen.Cursor := crDefault;

end;

procedure TMDIChild.ComboBox2Change(Sender: TObject);
begin
 ReadExecs(true);
end;

procedure TMDIChild.DateTimePicker1CloseUp(Sender: TObject);
var
 i : integer;
 found : boolean;
begin
{ found := false;
 for I := 0 to TreeView1.Items.Count - 1 do
  if TreeView1.Items[i].Level=2 then
   if TDataLeaf(TreeView1.Items[i].Data).dt = DateTimePicker1.Date then
    if TDataLeaf(TreeView1.Items[i].Data).men = ComboBox1.Text then
     begin
      found := true;
      TreeView1.Selected := TreeView1.Items[i];
      TreeView1.Selected.MakeVisible;
      PageControl1.ActivePageIndex := 0;
//      ShowLeaf(TDataLeaf(TreeView1.Selected.Data),true);
      break;
     end;

 if TreeView1.Items.Count>5 then
  if not MainForm.IfRegOK then Exit;

 if not found then
 begin
  ClearPanel;
  InitFirst;
  ReadExecs(true);
 end;}
end;

procedure TMDIChild.DateTimePicker1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 if key = VK_RETURN then
  DateTimePicker1CloseUp(Sender);
end;

procedure TMDIChild.InitFirst;
var
 Panel1, Panel2 : TPanel;
 i : integer;
begin
  MainForm.SB.Panels[0].Text := 'Новая тренировка';

  PageControl1.ActivePageIndex := 0;

  Modified := true;

  BigDataLeaf.Add(TDataLeaf.Create);
  CurrentLeaf := TDataLeaf(BigDataLeaf.Last);

  TDataLeaf(BigDataLeaf.Last).BigDataGrid := TObjectList.Create;
  CurrentLeaf.dt := DateTimePicker1.Date;
  CurrentLeaf.sportindex := -1;
  CurrentLeaf.menindex := -1;
  CurrentLeaf.saved := false;
  Inc(CurrentLeaf.trcount);

  for I := 0 to ListView1.Items.Count - 1 do
   ListView1.Items[i].Checked := false;

  GlobalX := 1;
  GlobalY := 0;

  Panel1 := TPanel.Create(ScrollBox1);
  Panel1.Parent := ScrollBox1;
  Panel1.Width := 262;
  Panel1.Height := 59;
  Panel1.Top := 26;
  Panel1.Left := 28;
  Panel1.Visible := true;
  TDataLeaf(BigDataLeaf.Last).BigDataGrid.Add(TDataContent.Create);
  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).trindex := CurrentLeaf.trcount;
  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).Panel := Panel1;
  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).ScrollBox := ScrollBox1;
  Inc(GlobalY);
  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).X := GlobalX;
  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).Y := GlobalY;
  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).InitVisual(26);
  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).BitBtnNext.OnClick := ClickNext;
  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).BitBtnPrev.OnClick := ClickPrev;
  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).CheckBox1.OnClick := CheckClick;
//  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).BitBtnPrev.Visible := false;

  Panel2 := TPanel.Create(ScrollBox1);
  Panel2.Parent := ScrollBox1;
  Panel2.Width := 85;
  Panel2.Height := 59;
  Panel2.Top := 26;
  Panel2.Left := 295;
  Panel2.Visible := true;
  TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).DataGrid.Add(TDataPoint.Create);
  TDataPoint(TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).DataGrid.Last).Panel := Panel2;
  TDataPoint(TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).DataGrid.Last).X := GlobalX;
  TDataPoint(TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).DataGrid.Last).InitVisual(1);

  ImageList1.GetBitmap(2,TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).BitBtnPrev.Glyph);
  ImageList1.GetBitmap(3,TDataContent(TDataLeaf(BigDataLeaf.Last).BigDataGrid.Last).BitBtnNext.Glyph);

  Button2.Top := 140;
  BitBtnUp.Top := 25;
  BitBtnDown.Top := 55;

  Label5.Caption := '';
end;

procedure TMDIChild.ListView1Click(Sender: TObject);
begin
 Modified := true;
end;

procedure TMDIChild.N1Click(Sender: TObject);
var
 ii : integer;
 found : boolean;
begin
 if TreeView1.Selected <> nil then
  if TreeView1.Selected.Level = 2 then
   begin
    if MessageBOX(Handle, PChar('Удалить тренировку '+TreeView1.Selected.Text+'?'), PChar(MainTitle) ,
    MB_YESNO or MB_ICONQUESTION) = IDYes then
     begin
      found := false;
      for II := 0 to BigDataLeaf.Count - 1 do
      if TDataLeaf(BigDataLeaf.Items[ii]).menindex<>-1 then
       begin
        if TDataLeaf(BigDataLeaf.Items[ii]) = TDataLeaf(TreeView1.Selected.Data) then
         begin
           found := true;
           break;
         end;
       end;
      if found then
       begin
        Modified := true;
        CurrentLeaf := TDataLeaf(BigDataLeaf.Items[ii]);
        ClearPanel;
        BigDataLeaf.Delete(ii);
        DeleteItem := True;
        TreeView1.Selected.Delete;
        DeleteItem := False;
        if TreeView1.Selected.Level = 2 then
//         ShowTreeNode(TreeView1.Selected)
        else
        begin
         InitFirst;
         ReadExecs(true);
        end;
       end;
     end;
   end;
end;

procedure TMDIChild.N2Click(Sender: TObject);
var
 i, k : integer;
 found : boolean;
begin
 if TreeView1.Selected <> nil then
  if (TreeView1.Selected.Level = 0) and (TreeView1.Items.Count>1) then
   begin
    if MessageBOX(Handle, PChar('Удалить мезоцикл "'+TreeView1.Selected.Text+'"?'), PChar(MainTitle) ,
    MB_YESNO or MB_ICONQUESTION) = IDYes then
     begin
        // Проверим присутсвие тренировок в мезоцикле
        found := false;
        for i := 0 to TreeView1.Selected.Count - 1 do
        begin
         found := TreeView1.Selected.Item[i].HasChildren;
         if found then break;
        end;
        if not found then
        begin
         TreeView1.Selected.Delete;
         // Обновим номера мезоциклов
         k := 1;
         for i := 0 to TreeView1.Items.Count - 1 do
         if TreeView1.Items[i].Level=0 then
         begin
          TreeView1.Items[i].Text := IntToStr(k) + ' Мезоцикл';
          Inc(k);
         end;
        end
        else
         MessageBOX(Handle, PChar('Невозможно удалить мезоцикл. В мезоцикле присутствуют тренировки...'), PChar(Caption) ,
         MB_ICONWARNING)

     end;
   end;
end;

procedure TMDIChild.N4Click(Sender: TObject);
begin
 SuperClearPanel;
end;

procedure TMDIChild.SuperClearPanel;
begin
  if DateForm.ShowModal = mrOk then
  begin
   ClearPanel;
   InitFirst;
   ReadExecs(true);
   DateTimePicker1.Date := DateForm.dt.Date;
  end;
end;

procedure TMDIChild.DelExecExecute(Sender: TObject);
begin
 BitBtnUpClick(Sender);
end;

procedure TMDIChild.DelTrainningExecute(Sender: TObject);
begin
 SuperClearPanel;
end;

procedure TMDIChild.DelZoneExecute(Sender: TObject);
begin
 ClickPrev(Sender);
end;

procedure TMDIChild.DockTabSet1Change(Sender: TObject; NewTab: Integer;
  var AllowChange: Boolean);
var
 TN, MN, RN : TTreeNode;
begin
 if not DockChanged then
 begin
  ClearPanel;
  InitFirst;
  ReadExecs(true);
  if TreeView1.Selected <> nil then
  begin
      TN := TreeView1.Selected;
      TN.Selected := false;

      if TN.Level = 2 then
       TN := TN.Parent.Parent;
      if TN.Level = 1 then
       TN := TN.Parent;

      TN.Selected := true;
      MN := TN.getFirstChild;
      while MN<>nil do
      begin
       if MN.Text = DockTabSet1.Tabs[NewTab] then
       begin
        if MN.HasChildren then
        begin
         RN := MN.GetLastChild;
         ShowTreeNode(RN);
         RN.Selected := true;
        end;
        break;
       end;
       MN := TN.GetNextChild(MN);
      end;
  end;
 end;
end;

end.
