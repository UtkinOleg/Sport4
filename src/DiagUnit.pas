unit DiagUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  ComObj, Dialogs, ComCtrls, StdCtrls, ExtCtrls, ToolWin, ImgList, Buttons;

type
  TDiagForm = class(TForm)
    pc: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    ListView1: TListView;
    ListView2: TListView;
    Button1: TButton;
    Button2: TButton;
    CancelBtn: TButton;
    Label2: TLabel;
    ListView3: TListView;
    Button3: TButton;
    ListView4: TListView;
    ListView5: TListView;
    ListView6: TListView;
    Button4: TButton;
    Button5: TButton;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    ButtonNext: TButton;
    ButtonBack: TButton;
    TabSheet3: TTabSheet;
    GroupBox1: TGroupBox;
    ImageList1: TImageList;
    ControlBar1: TControlBar;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ImageList2: TImageList;
    ControlBar3: TControlBar;
    ToolBar3: TToolBar;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ListView7: TListView;
    Label8: TLabel;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    RadioButton6: TRadioButton;
    RadioButton10: TRadioButton;
    ComboBox2: TComboBox;
    ComboBox3: TComboBox;
    Label6: TLabel;
    Label9: TLabel;
    Bevel1: TBevel;
    Button9: TButton;
    Memo1: TMemo;
    Label1: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Memo2: TMemo;
    Label12: TLabel;
    Memo3: TMemo;
    Label13: TLabel;
    Memo4: TMemo;
    TabSheet4: TTabSheet;
    GroupBox4: TGroupBox;
    Label14: TLabel;
    Edit1: TEdit;
    BitBtn1: TBitBtn;
    TabSheet5: TTabSheet;
    Memo5: TMemo;
    Label15: TLabel;
    GroupBox5: TGroupBox;
    Label16: TLabel;
    RadioButton2: TRadioButton;
    RadioButton7: TRadioButton;
    Label17: TLabel;
    OpenDialog1: TOpenDialog;
    Label18: TLabel;
    MezoB: TEdit;
    Label19: TLabel;
    MicroB: TEdit;
    Label20: TLabel;
    MezoE: TEdit;
    MicroE: TEdit;
    Label21: TLabel;
    Label22: TLabel;
    MezoBox: TComboBox;
    procedure RadioButton2Click(Sender: TObject);
    procedure RadioButton7Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure Label1MouseLeave(Sender: TObject);
    procedure Label1MouseEnter(Sender: TObject);
    procedure ComboBox3Change(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure ListView1Click(Sender: TObject);
    procedure ListView7Click(Sender: TObject);
    procedure ListView3Click(Sender: TObject);
    procedure RadioButton6Click(Sender: TObject);
    procedure RadioButton10Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ToolButton6Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ToolButton4Click(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ButtonBackClick(Sender: TObject);
    procedure ButtonNextClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    FName, ImportExportName : string;
    MasterType : byte;
  end;

var
  DiagForm: TDiagForm;

implementation

{$R *.dfm}

uses Main, Man, RazrForm, ChildWin;

procedure TDiagForm.BitBtn1Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
   Edit1.Text := OpenDialog1.FileName;
end;

procedure TDiagForm.Button1Click(Sender: TObject);
var
 i,j : integer;
 found : boolean;
 LI : TListItem;
begin

 for j := 0 to ListView1.Items.Count - 1 do
 begin
 if ListView1.Items[j].Checked then
  begin
   found := false;
   for i := 0 to ListView2.Items.Count - 1 do
   if ListView2.Items[i].Subitems[0] = ListView1.Items[j].SubItems[0] then
     begin
       found := true;
       break;
     end;
   if (not found) and (ListView2.Items.Count=0) then
   begin
    LI := ListView2.Items.Add;
    LI.Caption := IntToStr(ListView2.Items.Count);
    LI.SubItems.Add(ListView1.Items[j].Subitems[0]);
    break;
   end;
  end;
 end;

end;

procedure TDiagForm.Button2Click(Sender: TObject);
var
 i,j : integer;
 found : boolean;
 LI : TListItem;
begin
 repeat
  found := false;
  for j := 0 to ListView2.Items.Count - 1 do
   if ListView2.Items[j].Selected then
   begin
    found := true;
    break;
   end;
   if found then
    ListView2.Items[j].Delete;
 until not found;
end;


procedure TDiagForm.Button3Click(Sender: TObject);
var
 i,j : integer;
 found : boolean;
 LI : TListItem;
begin

 for j := 0 to ListView3.Items.Count - 1 do
 begin
 if ListView3.Items[j].Checked then
  begin
   found := false;
   for i := 0 to ListView4.Items.Count - 1 do
   if ListView4.Items[i].Subitems[0] = ListView3.Items[j].SubItems[0] then
     begin
       found := true;
       break;
     end;
   if not found then
   begin
    LI := ListView4.Items.Add;
    LI.Caption := IntToStr(ListView4.Items.Count);
    LI.SubItems.Add(ListView3.Items[j].Subitems[0]);
   end;
  end;
 end;
end;

procedure TDiagForm.Button4Click(Sender: TObject);
var
 i,j : integer;
 found : boolean;
 LI : TListItem;
begin

 for j := 0 to ListView3.Items.Count - 1 do
 begin
 if ListView3.Items[j].Checked then
  begin
   found := false;
   for i := 0 to ListView5.Items.Count - 1 do
   if ListView5.Items[i].Subitems[0] = ListView3.Items[j].SubItems[0] then
     begin
       found := true;
       break;
     end;
   if not found then
   begin
    LI := ListView5.Items.Add;
    LI.Caption := IntToStr(ListView5.Items.Count);
    LI.SubItems.Add(ListView3.Items[j].Subitems[0]);
   end;
  end;
 end;
end;

procedure TDiagForm.Button5Click(Sender: TObject);
var
 i,j : integer;
 found : boolean;
 LI : TListItem;
begin

 for j := 0 to ListView7.Items.Count - 1 do
 begin
 if ListView7.Items[j].Checked then
  begin
   found := false;
   for i := 0 to ListView6.Items.Count - 1 do
   if ListView6.Items[i].Subitems[0] = ListView7.Items[j].SubItems[0] then
     begin
       found := true;
       break;
     end;
   if not found then
   begin
    LI := ListView6.Items.Add;
    LI.Caption := IntToStr(ListView6.Items.Count);
    LI.SubItems.Add(ListView7.Items[j].Subitems[0]);
   end;
  end;
 end;

 repeat
  found := false;
  for j := 0 to ListView7.Items.Count - 1 do
   if ListView7.Items[j].Checked then
   begin
    found := true;
    break;
   end;
   if found then
    ListView7.Items[j].Delete;
 until not found;
end;

procedure TDiagForm.Button6Click(Sender: TObject);
var
 i,j : integer;
 found : boolean;
 LI : TListItem;
begin
 repeat
  found := false;
  for j := 0 to ListView4.Items.Count - 1 do
   if ListView4.Items[j].Selected then
   begin
    found := true;
    break;
   end;
   if found then
    ListView4.Items[j].Delete;
 until not found;
end;

procedure TDiagForm.Button7Click(Sender: TObject);
var
 i,j : integer;
 found : boolean;
 LI : TListItem;
begin
 repeat
  found := false;
  for j := 0 to ListView5.Items.Count - 1 do
   if ListView5.Items[j].Selected then
   begin
    found := true;
    break;
   end;
   if found then
    ListView5.Items[j].Delete;
 until not found;
end;

procedure TDiagForm.Button8Click(Sender: TObject);
var
 i,j : integer;
 found : boolean;
 LI : TListItem;
begin

 for j := 0 to ListView6.Items.Count - 1 do
 begin
 if ListView6.Items[j].Selected then
  begin
   found := false;
   for i := 0 to ListView7.Items.Count - 1 do
   if ListView7.Items[i].Subitems[0] = ListView6.Items[j].SubItems[0] then
     begin
       found := true;
       break;
     end;
   if not found then
   begin
    LI := ListView7.Items.Add;
    LI.Caption := IntToStr(ListView7.Items.Count);
    LI.SubItems.Add(ListView6.Items[j].Subitems[0]);
   end;
  end;
 end;

 repeat
  found := false;
  for j := 0 to ListView6.Items.Count - 1 do
   if ListView6.Items[j].Selected then
   begin
    found := true;
    break;
   end;
   if found then
    ListView6.Items[j].Delete;
 until not found;
end;

procedure TDiagForm.ButtonBackClick(Sender: TObject);
begin
if MasterType=0 then
begin
 if pc.ActivePageIndex = 1 then
 begin
  pc.ActivePageIndex := 0;
  ButtonBack.Enabled := false;
 end
 else
 if pc.ActivePageIndex = 2 then
 begin
  pc.ActivePageIndex := 1;
  ButtonNext.Caption := 'Далее >';
 end;
end
else
if MasterType=1 then
begin
 if pc.ActivePageIndex = 3 then
 begin
  pc.ActivePageIndex := 0;
  ButtonBack.Enabled := false;
 end
 else
 if pc.ActivePageIndex = 4 then
 begin
  pc.ActivePageIndex := 3;
  ButtonNext.Caption := 'Далее >';
 end;
end
else
if MasterType=2 then
begin
 if pc.ActivePageIndex = 0 then
 begin
  pc.ActivePageIndex := 3;
  ButtonBack.Enabled := false;
 end
 else
 if pc.ActivePageIndex = 4 then
 begin
  pc.ActivePageIndex := 0;
  ButtonNext.Caption := 'Далее >';
 end;
end
else
if MasterType=3 then
begin
 if pc.ActivePageIndex = 4 then
 begin
  pc.ActivePageIndex := 0;
  ButtonBack.Enabled := false;
  ButtonNext.Caption := 'Далее >';
 end;
end;

end;

procedure TDiagForm.ButtonNextClick(Sender: TObject);
var
 ST : TStringList;
 LI : TListItem;
 i : integer;

begin
if MasterType=0 then
begin
 if pc.ActivePageIndex = 0 then
 begin
  if ListView2.Items.Count = 0 then
   begin
    MessageBOX(Handle, PChar('Необходимо выбрать одного или нескольких спортсменов.'), PChar(MainTitle) ,
    MB_ICONWARNING);
    Exit;
   end;
  pc.ActivePageIndex := 1;
  ButtonBack.Enabled := true;
 end
 else
 if pc.ActivePageIndex = 1 then
 begin
  if (ListView4.Items.Count = 0) and (ListView5.Items.Count = 0) and (ListView6.Items.Count = 0) then
   begin
    MessageBOX(Handle, PChar('Необходимо выбрать одно или несколько упражнений.'), PChar(MainTitle) ,
    MB_ICONWARNING);
    Exit;
   end;
  Memo1.Lines.Clear;
  Memo2.Lines.Clear;
  Memo3.Lines.Clear;
  Memo4.Lines.Clear;
  ST := TStringList.Create;
  for I := 0 to ListView2.Items.Count - 1 do
   ST.Add(ListView2.Items[i].SubItems[0]);
  Memo1.Lines.AddStrings(ST);
  ST.Clear;
  for I := 0 to ListView4.Items.Count - 1 do
   ST.Add(ListView4.Items[i].SubItems[0]);
  Memo2.Lines.AddStrings(ST);
  ST.Clear;
  for I := 0 to ListView5.Items.Count - 1 do
   ST.Add(ListView5.Items[i].SubItems[0]);
  Memo3.Lines.AddStrings(ST);
  ST.Clear;
  for I := 0 to ListView6.Items.Count - 1 do
   ST.Add(ListView6.Items[i].SubItems[0]);
  Memo4.Lines.AddStrings(ST);
  ST.Clear;

  pc.ActivePageIndex := 2;
  ButtonNext.Caption := 'Готово';
 end
 else
 if pc.ActivePageIndex = 2 then
 begin
  // &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
   ModalResult := mrOk;
  // &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
 end;
end
else
if MasterType=1 then
begin
 if pc.ActivePageIndex = 0 then
 begin
  if ListView2.Items.Count = 0 then
   begin
    MessageBOX(Handle, PChar('Необходимо выбрать одного или нескольких спортсменов.'), PChar(MainTitle) ,
    MB_ICONWARNING);
    Exit;
   end;
  pc.ActivePageIndex := 3;
  ButtonBack.Enabled := true;
 end
 else
 if pc.ActivePageIndex = 3 then
 begin
  if Length(Edit1.Text) = 0 then
   begin
    MessageBOX(Handle, PChar('Необходимо указать имя файла.'), PChar(MainTitle) ,
    MB_ICONWARNING);
    Exit;
   end;
  Memo5.Lines.Clear;
  ST := TStringList.Create;
  for I := 0 to ListView2.Items.Count - 1 do
   ST.Add(ListView2.Items[i].SubItems[0]);
  Memo5.Lines.AddStrings(ST);
  ST.Free;
  Label17.Caption := 'Файл экспорта: ' + Edit1.Text;
  ButtonNext.Caption := 'Готово';
  pc.ActivePageIndex := 4;
 end
 else
 if pc.ActivePageIndex = 4 then
 begin
  ImportExportName := Edit1.Text;
  if RadioButton7.Checked then
  try
   i := StrToInt(MezoB.Text);
   i := StrToInt(MezoE.Text);
   i := StrToInt(MicroB.Text);
   i := StrToInt(MicroE.Text);
  except
    MessageBOX(Handle, PChar('Неправильно указаны параметры диапазона.'), PChar(MainTitle) ,
    MB_ICONWARNING);
    Exit;
  end;
  // &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
  ModalResult := mrOk;
  // &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
 end;
end
else
if MasterType=2 then
begin
 if pc.ActivePageIndex = 3 then
 begin
  if Length(Edit1.Text) = 0 then
   begin
    MessageBOX(Handle, PChar('Необходимо указать имя файла.'), PChar(MainTitle) ,
    MB_ICONWARNING);
    Exit;
   end;
 fname := Edit1.Text;

 ListView1.Items.Clear;
 ListView2.Items.Clear;

 ST := TStringList.Create;
 ST.Clear;

 MainForm.ReadExecsToString(ST);
 ComboBox2.Items.Clear;
 ComboBox2.Items.AddStrings(ST);
 ComboBox2.ItemIndex := 0;

 ST.Clear;
 MainForm.ReadExecs2ToString(0,ST);
 ListView3.Items.Clear;
 ListView4.Items.Clear;
 ListView5.Items.Clear;
 ListView6.Items.Clear;
 ListView7.Items.Clear;
 for i := 0 to ST.Count - 1 do
  begin
    LI := ListView3.Items.Add;
    LI.Caption := IntToStr(ListView3.Items.Count);
    LI.SubItems.Add(ST[i]);
  end;

 St.Clear;
 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'ofp.mtx';
 if RazrDlg.GetRazr(ST) then
 begin
  for i := 0 to ST.Count - 1 do
  begin
    LI := ListView7.Items.Add;
    LI.Caption := IntToStr(ListView7.Items.Count);
    LI.SubItems.Add(ST[i]);
  end;
 end;

 St.Clear;
 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'razr.mtx';
 if RazrDlg.GetRazr(ST) then
 begin
  ComboBox3.Items.Clear;
  ComboBox3.Items.Add('Все разряды');
  ComboBox3.Items.AddStrings(ST);
  ComboBox3.ItemIndex := 0;
 end;

 St.Clear;
 ManForm.fname := fname;
 ManForm.ScanDBToStrings(st);
 Label18.Visible := st.Count = 0;
 St.Clear;

 ManForm.ScanDBToStringsIndex(st,ComboBox2.Text,ComboBox3.ItemIndex-1);
 for I := 0 to st.Count - 1 do
 begin
  LI := ListView1.Items.Add;
  LI.Caption := IntToStr(ListView1.Items.Count);
  LI.SubItems.Add(st[i]);
 end;

 ST.Free;

  pc.ActivePageIndex := 0;
  ButtonBack.Enabled := true;
 end
 else
 if pc.ActivePageIndex = 0 then
 begin
  if ListView2.Items.Count = 0 then
   begin
    MessageBOX(Handle, PChar('Необходимо выбрать одного или нескольких спортсменов.'), PChar(MainTitle) ,
    MB_ICONWARNING);
    Exit;
   end;

  Memo5.Lines.Clear;
  ST := TStringList.Create;
  for I := 0 to ListView2.Items.Count - 1 do
   ST.Add(ListView2.Items[i].SubItems[0]);
  Memo5.Lines.AddStrings(ST);
  ST.Free;
  Label17.Caption := 'Файл импорта: ' + Edit1.Text;
  ButtonNext.Caption := 'Готово';
  pc.ActivePageIndex := 4;
 end
 else
 if pc.ActivePageIndex = 4 then
 begin
  ImportExportName := Edit1.Text;
  // &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
  ModalResult := mrOk;
  // &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
 end;

end
else
if MasterType=3 then
begin
 if pc.ActivePageIndex = 0 then
 begin
  if ListView2.Items.Count = 0 then
   begin
    MessageBOX(Handle, PChar('Необходимо выбрать одного или нескольких спортсменов.'), PChar(MainTitle) ,
    MB_ICONWARNING);
    Exit;
   end;
  Memo5.Lines.Clear;
  ST := TStringList.Create;
  for I := 0 to ListView2.Items.Count - 1 do
   ST.Add(ListView2.Items[i].SubItems[0]);
  Memo5.Lines.AddStrings(ST);
  ST.Free;
  Label17.Caption := 'Экспорт информации в Microsoft Excel';
  ButtonNext.Caption := 'Готово';
  pc.ActivePageIndex := 4;
  ButtonBack.Enabled := true;
 end
 else
 if pc.ActivePageIndex = 4 then
  ModalResult := mrOk;
end;

end;

procedure TDiagForm.ComboBox2Change(Sender: TObject);
var
  ST : TStringList;
  i : integer;
  LI : TListItem;
begin
  ST := TStringList.Create;
  MainForm.ReadExecs2ToString(ComboBox2.ItemIndex,ST);
  ListView3.Items.Clear;
  for i := 0 to ST.Count - 1 do
  begin
    LI := ListView3.Items.Add;
    LI.Caption := IntToStr(ListView3.Items.Count);
    LI.SubItems.Add(ST[i]);
  end;

  St.Clear;
  ListView1.Items.Clear;
  ManForm.fname := fname;
  ManForm.ScanDBToStringsIndex(st,ComboBox2.Text,ComboBox3.ItemIndex-1);
  for I := 0 to st.Count - 1 do
  begin
   LI := ListView1.Items.Add;
   LI.Caption := IntToStr(ListView1.Items.Count);
   LI.SubItems.Add(st[i]);
  end;

  ST.Free;
end;

procedure TDiagForm.ComboBox3Change(Sender: TObject);
var
  ST : TStringList;
  i : integer;
  LI : TListItem;
begin
  ST := TStringList.Create;
  ListView1.Items.Clear;
  ManForm.fname := fname;
  ManForm.ScanDBToStringsIndex(st,ComboBox2.Text,ComboBox3.ItemIndex-1);
  for I := 0 to st.Count - 1 do
  begin
   LI := ListView1.Items.Add;
   LI.Caption := IntToStr(ListView1.Items.Count);
   LI.SubItems.Add(st[i]);
  end;
end;

procedure TDiagForm.FormShow(Sender: TObject);
var
 ST : TStringList;
 i : integer;
 LI : TListItem;
begin
 if MasterType=0 then
  Caption := 'Мастер диаграмм'
 else
 if MasterType=1 then
  Caption := 'Экспорт информации'
 else
 if MasterType=2 then
  Caption := 'Импорт информации'
 else
 if MasterType=3 then
  Caption := 'Экспорт в Microsoft Excel';

 if (MasterType<=1) or (MasterType=3) then
 begin
 Edit1.Text := '';
// GroupBox6.Visible := false;
 pc.ActivePageIndex := 0;
 ButtonBack.Enabled := false;
 ButtonNext.Caption := 'Далее >';
 ButtonNext.Enabled := true;


 ListView1.Items.Clear;
// ListView2.Items.Clear;

 ST := TStringList.Create;
 ST.Clear;

 MainForm.ReadExecsToString(ST);
 ComboBox2.Items.Clear;
 ComboBox2.Items.AddStrings(ST);
 ComboBox2.ItemIndex := 0;

 ST.Clear;
 MainForm.ReadExecs2ToString(0,ST);

 ListView3.Items.Clear;
// ListView4.Items.Clear;
// ListView5.Items.Clear;
// ListView6.Items.Clear;
 ListView7.Items.Clear;

 for i := 0 to ST.Count - 1 do
  begin
    LI := ListView3.Items.Add;
    LI.Caption := IntToStr(ListView3.Items.Count);
    LI.SubItems.Add(ST[i]);
  end;

 St.Clear;
 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'ofp.mtx';
 if RazrDlg.GetRazr(ST) then
 begin
  for i := 0 to ST.Count - 1 do
  begin
    LI := ListView7.Items.Add;
    LI.Caption := IntToStr(ListView7.Items.Count);
    LI.SubItems.Add(ST[i]);
  end;
 end;

 St.Clear;
 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'razr.mtx';
 if RazrDlg.GetRazr(ST) then
 begin
  ComboBox3.Items.Clear;
  ComboBox3.Items.Add('Все разряды');
  ComboBox3.Items.AddStrings(ST);
  ComboBox3.ItemIndex := 0;
 end;

 St.Clear;
 RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'control.mtx';
 RazrDlg.GetRazr(ST);
// ComboBox1.Items.Clear;
//ComboBox1.Items.AddStrings(ST);

 St.Clear;
 ManForm.fname := fname;
 ManForm.ScanDBToStringsIndex(st,ComboBox2.Text,ComboBox3.ItemIndex-1);
 for I := 0 to st.Count - 1 do
 begin
  LI := ListView1.Items.Add;
  LI.Caption := IntToStr(ListView1.Items.Count);
  LI.SubItems.Add(st[i]);
 end;

 ST.Free;
 end
 else
 if MasterType=2 then
 begin
  Edit1.Text := '';
  // GroupBox6.Visible := true;
  pc.ActivePageIndex := 3;
  ButtonBack.Enabled := false;
  ButtonNext.Caption := 'Далее >';
  ButtonNext.Enabled := true;
 end;
 Label18.Visible := false;
end;

procedure TDiagForm.Label1MouseEnter(Sender: TObject);
begin
 Screen.Cursor := crHandPoint;
end;

procedure TDiagForm.Label1MouseLeave(Sender: TObject);
begin
 Screen.Cursor := crDefault;
end;

procedure TDiagForm.ListView1Click(Sender: TObject);
begin
 if ListView1.Selected <> nil then ListView1.Selected.Checked := true;
end;

procedure TDiagForm.ListView3Click(Sender: TObject);
begin
 if ListView3.Selected <> nil then ListView3.Selected.Checked := true;
end;

procedure TDiagForm.ListView7Click(Sender: TObject);
begin
 if ListView7.Selected <> nil then ListView7.Selected.Checked := true;
end;

procedure TDiagForm.RadioButton10Click(Sender: TObject);
begin
 MezoBox.Enabled := RadioButton10.Checked;
end;

procedure TDiagForm.RadioButton2Click(Sender: TObject);
begin
 MezoB.Enabled := RadioButton7.Checked;
 MezoE.Enabled := RadioButton7.Checked;
 MicroB.Enabled := RadioButton7.Checked;
 MicroE.Enabled := RadioButton7.Checked;
end;

procedure TDiagForm.RadioButton6Click(Sender: TObject);
begin
 MezoBox.Enabled := false;
end;

procedure TDiagForm.RadioButton7Click(Sender: TObject);
begin
 MezoB.Enabled := RadioButton7.Checked;
 MezoE.Enabled := RadioButton7.Checked;
 MicroB.Enabled := RadioButton7.Checked;
 MicroE.Enabled := RadioButton7.Checked;
end;

procedure TDiagForm.ToolButton1Click(Sender: TObject);
var
 i : integer;
begin
 for i := 0 to ListView1.Items.Count - 1 do
    ListView1.Items[i].Checked := true;
end;

procedure TDiagForm.ToolButton2Click(Sender: TObject);
var
 i : integer;
begin
 for i := 0 to ListView1.Items.Count - 1 do
    ListView1.Items[i].Checked := false;
end;

procedure TDiagForm.ToolButton3Click(Sender: TObject);
var
 i : integer;
begin
 for i := 0 to ListView2.Items.Count - 1 do
    ListView2.Items[i].Checked := true;
end;

procedure TDiagForm.ToolButton4Click(Sender: TObject);
var
 i : integer;
begin
 for i := 0 to ListView2.Items.Count - 1 do
    ListView2.Items[i].Checked := false;
end;

procedure TDiagForm.ToolButton5Click(Sender: TObject);
var
 i : integer;
begin
 for i := 0 to ListView3.Items.Count - 1 do
    ListView3.Items[i].Checked := true;
end;

procedure TDiagForm.ToolButton6Click(Sender: TObject);
var
 i : integer;
begin
 for i := 0 to ListView3.Items.Count - 1 do
    ListView3.Items[i].Checked := false;
end;

end.
