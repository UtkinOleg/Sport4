unit Spr;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ExtCtrls, ImgList, ToolWin;


type

  TExe = class
    ST : TStringList;
    constructor Create;
    destructor Destroy; override;
  end;

  TFormSpr = class(TForm)
    Splitter1: TSplitter;
    ImageList1: TImageList;
    Panel1: TPanel;
    ControlBar2: TControlBar;
    ToolBar1: TToolBar;
    ToolButtonAdd: TToolButton;
    ToolButton14: TToolButton;
    ToolButton15: TToolButton;
    ToolButton16: TToolButton;
    ListView1: TListView;
    Splitter2: TSplitter;
    Panel2: TPanel;
    ListView2: TListView;
    ControlBar1: TControlBar;
    ToolBar2: TToolBar;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButtonDel: TToolButton;
    procedure FormActivate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ListView2DblClick(Sender: TObject);
    procedure ToolButtonAddClick(Sender: TObject);
    procedure ToolButton16Click(Sender: TObject);
    procedure ToolButton8Click(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ListView1DblClick(Sender: TObject);
    procedure ListView2Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure ListView1Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButtonDelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    ExeList : TList;
    procedure SetModified;
    procedure SetEnable;
    procedure SetEnable2;
    procedure SetDisable2;
  public
    { Public declarations }
    Modified : boolean;
    ExecReadOnly : boolean;
    procedure SaveExecData;
  end;

var
  FormSpr: TFormSpr;

implementation

uses OKCNHLP1, Main;

{$R *.dfm}

constructor TExe.Create;
begin
  ST := TStringList.Create;
  inherited Create;
end;

destructor TExe.Destroy;
begin
  ST.Free;
  inherited Destroy;
end;

procedure TFormSpr.FormActivate(Sender: TObject);
var
 i : integer;
begin
 for I := 0 to MainForm.DockTabSet1.Tabs.Count - 1 do
  if Caption = Trim(MainForm.DockTabSet1.Tabs[i]) then
  begin
    MainForm.DockTabSet1.TabIndex := i;
    break;
  end;

  MainForm.N9.Enabled := false;
  MainForm.ToolButton4.Enabled := false;
  MainForm.N27.Enabled := false;
  MainForm.N18.Enabled := false;
  MainForm.ToolButton14.Enabled := false;
  MainForm.C1.Enabled := false;
  MainForm.ToolButton8.Enabled := false;
  MainForm.N14.Enabled := false;

  MainForm.ToolButton11.Enabled := false;
  MainForm.N16.Enabled := false;
  MainForm.ToolButton12.Enabled := false;
  MainForm.N17.Enabled := false;
  MainForm.ToolButton14.Enabled := false;
  MainForm.N24.Enabled := false;
  MainForm.SB.Panels[0].Text := '';
  MainForm.N27.Enabled := false;
end;

procedure TFormSpr.FormClose(Sender: TObject; var Action: TCloseAction);
var
  i : integer;
begin
  if MainForm.MDIChildCount = 1 then
   begin
    MainForm.ToolButton2.Enabled := false;
    MainForm.FileSaveItem.Enabled := false;
    MainForm.DockTabSet1.Visible := false;

    MainForm.ToolButton14.Enabled := false;
    MainForm.N24.Enabled := false;
   end;

  for i := ExeList.Count - 1 downto 0 do
   TExe(ExeList.Items[i]).Free;
  ExeList.Free;

  for I := 0 to MainForm.DockTabSet1.Tabs.Count - 1 do
   if Trim(MainForm.DockTabSet1.Tabs[i]) = Caption then
    begin
     MainForm.DockTabSet1.Tabs.Delete(i);
     break;
    end;

  Action := caFree;
end;

procedure TFormSpr.SaveExecData;
var
 F : TFileStream;
 i,j,len : integer;
 buffer : PAnsiChar;
 Exe : TExe;
 s : string;
begin
 if not ExecReadOnly then
  try
   F := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'exec.mtx',fmCreate);
   len := ExecVersion;
   F.Write(len,4);
   for I := 0 to ListView1.Items.Count - 1 do
   begin
    s := ListView1.Items[i].SubItems[0];
    len := Length(s);
    F.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    F.Write(buffer^,len);
    FreeMem(buffer);
    Exe := ListView1.Items[i].Data;

    len := Exe.ST.Count;
    F.Write(len,4);
    for j := 0 to Exe.ST.Count - 1 do
    begin
     s := Exe.ST.Strings[j];
     len := Length(s);
     F.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     F.Write(buffer^,len);
     FreeMem(buffer);
    end;
   end;
   F.Free;
   MainForm.ToolButton2.Enabled := false;
   MainForm.FileSaveItem.Enabled := false;
   Modified := false;
  except
   MessageBOX(Handle, 'Ошибка при записи справочника "'+ExecTitle+'"...', MainTitle ,
   MB_ICONERROR);
  end;
end;

procedure TFormSpr.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
 J : Integer;
begin
  if Modified then
  begin
    case MessageBOX(Handle, 'Сохранить изменения в справочнике "' +
      ExecTitle + '" ?', MainTitle ,
      MB_YESNOCANCEL or MB_ICONQUESTION) of
      IDYes:
        begin
          SaveExecData;
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

procedure TFormSpr.FormCreate(Sender: TObject);
var
 F : TFileStream;
 len,j,i,cnt : integer;
 buffer : PAnsiChar;
 s : string;
 LI : TListItem;
 Exe : TExe;
begin
 ExeList := TList.Create;
 SetModified;
 Modified := false;
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

     LI := ListView1.Items.Add;
     LI.Caption := IntToStr(ListView1.Items.Count);
     LI.SubItems.Add(s);
     ExeList.Add(TExe.Create);
     LI.Data := ExeList.Last;
     Exe := LI.Data;

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
        Exe.ST.Add(s);
       end;
     end;
    end;

    end;
   F.Free;
  except
   MessageBOX(Handle, 'Îøèáêà ïðè чтении справочника "'+ExecTitle+'"...', MainTitle ,
   MB_ICONERROR);
  end;
 end;
end;

procedure TFormSpr.FormShow(Sender: TObject);
begin
  MainForm.ToolButton8.Enabled := false;
  MainForm.N14.Enabled := false;

  ToolButtonAdd.Enabled := MainForm.AdminMode;
  ToolButton14.Enabled := MainForm.AdminMode;
  ToolButton15.Enabled := MainForm.AdminMode;
  ToolButton16.Enabled := MainForm.AdminMode;
  ExecReadOnly := not MainForm.AdminMode;
end;

procedure TFormSpr.ListView1Click(Sender: TObject);
var
 Exe : TExe;
 i : integer;
 LI : TListItem;
begin
 SetEnable;
 SetDisable2;
 ListView2.Items.Clear;
 if ListView1.Selected <> nil then
 begin
  Exe := ListView1.Selected.Data;
  for I := 0 to Exe.ST.Count - 1 do
   begin
    LI := ListView2.Items.Add;
    LI.Caption := IntToStr(ListView2.Items.Count);
    LI.SubItems.Add(Exe.ST.Strings[i]);
   end;
 end;
end;

procedure TFormSpr.ListView1DblClick(Sender: TObject);
begin
 if MainForm.AdminMode then
  ToolButton1Click(Sender);
end;

procedure TFormSpr.ListView2Click(Sender: TObject);
begin
 SetEnable2;
end;

procedure TFormSpr.ListView2DblClick(Sender: TObject);
begin
 if not ExecReadOnly then
  ToolButton7Click(Sender);
end;

procedure TFormSpr.SetModified;
begin
  if not ExecReadOnly then
   begin
    Modified := true;
    MainForm.ToolButton2.Enabled := true;
    MainForm.FileSaveItem.Enabled := true;
   end;
end;

procedure TFormSpr.SetEnable;
begin
  if not ExecReadOnly then
   begin
    ToolButtonAdd.Enabled := true;
    ToolButton14.Enabled := true;
    ToolButton15.Enabled := true;
    ToolButton16.Enabled := true;
   end;
end;

procedure TFormSpr.SetEnable2;
begin
  if not ExecReadOnly then
   begin
    ToolButton6.Enabled := true;
    ToolButton7.Enabled := true;
    ToolButton8.Enabled := true;
    ToolButtonDel.Enabled := true;
   end;
end;

procedure TFormSpr.SetDisable2;
begin
  if not ExecReadOnly then
   ToolButton6.Enabled := true;
  ToolButton7.Enabled := false;
  ToolButton8.Enabled := false;
  ToolButtonDel.Enabled := false;
end;

procedure TFormSpr.ToolButton16Click(Sender: TObject);
begin
   if MessageBOX(Handle, PChar('Удалить все виды спорта (упражнения также будут удалены) ?'), PChar(MainTitle) ,
    MB_YESNOCANCEL or MB_ICONQUESTION) = IDYes then
     begin
      SetModified;
      ListView1.Items.Clear;
      ListView2.Items.Clear;
     end;
end;

procedure TFormSpr.ToolButton1Click(Sender: TObject);
begin
 if ListView1.Selected <> nil then
 begin
  OKHelpBottomDlg.Caption := 'Редактировать вид спорта';
  OKHelpBottomDlg.Edit1.Text := ListView1.Selected.SubItems[0];
  if OKHelpBottomDlg.ShowModal = mrOk then
   begin
    SetModified;
    ListView1.Selected.SubItems[0] := OKHelpBottomDlg.Edit1.Text;
   end;
 end;
end;

procedure TFormSpr.ToolButton2Click(Sender: TObject);
begin
 if ListView1.Selected <> nil then
 begin
   if MessageBOX(Handle, PChar('Удалить вид спорта "'+ListView1.Selected.SubItems[0]+'" (упражнения для выбранного вида также будут удалены) ?'), PChar(MainTitle) ,
    MB_YESNOCANCEL or MB_ICONQUESTION) = IDYes then
     begin
      SetModified;
      ListView1.Selected.Delete;
      ListView2.Items.Clear;
     end;
 end;
end;

procedure TFormSpr.ToolButton5Click(Sender: TObject);
var
 Exe : TExe;
 LI : TListItem;
begin
 if ListView1.Selected <> nil then
 begin
  OKHelpBottomDlg.Caption := 'Введите новое упражнение для вида спорта "' + ListView1.Selected.SubItems[0]+'"';
  OKHelpBottomDlg.Edit1.Text := '';
  if OKHelpBottomDlg.ShowModal = mrOk then
  begin
   SetModified;
   Exe := ListView1.Selected.Data;
   LI := ListView2.Items.Add;
   LI.Caption := IntToStr(ListView2.Items.Count);
   LI.SubItems.Add(OKHelpBottomDlg.Edit1.Text);
   Exe.ST.Add(OKHelpBottomDlg.Edit1.Text);
  end;
 end;
end;

procedure TFormSpr.ToolButton7Click(Sender: TObject);
var
 Exe : TExe;
begin
 if ListView2.Selected <> nil then
 begin
  OKHelpBottomDlg.Caption := 'Редактировать упражнение';
  OKHelpBottomDlg.Edit1.Text := ListView2.Selected.SubItems[0];
  if OKHelpBottomDlg.ShowModal = mrOk then
   begin
    SetModified;
    Exe := ListView1.Selected.Data;
    ListView2.Selected.SubItems[0] := OKHelpBottomDlg.Edit1.Text;
    Exe.ST[ListView2.Selected.Index] := OKHelpBottomDlg.Edit1.Text;
   end;
 end;
end;

procedure TFormSpr.ToolButton8Click(Sender: TObject);
var
 Exe : TExe;
begin
 if ListView1.Selected <> nil then
 begin
   if MessageBOX(Handle, PChar('Удалить выбранное упражнение "'+ ListView2.Selected.SubItems[0]+'"?'), PChar(MainTitle) ,
    MB_YESNOCANCEL or MB_ICONQUESTION) = IDYes then
     begin
      SetModified;
      Exe := ListView1.Selected.Data;
      Exe.ST.Delete(ListView2.Selected.Index);
      ListView2.Selected.Delete;
     end;
 end;
end;

procedure TFormSpr.ToolButtonAddClick(Sender: TObject);
var
 LI : TListItem;
begin
 OKHelpBottomDlg.Caption := 'Введите новый вид спорта';
 OKHelpBottomDlg.Edit1.Text := '';
 if OKHelpBottomDlg.ShowModal = mrOk then
 begin
  SetModified;
  LI := ListView1.Items.Add;
  LI.Caption := IntToStr(ListView1.Items.Count);
  LI.SubItems.Add(OKHelpBottomDlg.Edit1.Text);
  ExeList.Add(TExe.Create);
  LI.Data := ExeList.Last;
 end;
end;

procedure TFormSpr.ToolButtonDelClick(Sender: TObject);
var
 Exe : TExe;
begin
   if MessageBOX(Handle, PChar('Удалить все упражнения ?'), PChar(MainTitle) ,
    MB_YESNOCANCEL or MB_ICONQUESTION) = IDYes then
     begin
      SetModified;
      Exe := ListView1.Selected.Data;
      Exe.ST.Clear;
      ListView2.Items.Clear;
     end;
end;

end.
