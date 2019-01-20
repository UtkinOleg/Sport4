unit Man;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ExtCtrls, ImgList, StdCtrls;

type

  ManRec = record
    Size, Pos, PosLast, ID : Integer;
  end;
  pManRec = ^ManRec;

  TManForm = class(TForm)
    ImageList1: TImageList;
    ControlBar1: TControlBar;
    ToolBar2: TToolBar;
    ToolButton9: TToolButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton4: TToolButton;
    ListView1: TListView;
    CancelBtn: TButton;
    HelpBtn: TButton;
    OKBtn: TButton;
    procedure FormShow(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ListView1DblClick(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton9Click(Sender: TObject);
  private
    { Private declarations }
  public
    fname : string;
    { Public declarations }
    procedure ScanDB;
    procedure ScanDBToStrings(var st : TStringList);
    procedure AddMem(M : TMemoryStream);
    procedure ScanDBToStringsIndex(var st : TStringList; sport : string; razr : integer);
  end;

var
  ManForm: TManForm;
  tManRec : pManRec;
  recList : TList;

implementation

uses MenFormDlg, Main, RazrForm, ChildWin;

{$R *.dfm}

procedure TManForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
// recList.Free;
end;

procedure TManForm.ScanDB;
var
 F : TFileStream;
 mensize, cnt, id, i,j,len : integer;
 buffer : PAnsiChar;
 s : string;
 LI : TListItem;
begin

 for I := 0 to ListView1.Items.Count - 1 do
   begin
    tManRec := ListView1.Items[i].Data;
    Dispose(tManRec);
   end;
 ListView1.Items.Clear;
 id := 0;
// recList := TList.Create;
 if FileExists(fname) then
    try
     F := TFileStream.Create(fname,fmOpenRead);
     F.Read(len,4);
     if len<=DBVersion then
     begin
      F.Read(len,4);
      mensize := len;
      while F.Position < mensize do
      begin
      New(tManRec);
//      recList.Add(tManRec);
      inc(id);
      tManRec^.ID := id;
      tManRec^.Pos := F.Position;
      // Размер записи
      F.Read(len,4);
      tManRec^.Size := len;
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
      LI := ListView1.Items.Add;
      LI.Caption := IntToStr(ListView1.Items.Count);
      LI.ImageIndex := 3;
      LI.SubItems.Add(s);
      LI.Data := tManRec;

      // Разряд
      F.Read(len,4);
      F.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      F.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := StrPas(buffer);
      FreeMem(buffer);
      LI.SubItems.Add(s);

      F.Read(cnt,4);
      s := '';
      for I := 0 to cnt - 1 do
      begin
       F.Read(len,4);
       buffer := AllocMem(len+1);
       buffer[len] := Chr(0);
       F.Read(buffer^,len);
       buffer[len]:=Chr(0);
       s := s + ' ' + StrPas(buffer);
       FreeMem(buffer);
      end;
      LI.SubItems.Add(s);
      tManRec^.PosLast := F.Position;
      end;
     end
     else
      MessageBOX(Handle, 'Версия справочника "'+ManTitle+'" устарела...', MainTitle ,
      MB_ICONWarning);

     F.Free;
     ToolButton1.Enabled := true;
     ToolButton2.Enabled := true;
     ToolButton4.Enabled := true;

    except
     MessageBOX(Handle, 'Ошибка при чтении справочника "'+ManTitle+'"...', MainTitle ,
     MB_ICONERROR);
    end;
end;

procedure TManForm.ScanDBToStrings(var st : TStringList);
var
 F : TFileStream;
 mensize, cnt, id, i,j,len : integer;
 buffer : PAnsiChar;
 s : string;
 LI : TListItem;
begin
 if FileExists(fname) then
    try
     F := TFileStream.Create(fname,fmOpenRead);
     F.Read(len,4);
     if len<=DBVersion then
     begin
      F.Read(len,4);
      mensize := len;
      while F.Position < mensize do
      begin
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
      st.Add(s);

      // Разряд
      F.Read(len,4);
      F.Read(len,4);
      F.Seek(len,soFromCurrent);

      F.Read(cnt,4);
      for I := 0 to cnt - 1 do
      begin
       F.Read(len,4);
       F.Seek(len,soFromCurrent);
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

procedure TManForm.ScanDBToStringsIndex(var st : TStringList; sport : string; razr : integer);
var
 F : TFileStream;
 mensize, cnt, id, i,j,len : integer;
 buffer : PAnsiChar;
 s, s2 : string;
 LI : TListItem;
 razrfound, found : boolean;
begin
 if FileExists(fname) then
    try
     F := TFileStream.Create(fname,fmOpenRead);
     F.Read(len,4);
     if len<=DBVersion then
     begin
      F.Read(len,4);
      mensize := len;
      while F.Position < mensize do
      begin
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

      found := false;
      // Разряд
      F.Read(len,4);
      razrfound := (len = razr) or (razr= -1);

      F.Read(len,4);
      F.Seek(len,soFromCurrent);

      F.Read(cnt,4);
      for I := 0 to cnt - 1 do
       begin
        F.Read(len,4);
        buffer := AllocMem(len+1);
        buffer[len] := Chr(0);
        F.Read(buffer^,len);
        buffer[len]:=Chr(0);
        s2 := StrPas(buffer);
        if s2 = sport then
          found := true;
        FreeMem(buffer);
       end;

      if found and razrfound then
       st.Add(s);

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

procedure TManForm.FormShow(Sender: TObject);
begin
 ScanDB;
end;

procedure TManForm.ListView1DblClick(Sender: TObject);
begin
 ToolButton1Click(Sender);
end;

procedure TManForm.ToolButton1Click(Sender: TObject);
var
  F : TFileStream;
  M : TMemoryStream;
  oldsize, mansize, last, begsize, endsize, cnt, id, i,j,len : integer;
  buffer : PAnsiChar;
  bbuffer, pbuffer : PByte;
  oldfam, oldname, oldsurname, s : string;
  LI : TListItem;
  ST : TStringList;
begin
 if ListView1.Selected <> nil then
 begin
  FioDlg.LabeledEdit1.Clear;
  FioDlg.LabeledEdit2.Clear;
  FioDlg.LabeledEdit3.Clear;
  FioDlg.ListBox1.Items.Clear;
  FioDlg.ComboBox1.Items.Clear;

  ST := TStringList.Create;
  RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'razr.mtx';
  if RazrDlg.GetRazr(ST) then
   FioDlg.ComboBox1.Items.AddStrings(ST)
  else
    MessageBOX(Handle, 'Справочник "'+RazrTitle+'" не загружен...', MainTitle ,
    MB_ICONWARNING);
  ST.Free;
  tManRec := ListView1.Selected.Data;
  try

      F := TFileStream.Create(fname,fmOpenRead);
      F.Read(len,4);
      F.Read(mansize,4);

      i:=1;
      while i<>tManRec^.ID do
      begin
       F.Read(len,4);
       F.Seek(len-4,soFromCurrent);
       inc(i);
      end;

      F.Read(len,4);
      oldsize := len;

      // Фамилия
      F.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      F.Read(buffer^,len);
      buffer[len]:=Chr(0);
      FioDlg.LabeledEdit1.Text := StrPas(buffer);
      oldfam := StrPas(buffer);
      FreeMem(buffer);
      // Имя
      F.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      F.Read(buffer^,len);
      buffer[len]:=Chr(0);
      FioDlg.LabeledEdit2.Text := StrPas(buffer);
      oldname := StrPas(buffer);
      FreeMem(buffer);
      // Отчество
      F.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      F.Read(buffer^,len);
      buffer[len]:=Chr(0);
      FioDlg.LabeledEdit3.Text := StrPas(buffer);
      oldsurname := StrPas(buffer);
      FreeMem(buffer);
      // Разряд
      F.Read(len,4);
      FioDlg.ComboBox1.ItemIndex := len;
      F.Read(len,4);
      buffer := AllocMem(len+1);
      buffer[len] := Chr(0);
      F.Read(buffer^,len);
      buffer[len]:=Chr(0);
      s := StrPas(buffer);
      FreeMem(buffer);

      F.Read(cnt,4);
      for I := 0 to cnt - 1 do
      begin
       F.Read(len,4);
       buffer := AllocMem(len+1);
       buffer[len] := Chr(0);
       F.Read(buffer^,len);
       buffer[len]:=Chr(0);
       s := StrPas(buffer);
       FreeMem(buffer);
       FioDlg.ListBox1.Items.Add(s);
      end;

   F.Free;
  except
   MessageBOX(Handle, 'Ошибка при чтении справочника "'+ManTitle+'"...', MainTitle ,
   MB_ICONERROR);
  end;

 if FioDlg.ShowModal = mrok then
  begin
   ListView1.Selected.SubItems.Clear;
   ListView1.Selected.SubItems.Add(FIODlg.LabeledEdit1.Text + ' ' + FIODlg.LabeledEdit2.Text + ' ' + FIODlg.LabeledEdit3.Text);
   ListView1.Selected.SubItems.Add(FIODlg.ComboBox1.Text);
   s := '';
   for I := 0 to FIODlg.ListBox1.Items.Count - 1 do
    s := s + ' ' + FIODlg.ListBox1.Items[i];
   ListView1.Selected.SubItems.Add(s);
   try
    F := TFileStream.Create(fname,fmOpenReadWrite);
    F.Seek(0,soFromBeginning);

    F.Read(len,4);
    F.Read(mansize,4);

    i:=1;
    while i<>tManRec^.ID do
    begin
     F.Read(len,4);
     F.Seek(len-4,soFromCurrent);
     inc(i);
    end;

    begsize := F.Position;

    bbuffer := AllocMem(begsize);
    F.Seek(0,soFromBeginning);
    F.Read(bbuffer^,begsize);

    F.Read(len,4);
    F.Seek(len-4,soFromCurrent);

    endsize := F.Size - begsize - len;
    pbuffer := AllocMem(endsize);
    F.Read(pbuffer^,endsize);

    F.Free;

    F := TFileStream.Create(fname,fmCreate);
    F.Write(bbuffer^,begsize);
    FreeMem(bbuffer);

    M := TMemoryStream.Create;
    // Фамилия
    s := FIODlg.LabeledEdit1.Text;
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);
    // Имя
    s := FIODlg.LabeledEdit2.Text;
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);
    // Отчество
    s := FIODlg.LabeledEdit3.Text;
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);
    // Разряд
    s := FIODlg.ComboBox1.Text;
    len := FIODlg.ComboBox1.ItemIndex;
    M.Write(len,4);
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);
    // Виды спорта
    len := FIODlg.ListBox1.Items.Count;
    M.Write(len,4);
    for I := 0 to FIODlg.ListBox1.Items.Count - 1 do
    begin
     s := FIODlg.ListBox1.Items[i];
     len := Length(s);
     M.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     M.Write(buffer^,len);
     FreeMem(buffer);
    end;

    M.Position := 0;
    len := M.Size+4;
    mansize := mansize - oldsize + len;
    F.Write(len,4);
    bbuffer := AllocMem(M.Size);
    M.Read(bbuffer^,M.Size);
    F.Write(bbuffer^,M.Size);
    M.Free;
    FreeMem(bbuffer);

    F.Write(pbuffer^,endsize);
    FreeMem(pbuffer);

    F.Seek(0,soFromBeginning);
    F.Read(len,4);
    F.Write(mansize,4);

    F.Free;

    // Изменим БД
    for I := 0 to TMDIChild(MainForm.ActiveMDIChild).BigDataLeaf.Count - 1 do
      if TDataLeaf(TMDIChild(MainForm.ActiveMDIChild).BigDataLeaf.Items[i]).menindex<>-1 then
       begin
        if TDataLeaf(TMDIChild(MainForm.ActiveMDIChild).BigDataLeaf.Items[i]).men = oldfam + ' ' + oldname + ' ' + oldsurname then
         begin
          TDataLeaf(TMDIChild(MainForm.ActiveMDIChild).BigDataLeaf.Items[i]).men :=
          FIODlg.LabeledEdit1.Text + ' ' + FIODlg.LabeledEdit2.Text + ' ' + FIODlg.LabeledEdit3.Text;
         end;
       end;
    TMDIChild(MainForm.ActiveMDIChild).SaveDataBase;
   except
    MessageBOX(Handle, 'Ошибка при изменении справочника "'+ManTitle+'"...', MainTitle ,
    MB_ICONERROR);
   end;
  end;
 end;
end;

procedure TManForm.ToolButton2Click(Sender: TObject);
var
  F : TFileStream;
  M : TMemoryStream;
  mansize, last, begsize, endsize, cnt, id, i,j,len : integer;
  buffer : PAnsiChar;
  bbuffer, pbuffer : PByte;
  oldfam, oldname, oldsurname, s : string;
  LI : TListItem;
  ST : TStringList;
begin
 if ListView1.Selected <> nil then
  if MessageBOX(Handle, PChar('Удалить спортсмена '+ListView1.Selected.SubItems[0]+'? Внимание!' +
   ' В основной базе данных также будет удалена вся информация по '+
   'выбранному спортсмену.'), PChar(MainTitle) ,
     MB_YESNO or MB_ICONQUESTION) = IDYes then
   begin
   tManRec := ListView1.Selected.Data;
   try
    F := TFileStream.Create(fname,fmOpenReadWrite);
    F.Seek(0,soFromBeginning);

    F.Read(len,4);
    F.Read(mansize,4);

    i:=1;
    while i<>tManRec^.ID do
    begin
     F.Read(len,4);
     F.Seek(len-4,soFromCurrent);
     inc(i);
    end;

    begsize := F.Position;

    bbuffer := AllocMem(begsize);
    F.Seek(0,soFromBeginning);
    F.Read(bbuffer^,begsize);

    F.Read(len,4);
    mansize := mansize - len;
    F.Seek(len-4,soFromCurrent);

    endsize := F.Size - begsize - len;
    pbuffer := AllocMem(endsize);
    F.Read(pbuffer^,endsize);
    F.Free;

    F := TFileStream.Create(fname,fmCreate);
    F.Write(bbuffer^,begsize);
    FreeMem(bbuffer);
    F.Write(pbuffer^,endsize);
    FreeMem(pbuffer);
    F.Seek(0,soFromBeginning);
    F.Read(len,4);
    F.Write(mansize,4);

    F.Free;

    // Изменим БД
{    ST := TStringList.Create;
    for I := 0 to TMDIChild(MainForm.ActiveMDIChild).BigDataLeaf.Count - 1 do
      if TDataLeaf(TMDIChild(MainForm.ActiveMDIChild).BigDataLeaf.Items[i]).menindex<>-1 then
       begin
        if TDataLeaf(TMDIChild(MainForm.ActiveMDIChild).BigDataLeaf.Items[i]).men = ListView1.Selected.SubItems[0] then
          ST.Add(IntToStr(i));
       end;
    for i:=0 to ST.Count-1 do
     TMDIChild(MainForm.ActiveMDIChild).BigDataLeaf.Delete(StrToInt(ST[i]));
    ST.Free;
    TMDIChild(MainForm.ActiveMDIChild).SaveDataBase(fname);}

   except
    MessageBOX(Handle, 'Ошибка при изменении справочника "'+ManTitle+'"...', MainTitle ,
    MB_ICONERROR);
   end;
   ScanDB;
   end;
end;

procedure TManForm.ToolButton9Click(Sender: TObject);
var
 F : TFileStream;
 mansize, id, i,j,len : integer;
 buffer : PAnsiChar;
 pbuffer : PByte;
 s : string;
 LI : TListItem;
 ST : TStringList;
 M,DBM : TMemoryStream;
begin
  FioDlg.LabeledEdit1.Clear;
  FioDlg.LabeledEdit2.Clear;
  FioDlg.LabeledEdit3.Clear;
  FioDlg.ListBox1.Items.Clear;
  FioDlg.ComboBox1.Items.Clear;

  ST := TStringList.Create;
  RazrDlg.fname := ExtractFilePath(paramstr(0)) + 'razr.mtx';
  if RazrDlg.GetRazr(ST) then
   FioDlg.ComboBox1.Items.AddStrings(ST)
  else
    MessageBOX(Handle, 'Справочник "'+RazrTitle+'" не загружен...', MainTitle ,
    MB_ICONWARNING);
  ST.Free;


 if FioDlg.ShowModal = mrok then
  begin
  LI := ListView1.Items.Add;
  LI.Caption := IntToStr(ListView1.Items.Count);
  LI.ImageIndex := 3;
  LI.SubItems.Add(FIODlg.LabeledEdit1.Text + ' ' + FIODlg.LabeledEdit2.Text + ' ' + FIODlg.LabeledEdit3.Text);
  LI.SubItems.Add(FIODlg.ComboBox1.Text);
  s := '';
  for I := 0 to FIODlg.ListBox1.Items.Count - 1 do
   s := s + ' ' + FIODlg.ListBox1.Items[i];
  LI.SubItems.Add(s);
  ToolButton1.Enabled := true;
  ToolButton2.Enabled := true;
  ToolButton4.Enabled := true;
  try
    DBM := TMemoryStream.Create;
    if FileExists(fname) then
    begin
     F := TFileStream.Create(fname,fmOpenReadWrite);
     F.Read(len,4);
     F.Read(mansize,4);
     //
     F.Seek(mansize,soFromCurrent);
     len := F.Size - mansize - 8;
     if len>0 then
       DBM.CopyFrom(F,len);
     len := DBM.Size;
     F.Seek(0,soFromBeginning);
     F.Read(len,4);
    end
    else
    begin
     F := TFileStream.Create(fname,fmCreate);
     len := DBVersion;
     F.Write(len,4);
     mansize := 0;
    end;

    New(tManRec);
//    recList.Add(tManRec);
    tManRec^.Pos := F.Position;
    tManRec^.Id := ListView1.Items.Count;
    LI.Data := tManRec;

    M := TMemoryStream.Create;
    // Фамилия
    s := FIODlg.LabeledEdit1.Text;
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);
    // Имя
    s := FIODlg.LabeledEdit2.Text;
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);
    // Отчество
    s := FIODlg.LabeledEdit3.Text;
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);
    // Разряд
    s := FIODlg.ComboBox1.Text;
    len := FIODlg.ComboBox1.ItemIndex;
    M.Write(len,4);
    len := Length(s);
    M.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    M.Write(buffer^,len);
    FreeMem(buffer);
    // Виды спорта
    len := FIODlg.ListBox1.Items.Count;
    M.Write(len,4);
    for I := 0 to FIODlg.ListBox1.Items.Count - 1 do
    begin
     s := FIODlg.ListBox1.Items[i];
     len := Length(s);
     M.Write(len,4);
     buffer := AllocMem(len+1);
     buffer := StrPCopy(buffer, s);
     M.Write(buffer^,len);
     FreeMem(buffer);
    end;

    len := mansize + M.Size + 4;
    F.Write(len,4);

    F.Seek(mansize,soFromCurrent);
    M.Position := 0;
    len := M.Size + 4;
    F.Write(len,4);
    F.CopyFrom(M,M.Size);
    DBM.Position := 0;
    F.CopyFrom(DBM,DBM.Size);

    M.Free;
    DBM.Free;
    F.Free;
  except
   MessageBOX(Handle, 'Ошибка записи справочника "'+ManTitle+'"...', MainTitle ,
   MB_ICONERROR);
  end;

  end;
end;


procedure TManForm.AddMem(M : TMemoryStream);
var
 F : TFileStream;
 mansize, id, i,j,len : integer;
 buffer : PAnsiChar;
 pbuffer : PByte;
 s : string;
 LI : TListItem;
 ST : TStringList;
 DBM : TMemoryStream;
begin

  try
    DBM := TMemoryStream.Create;
    if not FileExists(fname) then
    begin
     F := TFileStream.Create(fname,fmCreate);
     len := DBVersion;
     F.Write(len,4);
     mansize := 0;
     F.Write(mansize,4);
     F.Position := 0;
    end
    else
     F := TFileStream.Create(fname,fmOpenReadWrite);

    F.Read(len,4);
    F.Read(mansize,4);
     //
    F.Seek(mansize,soFromCurrent);
    len := F.Size - mansize - 8;
    if len>0 then
       DBM.CopyFrom(F,len);
    len := DBM.Size;
    F.Seek(0,soFromBeginning);
    F.Read(len,4);

    len := mansize + M.Size + 4;
    F.Write(len,4);

    F.Seek(mansize,soFromCurrent);
    M.Position := 0;
    len := M.Size + 4;
    F.Write(len,4);
    F.CopyFrom(M,M.Size);
    DBM.Position := 0;
    F.CopyFrom(DBM,DBM.Size);

    DBM.Free;
    F.Free;

  except
   MessageBOX(Handle, 'Ошибка записи справочника "'+ManTitle+'"...', MainTitle ,
   MB_ICONERROR);
  end;

end;

end.
