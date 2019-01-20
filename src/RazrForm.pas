unit RazrForm;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls, 
  Buttons, ExtCtrls, OKCANCL1;

type
  TRazrDlg = class(TOKBottomDlg)
    HelpBtn: TButton;
    ListBox1: TListBox;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    procedure ListBox1DblClick(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    fname : string;
    stitle, title : string;
    function GetRazr(var ST : TStringList):boolean;
  end;

var
  RazrDlg: TRazrDlg;

implementation

uses Main, OKCNHLP1;

{$R *.dfm}

procedure TRazrDlg.Button1Click(Sender: TObject);
begin
  inherited;
  OKHelpBottomDlg.Caption := 'Введите новый '+stitle;
  OKHelpBottomDlg.Edit1.Text := '';
  if OKHelpBottomDlg.ShowModal = mrOk then
   ListBox1.Items.Add(OKHelpBottomDlg.Edit1.Text);
end;

procedure TRazrDlg.Button2Click(Sender: TObject);
begin
  inherited;
  if ListBox1.ItemIndex<>-1 then
  OKHelpBottomDlg.Caption := 'Изменить '+stitle;
  OKHelpBottomDlg.Edit1.Text := ListBox1.Items[ListBox1.ItemIndex];
  if OKHelpBottomDlg.ShowModal = mrOk then
   ListBox1.Items[ListBox1.ItemIndex] := OKHelpBottomDlg.Edit1.Text;
end;

procedure TRazrDlg.Button3Click(Sender: TObject);
begin
  inherited;
  if ListBox1.ItemIndex<>-1 then
   if MessageBOX(Handle, PChar('Удалить '+stitle+'?'), PChar(MainTitle) ,
    MB_YESNOCANCEL or MB_ICONQUESTION) = IDYes then
      ListBox1.DeleteSelected;
end;

function TRazrDlg.GetRazr(var ST : TStringList):boolean;
var
 F : TFileStream;
 len,j,i,cnt : integer;
 buffer : PAnsiChar;
 s : string;
begin
 if FileExists(fname) then
 begin
  try
   F := TFileStream.Create(fname,fmOpenRead);
   F.Read(len,4);
   if len > RazrVersion then
    Result := false
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
     ST.Add(s);
    end;
   end;
   F.Free;
   Result := True;
  except
   Result := False;
  end;
 end;
end;

procedure TRazrDlg.FormShow(Sender: TObject);
var
 ST : TStringList;
begin
 Button1.Enabled := MainForm.AdminMode;
 Button2.Enabled := MainForm.AdminMode;
 Button3.Enabled := MainForm.AdminMode;

 Caption := title;
 ListBox1.Items.Clear;
 ST := TStringList.Create;
 if GetRazr(ST) then
  ListBox1.Items.AddStrings(ST)
 else
    MessageBOX(Handle, Pchar('Справочник "'+Title+'" не загружен...'), MainTitle ,
    MB_ICONWARNING);
 ST.Free;
end;

procedure TRazrDlg.HelpBtnClick(Sender: TObject);
begin
  Application.HelpContext(HelpContext);
end;

procedure TRazrDlg.ListBox1DblClick(Sender: TObject);
begin
  inherited;
  if MainForm.AdminMode then
   Button2Click(Sender);
end;

procedure TRazrDlg.OKBtnClick(Sender: TObject);
var
 F : TFileStream;
 i,j,len : integer;
 buffer : PAnsiChar;
 s : string;
begin
 if MainForm.AdminMode then
  try
   F := TFileStream.Create(fname,fmCreate);
   len := RazrVersion;
   F.Write(len,4);
   for I := 0 to ListBox1.Items.Count - 1 do
   begin
    s := ListBox1.Items[i];
    len := Length(s);
    F.Write(len,4);
    buffer := AllocMem(len+1);
    buffer := StrPCopy(buffer, s);
    F.Write(buffer^,len);
    FreeMem(buffer);
   end;
   F.Free;
  except
   MessageBOX(Handle, Pchar('Îøèáêà ïðè çàïèñè справочника "'+Title+'"...'),
   MainTitle ,
   MB_ICONERROR);
  end;
end;

end.

