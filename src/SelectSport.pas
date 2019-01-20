unit SelectSport;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls;

type
  TSelectSportDlg = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Bevel1: TBevel;
    ListBox1: TListBox;
    procedure OKBtnClick(Sender: TObject);
    procedure ListBox1DblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SelectSportDlg: TSelectSportDlg;

implementation

uses Main;

{$R *.dfm}

procedure TSelectSportDlg.FormShow(Sender: TObject);
var
 F : TFileStream;
 len,j,i,cnt : integer;
 buffer : PAnsiChar;
 s : string;
begin
 ListBox1.Items.Clear;
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
     ListBox1.Items.Add(s);

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
   MessageBOX(Handle, 'Îøèáêà ïðè чтении справочника "'+ExecTitle+'"...', MainTitle ,
   MB_ICONERROR);
  end;
 end;
end;

procedure TSelectSportDlg.ListBox1DblClick(Sender: TObject);
begin
 OKBtnClick(Sender);
end;

procedure TSelectSportDlg.OKBtnClick(Sender: TObject);
begin
 if ListBox1.ItemIndex <> -1 then
   ModalResult := mrOk;
end;

end.
