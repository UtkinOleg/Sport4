unit OptionUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TOptionForm = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    HelpBtn: TButton;
    GroupBox1: TGroupBox;
    Edit1: TEdit;
    Button1: TButton;
    SaveDialog1: TSaveDialog;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OptionForm: TOptionForm;

implementation

{$R *.dfm}

uses Main;

procedure TOptionForm.Button1Click(Sender: TObject);
var
 F, F2 : TFileStream;
 len : integer;
 s : string;
 buffer : PChar;
begin
 SaveDialog1.Filter := 'Обновление программы Спорт 4 (*.data4)|*.data4';
 SaveDialog1.FileName := 'sportupdate.data4';
 SaveDialog1.Title := 'Создание обновления';
 if SaveDialog1.Execute then
 try
   F := TFileStream.Create(SaveDialog1.FileName,fmCreate);
   s := 'SPORT4UPDATE';
   len := Length(s);
   F.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   F.Write(buffer^,len);
   FreeMem(buffer);

   s := 'exec.mtx';
   len := Length(s);
   F.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   F.Write(buffer^,len);
   FreeMem(buffer);

   F2 := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'exec.mtx',fmOpenRead);
   len := F2.Size;
   F.Write(len,4);
   F.CopyFrom(F2,F2.Size);
   F2.Free;

   s := 'razr.mtx';
   len := Length(s);
   F.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   F.Write(buffer^,len);
   FreeMem(buffer);

   F2 := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'razr.mtx',fmOpenRead);
   len := F2.Size;
   F.Write(len,4);
   F.CopyFrom(F2,F2.Size);
   F2.Free;

   s := 'relax.mtx';
   len := Length(s);
   F.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   F.Write(buffer^,len);
   FreeMem(buffer);

   F2 := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'relax.mtx',fmOpenRead);
   len := F2.Size;
   F.Write(len,4);
   F.CopyFrom(F2,F2.Size);
   F2.Free;

   s := 'ofp.mtx';
   len := Length(s);
   F.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   F.Write(buffer^,len);
   FreeMem(buffer);

   F2 := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'ofp.mtx',fmOpenRead);
   len := F2.Size;
   F.Write(len,4);
   F.CopyFrom(F2,F2.Size);
   F2.Free;

   s := 'control.mtx';
   len := Length(s);
   F.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   F.Write(buffer^,len);
   FreeMem(buffer);

   F2 := TFileStream.Create(ExtractFilePath(paramstr(0)) + 'control.mtx',fmOpenRead);
   len := F2.Size;
   F.Write(len,4);
   F.CopyFrom(F2,F2.Size);
   F2.Free;

   s := 'URL';
   len := Length(s);
   F.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   F.Write(buffer^,len);
   FreeMem(buffer);

   s := Edit1.Text;
   len := Length(s);
   F.Write(len,4);
   buffer := AllocMem(len+1);
   buffer := StrPCopy(buffer, s);
   F.Write(buffer^,len);
   FreeMem(buffer);

 finally
   MessageBOX(Handle, 'Обновление создано...', MainTitle ,
   MB_ICONINFORMATION);
   F.Free;
 end;
end;

procedure TOptionForm.FormShow(Sender: TObject);
begin
 Edit1.SetFocus;
end;

end.
