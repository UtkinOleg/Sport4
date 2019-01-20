program s4;

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  Main in 'Main.pas' {MainForm},
  ChildWin in 'ChildWin.pas' {MDIChild},
  about in 'about.pas' {AboutBox},
  Spr in 'Spr.pas' {FormSpr},
  OKCNHLP1 in 'OKCNHLP1.pas' {OKHelpBottomDlg},
  Man in 'Man.pas' {ManForm},
  MenFormDlg in 'MenFormDlg.pas' {FIODlg},
  SelectSport in 'SelectSport.pas' {SelectSportDlg},
  RazrForm in 'RazrForm.pas' {RazrDlg},
  DiagUnit in 'DiagUnit.pas' {DiagForm},
  OutputUnit in 'OutputUnit.pas' {OutputForm},
  RegUnit in 'RegUnit.pas' {RegForm},
  OptionUnit in 'OptionUnit.pas' {OptionForm},
  ProgressUnit in 'ProgressUnit.pas' {ProgressForm},
  SelectUnit in 'SelectUnit.pas' {SelectForm},
  DateUnit in 'DateUnit.pas' {DateForm},
  OKCANCL1 in 'C:\Program Files (x86)\Embarcadero\RAD Studio\7.0\ObjRepos\en\DelphiWin32\OKCANCL1.PAS' {OKBottomDlg},
  HTMLUnit in 'HTMLUnit.pas' {HTMLForm},
  BaseUnit in 'BaseUnit.pas' {BaseForm},
  ServerUnit in 'ServerUnit.pas' {ServerForm},
  IdFTPListParseBase in 'C:\Program Files (x86)\Embarcadero\RAD Studio\7.0\source\Indy\Indy10\Protocols\IdFTPListParseBase.pas';

{$E exe}

{$R *.RES}


function isRunning(aUniqueString:String): Boolean;
var
  hMutex: THandle;
begin
   Result := False;
   hMutex := CreateMutex(nil,False,PChar(aUniqueString));
   if GetLastError = ERROR_ALREADY_EXISTS then
   begin
     ShowMessage('Программа Спорт 4 уже запущена.');
     Result := True;
     CloseHandle(hMutex);
   end;
end;

begin
 if isRunning('SSSport4') then Exit;

// if paramstr(1) = 'SPORT4' then
// begin
  Application.Initialize;
  Application.Title := 'Спорт 4.0';
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.CreateForm(TOKHelpBottomDlg, OKHelpBottomDlg);
  Application.CreateForm(TFIODlg, FIODlg);
  Application.CreateForm(TManForm, ManForm);
  Application.CreateForm(TSelectSportDlg, SelectSportDlg);
  Application.CreateForm(TRazrDlg, RazrDlg);
  Application.CreateForm(TDiagForm, DiagForm);
  Application.CreateForm(TRegForm, RegForm);
  Application.CreateForm(TOptionForm, OptionForm);
  Application.CreateForm(TProgressForm, ProgressForm);
  Application.CreateForm(TSelectForm, SelectForm);
  Application.CreateForm(TDateForm, DateForm);
  Application.CreateForm(TOKBottomDlg, OKBottomDlg);
  Application.CreateForm(TBaseForm, BaseForm);
  Application.CreateForm(TServerForm, ServerForm);
  Application.Run;
// end;
end.
