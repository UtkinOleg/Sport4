unit OKCNHLP1;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls, 
  Buttons, ExtCtrls, OKCANCL1;

type
  TOKHelpBottomDlg = class(TOKBottomDlg)
    HelpBtn: TButton;
    Edit1: TEdit;
    procedure OKBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKHelpBottomDlg: TOKHelpBottomDlg;

implementation

{$R *.dfm}

procedure TOKHelpBottomDlg.FormShow(Sender: TObject);
begin
  Edit1.SetFocus;
  inherited;
end;

procedure TOKHelpBottomDlg.HelpBtnClick(Sender: TObject);
begin
  Application.HelpContext(HelpContext);
end;

procedure TOKHelpBottomDlg.OKBtnClick(Sender: TObject);
begin
  inherited;
  if Length(Trim(Edit1.Text))>0 then
    ModalResult := mrOk;  
end;

end.
 
