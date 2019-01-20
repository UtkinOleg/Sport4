unit RegUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TRegForm = class(TForm)
    CancelBtn: TButton;
    HelpBtn: TButton;
    OKBtn: TButton;
    LabeledEdit2: TLabeledEdit;
    Label1: TLabel;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  RegForm: TRegForm;

implementation

uses Main;

{$R *.dfm}

end.
