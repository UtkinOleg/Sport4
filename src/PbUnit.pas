unit PbUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls;

type
  TPbForm = class(TForm)
    pb: TProgressBar;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  PbForm: TPbForm;

implementation

{$R *.dfm}

end.
