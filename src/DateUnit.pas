unit DateUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls;

type
  TDateForm = class(TForm)
    Button1: TButton;
    dt: TDateTimePicker;
    Button2: TButton;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DateForm: TDateForm;

implementation

{$R *.dfm}

end.
