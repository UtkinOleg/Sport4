unit ServerUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  TServerForm = class(TForm)
    ListView1: TListView;
    Button1: TButton;
    Button2: TButton;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ServerForm: TServerForm;

implementation

{$R *.dfm}

end.
