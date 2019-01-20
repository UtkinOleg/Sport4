unit SelectUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TSelectForm = class(TForm)
    ButtonNext: TButton;
    CancelBtn: TButton;
    GroupBox1: TGroupBox;
    ComboBox1: TComboBox;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SelectForm: TSelectForm;

implementation

{$R *.dfm}

end.
