unit MenFormDlg;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, OKCANCL1;

type
  TFIODlg = class(TOKBottomDlg)
    HelpBtn: TButton;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    Label1: TLabel;
    ComboBox1: TComboBox;
    ListBox1: TListBox;
    Button1: TButton;
    Button2: TButton;
    Label2: TLabel;
    procedure OKBtnClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FIODlg: TFIODlg;

implementation

uses SelectSport, Main, RazrForm;

{$R *.dfm}

procedure TFIODlg.Button1Click(Sender: TObject);
begin
  inherited;
  if SelectSportDlg.ShowModal = mrOk then
    ListBox1.Items.Add(SelectSportDlg.ListBox1.Items[SelectSportDlg.ListBox1.ItemIndex]);
end;

procedure TFIODlg.Button2Click(Sender: TObject);
begin
  inherited;
  if ListBox1.ItemIndex<>-1 then
   if MessageBOX(Handle, PChar('”далить вид спорта ?'), PChar(MainTitle) ,
    MB_YESNO or MB_ICONQUESTION) = IDYes then
      ListBox1.DeleteSelected;
end;

procedure TFIODlg.FormShow(Sender: TObject);
begin
  inherited;
  LabeledEdit1.SetFocus;
end;

procedure TFIODlg.HelpBtnClick(Sender: TObject);
begin
  Application.HelpContext(HelpContext);
end;

procedure TFIODlg.OKBtnClick(Sender: TObject);
begin
  inherited;
  if Length(LabeledEdit1.Text)=0 then
    MessageBOX(Handle, '¬ведите фамилию спортсмена.', MainTitle ,
    MB_ICONERROR)
  else
  if Length(LabeledEdit2.Text)=0 then
    MessageBOX(Handle, '¬ведите им€ спортсмена.', MainTitle ,
    MB_ICONERROR)
  else
  if Length(LabeledEdit3.Text)=0 then
    MessageBOX(Handle, '¬ведите отчество спортсмена.', MainTitle ,
    MB_ICONERROR)
  else
  if ComboBox1.ItemIndex=-1 then
    MessageBOX(Handle, '”кажите разр€д спортсмена.', MainTitle ,
    MB_ICONERROR)
  else
  if LIstBox1.Items.Count=0 then
    MessageBOX(Handle, '”кажите виды спорта.', MainTitle ,
    MB_ICONERROR)
  else
   ModalResult := mrOk;
end;

end.

