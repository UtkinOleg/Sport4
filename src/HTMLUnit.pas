unit HTMLUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleCtrls, SHDocVw, ActiveX;

type
  THTMLForm = class(TForm)
    wb: TWebBrowser;
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  HTMLForm: THTMLForm;

implementation

uses Main;

{$R *.dfm}

procedure THTMLForm.FormActivate(Sender: TObject);
var
 i : integer;
begin
 for I := 0 to MainForm.DockTabSet1.Tabs.Count - 1 do
  if Caption = Trim(MainForm.DockTabSet1.Tabs[i]) then
  begin
    MainForm.DockTabSet1.TabIndex := i;
    break;
  end;

 MainForm.N18.Enabled := false;
 MainForm.N9.Enabled := true;
 MainForm.ToolButton4.Enabled := true;
 MainForm.N27.Enabled := true;
 MainForm.ToolButton2.Enabled := true;
 MainForm.FileSaveItem.Enabled := true;
 MainForm.C1.Enabled := false;
 MainForm.ToolButton8.Enabled := false;
 MainForm.N14.Enabled := false;

 MainForm.ToolButton14.Enabled := true;
 MainForm.N24.Enabled := true;

 MainForm.ToolButton11.Enabled := false;
 MainForm.N16.Enabled := false;
 MainForm.ToolButton12.Enabled := false;
 MainForm.N17.Enabled := false;
 MainForm.SB.Panels[0].Text := '';

end;

procedure THTMLForm.FormClose(Sender: TObject; var Action: TCloseAction);
var
 i : integer;
begin
  for I := 0 to MainForm.DockTabSet1.Tabs.Count - 1 do
   if Trim(MainForm.DockTabSet1.Tabs[i]) = Caption then
    begin
     MainForm.DockTabSet1.Tabs.Delete(i);
     break;
    end;

  MainForm.ToolButton14.Enabled := false;
  MainForm.N9.Enabled := false;
  MainForm.ToolButton4.Enabled := false;
  MainForm.N27.Enabled := false;

  if MainForm.MDIChildCount = 1 then
   begin
    MainForm.ToolButton2.Enabled := false;
    MainForm.FileSaveItem.Enabled := false;
    MainForm.DockTabSet1.Visible := false;
    MainForm.N18.Enabled := false;
    MainForm.ToolButton14.Enabled := false;
    MainForm.N24.Enabled := false;
   end;

 Action := caFree;
end;

procedure THTMLForm.FormShow(Sender: TObject);
begin
{ with wb do
    if Document <> nil then
     with Application as IOleobject do
       DoVerb(OLEIVERB_UIACTIVATE, nil, wb, 0, Handle, GetClientRect); }
end;

end.
