unit UnitPreview;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls, Grids;

type
  TFormPreview = class(TForm)
    re: TRichEdit;
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormPreview: TFormPreview;

implementation

Uses Main;

{$R *.dfm}

procedure TFormPreview.FormActivate(Sender: TObject);
var
 i : integer;
begin
 for I := 0 to MainForm.DockTabSet1.Tabs.Count - 1 do
  if Caption = MainForm.DockTabSet1.Tabs[i] then
  begin
    MainForm.DockTabSet1.TabIndex := i;
    break;
  end;

 MainForm.N18.Enabled := false;
 MainForm.N9.Enabled := true;
 MainForm.ToolButton2.Enabled := false;
 MainForm.FileSaveItem.Enabled := false;
 MainForm.C1.Enabled := false;
 MainForm.ToolButton8.Enabled := false;
 MainForm.N14.Enabled := false;

 MainForm.ToolButton14.Enabled := false;
 MainForm.N24.Enabled := false;

 MainForm.ToolButton11.Enabled := false;
 MainForm.N16.Enabled := false;
 MainForm.ToolButton12.Enabled := false;
 MainForm.N17.Enabled := false;
 MainForm.SB.Panels[0].Text := '';
 MainForm.N27.Enabled := false;
end;

procedure TFormPreview.FormClose(Sender: TObject; var Action: TCloseAction);
var
 i : integer;
begin
  for I := 0 to MainForm.DockTabSet1.Tabs.Count - 1 do
   if MainForm.DockTabSet1.Tabs[i] = Caption then
    begin
     MainForm.DockTabSet1.Tabs.Delete(i);
     break;
    end;

  MainForm.ToolButton14.Enabled := false;
  MainForm.N9.Enabled := false;

  if MainForm.MDIChildCount = 1 then
   begin
    MainForm.N9.Enabled := false;
    MainForm.ToolButton2.Enabled := false;
    MainForm.FileSaveItem.Enabled := false;
    MainForm.DockTabSet1.Visible := false;
    MainForm.N18.Enabled := false;
    MainForm.ToolButton14.Enabled := false;
    MainForm.N24.Enabled := false;
   end;

 Action := caFree;
end;

end.
