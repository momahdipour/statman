unit MemberUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, XPMenu;

type
  TMemberForm = class(TForm)
    DettachPopup: TPopupMenu;
    StayOnTop1: TMenuItem;
    XPMenu1: TXPMenu;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormHide(Sender: TObject);
    procedure StayOnTop1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MemberForm: TMemberForm;

implementation

uses
  MainUnit;

{$R *.dfm}

procedure TMemberForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  MainForm.DettachAttachPane.Click;
end;

procedure TMemberForm.FormHide(Sender: TObject);
begin
  with MainForm do
  begin
    DettachAttachPane.Caption:='Dettach Member Pane';
    MemberGroup.Parent:=MainForm;
    MemberGroup.Left:=3;
    MemberGroup.Top:=50;
    MemberPane1.Enabled:=True;
    MemberPane2.Enabled:=True;
  end;  
end;

procedure TMemberForm.StayOnTop1Click(Sender: TObject);
begin
  if (Sender as TMenuItem).Checked then
    MemberForm.FormStyle:=fsStayOnTop
  else
    MemberForm.FormStyle:=fsNormal;
end;

end.
