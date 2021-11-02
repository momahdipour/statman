unit XLSSheetNameUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, xpCheckBox, xpButton;

type
  TXLSSheetName = class(TForm)
    Label1: TLabel;
    SheetName: TEdit;
    NoSheetName: TCheckBox;
    xpButton1: TxpButton;
    OKBtn: TxpButton;
    procedure SheetNameChange(Sender: TObject);
    procedure NoSheetNameClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  XLSSheetName: TXLSSheetName;

implementation

{$R *.dfm}

procedure TXLSSheetName.SheetNameChange(Sender: TObject);
begin
  if Length(SheetName.Text)=0 then
    OKBtn.Enabled:=False
  else
    OKBtn.Enabled:=True;  
end;

procedure TXLSSheetName.NoSheetNameClick(Sender: TObject);
begin
  SheetName.Enabled:=not(NoSheetName.Checked);
  OKBtn.Enabled:=False;
  if NoSheetName.Checked then
    OKBtn.Enabled:=True
  else if Length(SheetName.Text)>0 then
    OKBtn.Enabled:=True;
end;

end.
