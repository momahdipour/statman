unit MainAboutUnit;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, psvDialogs, psvBasiclbl, psvWebLabel;

type
  TMainAboutBox = class(TForm)
    Panel1: TPanel;
    ProgramIcon: TImage;
    ProductName: TLabel;
    Version: TLabel;
    Copyright: TLabel;
    Comments: TLabel;
    OKButton: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    BitBtn1: TBitBtn;
    psvObjectPropertiesDialog1: TpsvObjectPropertiesDialog;
    psvWebLabel1: TpsvWebLabel;
    procedure OKButtonClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainAboutBox: TMainAboutBox;

implementation

{$R *.dfm}

procedure TMainAboutBox.OKButtonClick(Sender: TObject);
begin
  Hide;
end;

procedure TMainAboutBox.BitBtn1Click(Sender: TObject);
begin
  psvObjectPropertiesDialog1.Execute;
end;

end.
 
