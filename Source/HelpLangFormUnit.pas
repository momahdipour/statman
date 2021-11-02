unit HelpLangFormUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, xpCheckBox, xpButton, ExtCtrls, RTFLabel;

type
  THelpLangForm = class(TForm)
    LangList: TListBox;
    PromptCheck: TxpCheckBox;
    psvRTFLabel1: TpsvRTFLabel;
    Display: TxpButton;
    Cancelbtn: TxpButton;
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure CancelbtnClick(Sender: TObject);
    procedure FormDeactivate(Sender: TObject);
    procedure DisplayClick(Sender: TObject);
    procedure LangListKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure PromptCheckClick(Sender: TObject);
    procedure LangListDblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormHide(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  HelpLangForm: THelpLangForm;

implementation

uses MainUnit;

{$R *.dfm}

procedure THelpLangForm.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_ESCAPE then
    Cancelbtn.OnClick(Cancelbtn);
end;

procedure THelpLangForm.CancelbtnClick(Sender: TObject);
begin
  Hide;
end;

procedure THelpLangForm.FormDeactivate(Sender: TObject);
begin
  Hide;
end;

procedure THelpLangForm.DisplayClick(Sender: TObject);
begin
  case LangList.ItemIndex of
    0: Application.HelpFile:=ExeDir+'STATMAN.HLP';
    1: Application.HelpFile:=ExeDir+'STATMANPERSIAN.HLP';
  end;
  Application.HelpContext(HelpTopic);
  Hide;
end;

procedure THelpLangForm.LangListKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_RETURN then
    Display.OnClick(Display);
end;

procedure THelpLangForm.PromptCheckClick(Sender: TObject);
begin
  MainForm.PromptForLan.Checked:=not(PromptCheck.Checked);
  MainForm.ShowDefaultLan.Checked:=PromptCheck.Checked;
end;

procedure THelpLangForm.LangListDblClick(Sender: TObject);
begin
  if LangList.ItemIndex>=0 then
    Display.OnClick(Display);
end;

procedure THelpLangForm.FormShow(Sender: TObject);
begin
  Application.NormalizeAllTopMosts;
end;

procedure THelpLangForm.FormHide(Sender: TObject);
begin
  Application.RestoreTopMosts;
end;

end.
