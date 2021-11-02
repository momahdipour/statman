unit TipUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls;

const
    TipStrs:array[1..8] of String=(
        'You can jump to the next section of Member Pane by pressing Alt+End and to the previous section of Member Pane by pressing Alt+Home.',
        'You can get a snapshot of an active form easily by pressing Ctrl+Shift+P.',
        'There are two languages supported in this version of StatMan which you can access them in the Tools menu.As you will see in the Tools menu,these two languages are English and Persian.',
        'Now it is possible to save your work project in a file to work on it at another time.',
        'The "Chart History" feature makes it possible to review or reload a chart which you have created before.',
        'You can format each cell of each data sheet.For example you can change the color of a data sheet cell by right clicking on the cell and selecting "Color..." item.',
        'You can save a data sheet as a XLS(MS Excel Sheet file format) file to open and edit it in MS Excel.',
        'The "Fast Calc" feature allows you to calculate basic math expressions very easy and fast.This tool is shown on the statusbar or as a toolbar.'
        );

type
  TTipForm = class(TForm)
    NextTip: TButton;
    CloseTip: TButton;
    ShowTips: TCheckBox;
    PrevTip: TButton;
    StaticText1: TStaticText;
    TipMemo: TRichEdit;
    StartUpTimer: TTimer;
    Panel1: TPanel;
    Image1: TImage;
    Label1: TLabel;
    Memo1: TMemo;
    procedure FormShow(Sender: TObject);
    procedure NextTipClick(Sender: TObject);
    procedure PrevTipClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ShowTipsClick(Sender: TObject);
    procedure StartUpTimerTimer(Sender: TObject);
  private
    { Private declarations }
  public
    TipPage:Integer;
    { Public declarations }
  end;

var
  TipForm: TTipForm;

implementation

uses
  MainUnit;

{$R *.dfm}

procedure TTipForm.FormShow(Sender: TObject);
begin
  TipPage:=1;
  TipMemo.Lines.Clear;
  TipMemo.Lines.Append(TipStrs[1]);
  PrevTip.Enabled:=False;
  NextTip.Enabled:=True;
end;

procedure TTipForm.NextTipClick(Sender: TObject);
begin
PrevTip.Enabled:=True;
  Inc(TipPage);
  if TipPage=8 then
    NextTip.Enabled:=False;
  TipMemo.Lines.Clear;
  TipMemo.Lines.Append(TipStrs[TipPage]);
end;

procedure TTipForm.PrevTipClick(Sender: TObject);
begin
  NextTip.Enabled:=True;
  Dec(TipPage);
  if TipPage=1 then
    PrevTip.Enabled:=False;
  TipMemo.Lines.Clear;
  TipMemo.Lines.Append(TipStrs[TipPage]);
end;

procedure TTipForm.FormCreate(Sender: TObject);
begin
  if MainForm.TipOfTheDay.Tag=1 then
    ShowTips.Checked:=True
  else
    ShowTips.Checked:=False;
  if ShowTips.Checked then
    StartUpTimer.Enabled:=True;
end;

procedure TTipForm.ShowTipsClick(Sender: TObject);
begin
  if ShowTips.Checked then
    MainForm.TipOfTheDay.Tag:=1
  else
    MainForm.TipOfTheDay.Tag:=0;  
end;

procedure TTipForm.StartUpTimerTimer(Sender: TObject);
begin
  StartUpTimer.Enabled:=False;
  ShowModal;
end;

end.
