unit FormulaUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, Parser10, Menus, ComCtrls,
  ImgList, AnimatedButton;

const
  MaxVisibleChars=20;

type
  TFormulaForm = class(TForm)
    Panel2: TPanel;
    FmlEdit: TEdit;
    FXImage: TImage;
    GroupBox1: TGroupBox;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    SpeedButton7: TSpeedButton;
    SpeedButton8: TSpeedButton;
    SpeedButton9: TSpeedButton;
    SpeedButton10: TSpeedButton;
    SpeedButton11: TSpeedButton;
    SpeedButton12: TSpeedButton;
    SpeedButton13: TSpeedButton;
    SpeedButton14: TSpeedButton;
    SpeedButton15: TSpeedButton;
    SpeedButton16: TSpeedButton;
    SpeedButton17: TSpeedButton;
    SpeedButton18: TSpeedButton;
    SpeedButton19: TSpeedButton;
    SpeedButton20: TSpeedButton;
    SpeedButton21: TSpeedButton;
    SpeedButton22: TSpeedButton;
    SpeedButton23: TSpeedButton;
    SpeedButton24: TSpeedButton;
    SpeedButton25: TSpeedButton;
    Panel1: TPanel;
    BSbtn: TSpeedButton;
    Clrbtn: TSpeedButton;
    GroupBox2: TGroupBox;
    Edit1: TLabeledEdit;
    Testbtn: TBitBtn;
    REdit: TLabeledEdit;
    OKbtn: TBitBtn;
    Cancelbtn: TBitBtn;
    Parser1: TParser;
    FmlAboutPopup: TPopupMenu;
    About1: TMenuItem;
    ImageList1: TImageList;
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton24Click(Sender: TObject);
    procedure BSbtnClick(Sender: TObject);
    procedure ClrbtnClick(Sender: TObject);
    procedure TestbtnClick(Sender: TObject);
    procedure Parser1ParserError(Sender: TObject; E: Exception);
    procedure FormShow(Sender: TObject);
    procedure FmlEditChange(Sender: TObject);
    procedure OKbtnClick(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
    procedure FXImageClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormulaForm: TFormulaForm;
  HHP,HP:LongInt;
  ParserError:Boolean=False;
  NewShow:Boolean=True;
  Formula:String[80];

implementation

uses MainUnit, FmlAboutUnit;

{$R *.dfm}

procedure TFormulaForm.SpeedButton3Click(Sender: TObject);
begin
  Formula:=Formula+(Sender as TSpeedButton).Caption;
  FmlEdit.Text:=FmlEdit.Text+(Sender as TSpeedButton).Caption;
end;

procedure TFormulaForm.SpeedButton24Click(Sender: TObject);
begin
  if (Sender as TSpeedButton).Caption = 'sin' then
    Formula:=Formula+'A'
  else if (Sender as TSpeedButton).Caption = 'cos' then
    Formula:=Formula+'B'
  else if (Sender as TSpeedButton).Caption = 'tan' then
    Formula:=Formula+'C'
  else if (Sender as TSpeedButton).Caption = 'sqrt' then
    Formula:=Formula+'D'
  else if (Sender as TSpeedButton).Caption = 'Exp' then
    Formula:=Formula+'E'
  else if (Sender as TSpeedButton).Caption = 'Ln' then
    Formula:=Formula+'F';
  FmlEdit.Text:=FmlEdit.Text+(Sender as TSpeedButton).Caption+'(';
end;

procedure TFormulaForm.BSbtnClick(Sender: TObject);
var a:Integer;
begin
  If Length(Formula)>=1 then
  begin
    case Formula[Length(Formula)] of
      '0'..'9','.','^','+','-','*','/','X','(',')': a:=1;
      'A'..'C','E': a:=4;
      'F': a:=3;
      'D': a:=5;
    end;
    FmlEdit.Text:=Copy(FmlEdit.Text,1,Length(FmlEdit.Text)-a);
    if Length(Formula)>=1 then
      Formula:=Copy(Formula,1,Length(Formula)-1)
    else
      Formula:='';
  end;
end;

procedure TFormulaForm.ClrbtnClick(Sender: TObject);
begin
  Formula:='';
  FmlEdit.Text:='F(X)=';
end;

procedure TFormulaForm.TestbtnClick(Sender: TObject);
var FX:String;
begin
  Edit1.SetFocus;
  Edit1.SelectAll;
  if Length(Formula)=0 then
  begin
    ShowMessage('You should create a formula to test it with a test value.');
    Exit;
  end;
  FX:=Copy(FmlEdit.Text,6,Length(FmlEdit.Text)-5);
  if pos('X',FX)<=0 then
  begin
    ShowMessage('The formula does not include a "X" variable.');
    Exit;
  end;
  if Length(Edit1.Text)=0 then
  begin
    ShowMessage('You should enter a test value to test the formula.');
    Exit;
  end;  
  if MainForm.IsValidInt(Edit1.Text) or (MainForm.IsValidDec(Edit1.Text)) then
  begin
    ParserError:=False;
    Parser1.Expression:=FX;
    if ParserError then
    begin
      REdit.Text:='{Error}';
      ShowMessage('Unable to calculate the result of F(X) for the test value because of a math error.Please check the formula expression or test value to be valid.');
    end
    else
    begin
      try
        Parser1.X:=MainForm.StrToExt(Edit1.Text);
        REdit.Text:=Format('%g', [Parser1.Value]);
      except
        on EZeroDivide do
        begin
          REdit.Text:='{Error}';
          ShowMessage('The value you have entered has caused a division by zero in the formula.');
        end;
        on E:EInvalidOp do
        begin
          REdit.Text:='{Error}';
          if MainForm.StrToExt(Edit1.Text)=0 then
            ShowMessage('The value you have entered has caused a division by zero in the formula.')
          else
            ShowMessage(E.Message);
        end;
      end;
    end;
  end
  else
  begin
    REdit.Text:='{Invalid Value}';
    ShowMessage('The test value you have entered is not a valid number.');
    Edit1.SelectAll;
    Edit1.SetFocus;
  end;
end;

procedure TFormulaForm.Parser1ParserError(Sender: TObject; E: Exception);
begin
  ParserError:=True;
end;

procedure TFormulaForm.FormShow(Sender: TObject);
var i:Integer;
begin
  HHP:=Application.HintHidePause;
  HP:=Application.HintPause;
  Application.HintHidePause:=9000000;
  Application.HintPause:=0;
  Formula:='';
  i:=6;
  while i<=Length(FmlEdit.Text) do
  begin
    case FmlEdit.Text[i] of
      'X','0'..'9','(',')','+','-','*','/','^','.': begin
        Formula:=Copy(FmlEdit.Text,i,1);
        Inc(i);
      end;
      's': begin
        if FmlEdit.Text[i+1]='i' then
        begin
          Formula:=Formula+'A';
          Inc(i,4);
        end
        else
        begin
          Formula:=Formula+'D';
          Inc(i,5);
        end;
      end;
      'c': begin
        Formula:=Formula+'B';
        Inc(i,4);
      end;
      't': begin
        Formula:=Formula+'C';
        Inc(i,4);
      end;
      'E': begin
        Formula:=Formula+'E';
        Inc(i,4);
      end;
      'L': begin
        Formula:=Formula+'F';
        Inc(i,3);
      end;
    end; {case}
  end; {while}
  FXImage.Visible:=True;
end;

procedure TFormulaForm.FmlEditChange(Sender: TObject);
var FX:String;
begin
  FmlEdit.Hint:=FmlEdit.Text;
  if Cancelbtn.Visible then
  begin
    OKbtn.Enabled:=True;
    if Length(Formula)<1 then
       OKbtn.Enabled:=False;
    FX:=Copy(FmlEdit.Text,6,Length(FmlEdit.Text)-5);
    if pos('X',FX)<=0 then
      OKbtn.Enabled:=False;
    ParserError:=False;
    Parser1.Expression:=FX;
    if ParserError then
      OKbtn.Enabled:=False;
  end;
end;

procedure TFormulaForm.OKbtnClick(Sender: TObject);
begin
  NewShow:=True;
  Application.HintHidePause:=HHp;
  Application.HintPause:=HP;
end;

procedure TFormulaForm.About1Click(Sender: TObject);
begin
  FmlAboutForm.ShowModal;
end;

procedure TFormulaForm.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
  if Key=Chr(13) then
  begin
    Key:=Chr(0);
    Testbtn.Click;
  end;  
end;

procedure TFormulaForm.FXImageClick(Sender: TObject);
begin
  FmlAboutPopup.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
end;

end.
