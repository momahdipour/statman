unit GridReplaceUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, xpButton, xpCheckBox, StdCtrls, AdvGrid;

type
  TGridReplaceForm = class(TForm)
    Label1: TLabel;
    TextCombo: TComboBox;
    GroupBox3: TGroupBox;
    CaseSensitive: TxpCheckBox;
    WholeWords: TxpCheckBox;
    Findbtn: TxpButton;
    Closebtn: TxpButton;
    Label2: TLabel;
    ReplaceCombo: TComboBox;
    procedure FindbtnClick(Sender: TObject);
    procedure ClosebtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    Grid:TAdvStringGrid;
    { Public declarations }
  end;

var
  GridReplaceForm: TGridReplaceForm;

implementation

{$R *.dfm}

function InStrPos(const S1,S2:String;CaseCheck:Boolean;var Pos:Integer):Boolean;
var i:Integer;  SubS:String;
begin
  Result:=False;
  if Length(S1)>Length(S2) then Exit;
  if Length(S1)=0 then Exit;
  for i:=1 to (Length(S2)-Length(S1)+1) do
  begin
    SubS:=Copy(S2,i,Length(S1));
    if CaseCheck then
    begin
      if SubS=S1 then
      begin
        Result:=True;
        Pos:=i;
        Exit;
      end;
    end
    else
    begin
      if LowerCase(SubS)=LowerCase(S1) then
      begin
        Result:=True;
        Pos:=i;
        Exit;
      end;
    end;
  end;
end;

procedure TGridReplaceForm.FindbtnClick(Sender: TObject);
var i,j,k:Integer;  S:String;  Found:Boolean;
    Pos:Integer;
begin
  if (TextCombo.Text='') or (ReplaceCombo.Text=TextCombo.Text) then
    Exit;
  if TextCombo.Items.IndexOf(TextCombo.Text)<0 then
    TextCombo.Items.Append(TextCombo.Text);
  if ReplaceCombo.Items.IndexOf(ReplaceCombo.Text)<0 then
    ReplaceCombo.Items.Append(ReplaceCombo.Text);  
  for i:=1 to (Grid.RowCount-2) do
    for j:=1 to (Grid.ColCount-1) do
      if Length(Grid.Cells[j,i])>0 then
      begin
        Found:=False;
        S:=Grid.Cells[j,i];
        if WholeWords.Checked and CaseSensitive.Checked then
        begin
          if S=TextCombo.Text then
            Found:=True;
        end
        else if WholeWords.Checked and not(CaseSensitive.Checked) then
        begin
          if LowerCase(S)=LowerCase(TextCombo.Text) then
            Found:=True;
        end
        else if CaseSensitive.Checked and not(WholeWords.Checked) then
        begin
          if InStrPos(TextCombo.Text,S,True,Pos) then
            Found:=True;
        end
        else
        begin
          if InStrPos(TextCombo.Text,S,False,Pos) then
            Found:=True;
        end;
        if Found then
        begin
          if WholeWords.Checked then
            S:=ReplaceCombo.Text
          else
          begin
            S:=Copy(S,1,Pos-1)+ReplaceCombo.Text+Copy(S,Pos+Length(TextCombo.Text),Length(S)-Length(Copy(S,1,Pos-1)+TextCombo.Text));
            Grid.Cells[j,i]:=S;
          end;
        end;
      end;
end;

procedure TGridReplaceForm.ClosebtnClick(Sender: TObject);
begin
  GridReplaceForm.Hide;
end;

end.
