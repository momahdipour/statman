unit GridFindUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, xpButton, xpCheckBox, StdCtrls, AdvGrid, MainUnit;

type
  TGridFindForm = class(TForm)
    Label1: TLabel;
    TextCombo: TComboBox;
    GroupBox3: TGroupBox;
    Findbtn: TxpButton;
    Closebtn: TxpButton;
    CaseSensitive: TxpCheckBox;
    WholeWords: TxpCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure TextComboChange(Sender: TObject);
    procedure ClosebtnClick(Sender: TObject);
    procedure CaseSensitiveClick(Sender: TObject);
    procedure WholeWordsClick(Sender: TObject);
    procedure FindbtnClick(Sender: TObject);
    procedure FormHide(Sender: TObject);
    procedure TextComboKeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormActivate(Sender: TObject);
  private
    FirstSearch:Boolean;
    SearchedPoints:array of TPoint;
    { Private declarations }
  public
    Grid:TAdvStringGrid;
    { Public declarations }
  end;

var
  GridFindForm: TGridFindForm;

implementation

{$R *.dfm}

procedure TGridFindForm.FormCreate(Sender: TObject);
begin
  FirstSearch:=True;
end;

procedure TGridFindForm.TextComboChange(Sender: TObject);
begin
  FirstSearch:=True;
end;

procedure TGridFindForm.ClosebtnClick(Sender: TObject);
begin
  GridFindForm.Hide;
end;

procedure TGridFindForm.CaseSensitiveClick(Sender: TObject);
begin
  FirstSearch:=True;
end;

procedure TGridFindForm.WholeWordsClick(Sender: TObject);
begin
  FirstSearch:=True;
end;

function InStr(const S1,S2:String;CaseCheck:Boolean):Boolean;
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
        Exit;
      end;
    end
    else
    begin
      if LowerCase(SubS)=LowerCase(S1) then
      begin
        Result:=True;
        Exit;
      end;
    end;
  end;
end;

procedure TGridFindForm.FindbtnClick(Sender: TObject);
var i,j,k:Integer;  S:String;  Found:Boolean;
    P:TPoint;
begin
  if TextCombo.Items.IndexOf(TextCombo.Text)<0 then
    TextCombo.Items.Append(TextCombo.Text);
  if FirstSearch then
  begin
    SetLength(SearchedPoints,0);
    FirstSearch:=False;
  end;
  Found:=False;
  for i:=1 to (Grid.RowCount-2) do
    for j:=1 to (Grid.ColCount-1) do
      if Length(Grid.Cells[j,i])>0 then
      begin
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
          if InStr(TextCombo.Text,S,True) then
            Found:=True;
        end
        else
        begin
          if InStr(TextCombo.Text,S,False) then
            Found:=True;
        end;
        for k:=0 to High(searchedPoints) do
          if (SearchedPoints[k].X=j) and (SearchedPoints[k].Y=i) then
          begin
            Found:=False;
            Break;
          end;
        if Found then
        begin
          SetLength(SearchedPoints,High(SearchedPoints)+2);
          P.X:=j;
          P.Y:=i;
          SearchedPoints[High(SearchedPoints)]:=P;
          Grid.Row:=i;
          Grid.Col:=j;
          Exit;
        end;
      end;
  if not(Found) then
  begin
    if High(SearchedPoints)=-1 then
      ShowMessage('Could not find any occurrence of serach string.')
    else
      ShowMessage('Search string was not found.');
    FirstSearch:=True;
  end;      
end;

procedure TGridFindForm.FormHide(Sender: TObject);
begin
  SetLength(SearchedPoints,0);
end;

procedure TGridFindForm.TextComboKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key=Chr(13)) and Findbtn.Enabled then
  begin
    Key:=Chr(0);
    Findbtn.OnClick(Findbtn);
  end;  
end;

procedure TGridFindForm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_ESCAPE then
    Closebtn.OnClick(Closebtn);
end;

procedure TGridFindForm.FormActivate(Sender: TObject);
begin
  if (ActiveSheet<1) or (ActiveSheet>5) then
    Findbtn.Enabled:=False
  else
    Findbtn.Enabled:=True;
  Findbtn.Repaint;    
end;

end.
