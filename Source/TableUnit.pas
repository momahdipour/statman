unit TableUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Spin, Buttons, MainUnit, Grids, 
  ComCtrls, ExplBtn, xpCheckBox, xpButton, ExtCtrls, AdvGrid, Languages;

type
  TTableForm = class(TForm)
    PageControl1: TPageControl;
    QuantitativePage: TTabSheet;
    QualitativePage: TTabSheet;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    AutoCheck: TCheckBox;
    BaseCombo1: TComboBox;
    BaseCombo2: TComboBox;
    AssignPanel: TGroupBox;
    AssignLabel: TLabel;
    ManualRadio: TRadioButton;
    SheetRadio: TRadioButton;
    SpanCombo: TComboBox;
    ItemList: TListBox;
    RangeGroupBox: TGroupBox;
    FromLabel: TLabel;
    ToLabel: TLabel;
    ToSpin: TSpinEdit;
    FromSpin: TSpinEdit;
    AssignRange: TBitBtn;
    SpanList: TListBox;
    StaticText1: TStaticText;
    Label3: TLabel;
    Sortbtn: TExplorerButton;
    ShowFreqbtn: TxpButton;
    QFreqAuto: TxpCheckBox;
    Label4: TLabel;
    QSheetLabel: TLabel;
    QGrid: TStringGrid;
    SortPopup: TExplorerPopup;
    SortByValue: TRadioButton;
    SortByTitle: TRadioButton;
    SortByFreq: TRadioButton;
    GroupBox2: TGroupBox;
    Ascending: TRadioButton;
    Descending: TRadioButton;
    CaseSensitive: TxpCheckBox;
    QDataCombo: TComboBox;
    CancelBtn: TxpButton;
    OKbtn: TxpButton;
    Image1: TImage;
    procedure OkBtnClick(Sender: TObject);
    procedure AutoCheckClick(Sender: TObject);
    procedure SheetRadioClick(Sender: TObject);
    procedure ManualRadioClick(Sender: TObject);
    procedure ItemListClick(Sender: TObject);
    procedure SpanListClick(Sender: TObject);
    procedure AssignRangeClick(Sender: TObject);
    procedure BaseCombo2Change(Sender: TObject);
    procedure FormHide(Sender: TObject);
    procedure FromSpinKeyPress(Sender: TObject; var Key: Char);
    procedure FromSpinChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SortByValueClick(Sender: TObject);
    procedure SortByTitleClick(Sender: TObject);
    procedure SortByFreqClick(Sender: TObject);
    procedure SortbtnDropDownClick(Sender: TObject);
    procedure SortPopupClose(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure SortbtnClick(Sender: TObject);
    procedure CancelbtnClick(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure QFreqAutoClick(Sender: TObject);
    procedure QDataComboChange(Sender: TObject);
    procedure QGridMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure ShowFreqbtnClick(Sender: TObject);
    procedure QGridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    Quantitative:Boolean;
    { Private declarations }
  public
    { Public declarations }
    function GetSheet(Name:String):SheetSettings;
    function GetGrid(Name:String):TAdvStringGrid;
    procedure LoadGridAsString(List:TListBox;Grid:TStringGrid);
  end;

var
  TableForm: TTableForm;
  CanCreate:Boolean;
  UseRanges:Boolean;
  RangeTemp:array of RangeType;

implementation

uses HelpLangFormUnit;

{$R *.dfm}

procedure TTableForm.LoadGridAsString(List:TListBox;Grid:TStringGrid);
var i,j:Integer;
begin
  for i:=1 to (Grid.ColCount-1) do
    for j:=1 to (Grid.RowCount-1) do
      if Length(Grid.Cells[i,j])>0 then
        List.Items.Append(Grid.Cells[i,j]);                 
end;

function TTableForm.GetGrid(Name:String):TAdvStringGrid;
begin
  if Name='Data Sheet 1' then Result:=MainForm.SGrid1
  else if Name='Data Sheet 2' then Result:=MainForm.SGrid2
  else if Name='Data Sheet 3' then Result:=MainForm.SGrid3
  else if Name='Data Sheet 4' then Result:=MainForm.SGrid4
  else if Name='Data Sheet 5' then Result:=MainForm.SGrid5;
end;

function TTableForm.GetSheet(Name:String):SheetSettings;
begin
  if Name='Data Sheet 1' then Result:=MainUnit.Sheet1
  else if Name='Data Sheet 2' then Result:=MainUnit.Sheet2
  else if Name='Data Sheet 3' then Result:=MainUnit.Sheet3
  else if Name='Data Sheet 4' then Result:=MainUnit.Sheet4
  else if Name='Data Sheet 5' then Result:=MainUnit.Sheet5;
end;

procedure TTableForm.OkBtnClick(Sender: TObject);
var grid:TAdvStringGrid;
    i,j,AssignCount:Integer;
    S:String;
begin
  if Quantitative and (BaseCombo1.Items.Count=0) then
  begin
    ShowMessage('There is no data sheet with integer or decimal data type.To create a table, you should make a data sheet with integer or decimal data type to provide the table frequency values.');
    Exit;
  end;
  MainForm.NeedChartRefresh:=True;
  if Quantitative then
  begin
    TableIsQualitative:=False;
    if (TableForm.BaseCombo1.Items.Count=1) and (TableForm.BaseCombo2.Items.Count=0) then
    begin
      if Sheet1.Used then MainForm.FillValueList(MainForm.SGrid1)
      else if Sheet2.Used then MainForm.FillValueList(MainForm.SGrid2)
      else if Sheet3.Used then MainForm.FillValueList(MainForm.SGrid3)
      else if Sheet4.Used then MainForm.FillValueList(MainForm.SGrid4)
      else if Sheet5.Used then MainForm.FillValueList(MainForm.SGrid5);
      if VListCount<2 then
      begin
        ShowMessage('You should enter at least two numbers in a sheet to create the frequency table.');
        Exit;
      end;
      MainForm.FillRangeList;
      MainForm.CreateTable;
      CanCreateAChart:=True;
      MainForm.AutosizeTabelColumn;
      Exit;
    end;
    if BaseCombo2.Items.Count=0 then
    begin
      S:=BaseCombo1.Items.Strings[BaseCombo1.ItemIndex];
      if S='Data Sheet 1' then MainForm.FillValueList(MainForm.SGrid1)
      else if S='Data Sheet 2' then MainForm.FillValueList(MainForm.SGrid2)
      else if S='Data Sheet 3' then MainForm.FillValueList(MainForm.SGrid3)
      else if S='Data Sheet 4' then MainForm.FillValueList(MainForm.SGrid4)
      else if S='Data Sheet 5' then MainForm.FillValueList(MainForm.SGrid5);
      MainForm.FillRangeList;
      MainForm.CreateTable;
      CanCreateAChart:=True;
      MainForm.AutosizeTabelColumn;
      Exit;
    end;
  CanCreate:=True;
  UseRanges:=False;
  RangeCount:=0;
  if not(AutoCheck.Checked) then
  begin
    grid:=GetGrid(BaseCombo2.Items.Strings[BaseCombo2.ItemIndex]);
    if AssignPanel.Enabled=False then
    begin
      for i:=1 to (grid.ColCount-1) do
        for j:=1 to (grid.RowCount-1) do
          if Length(grid.Cells[i,j])>0 then
          begin
            S:=grid.Cells[i,j];
            Inc(RangeCount);
            RangeList[RangeCount].LBound:=StrToInt(Copy(S,2,Pos(',',S)-2));
            RangeList[RangeCount].UBound:=StrToInt(Copy(S,Pos(',',S)+1,Length(S)-Pos(',',S)-1));
            RangeNameList[RangeCount]:=S;
          end;
    end
  else
  begin
    if ManualRadio.Checked then
    begin
      AssignCount:=0;
      for i:=0 to (SpanList.Items.Count-1) do
        if SpanList.Items.Strings[i]<>'Not Assigned' then
          Inc(AssignCount);
      if AssignCount<2 then
      begin
        ShowMessage('You should at least assign two ranges.');
        CanCreate:=False;
        Exit;
      end;
      for i:=0 to High(RangeTemp) do
        if not((RangeTemp[i].LBound=0) and (RangeTemp[i].UBound=0)) then
          begin
            Inc(RangeCount);
            RangeList[RangeCount]:=RangeTemp[i];
            RangeNameList[RangeCount]:=ItemList.Items.Strings[i];
            end;
        end
        else if SheetRadio.Checked then
        begin
          grid:=GetGrid(SpanCombo.Items.Strings[SpanCombo.ItemIndex]);
          for i:=1 to (grid.ColCount-1) do
          for j:=1 to (grid.RowCount-1) do
            if Length(grid.Cells[i,j])>0 then
            begin
              S:=grid.Cells[i,j];
              Inc(RangeCount);
              RangeList[RangeCount].LBound:=StrToInt(Copy(S,2,Pos(',',S)-2));
              RangeList[RangeCount].UBound:=StrToInt(Copy(S,Pos(',',S)+1,Length(S)-Pos(',',S)-1));
            end;
          if RangeCount<2 then
          begin
            ShowMessage('There are not enough ranges in '+SpanCombo.Items.Strings[SpanCombo.ItemIndex]+'.');
            CanCreate:=False;
            Exit;
          end;
          AssignCount:=0;
          grid:=GetGrid(BaseCombo2.Items.Strings[BaseCombo2.ItemIndex]);
          for i:=1 to (grid.ColCount-1) do
          begin
            for j:=1 to (grid.RowCount-1) do
              if Length(grid.Cells[i,j])>0 then
              begin
                Inc(AssignCount);
                RangeNameList[AssignCount]:=grid.Cells[i,j];
                if AssignCount=RangeCount then Break;
              end;
            if AssignCount=RangeCount then Break;
          end;
          if AssignCount<RangeCount then
            RangeCount:=AssignCount;
          if RangeCount<2 then
          begin
            ShowMessage('There are not enough ranges in '+BaseCombo2.Items.Strings[BaseCombo2.ItemIndex]+'.');
            CanCreate:=False;
            Exit;
          end;
        end;
      end;
    end;
    if CanCreate then
    begin
        grid:=GetGrid(BaseCombo1.Items.Strings[BaseCombo1.ItemIndex]);
        MainForm.FillValueList(grid);
        if AutoCheck.Checked then
          MainForm.FillRangeList;
        MainForm.CreateTable;
        CanCreateAChart:=True;
        MainForm.AutosizeTabelColumn;
      end;
  end
  else
  begin
    if not(Quantitative) and not(QualitativePage.Enabled) then
      Exit;
    if not(Quantitative) then
      with MainForm do
      begin
        TableGrid.RowCount:=QGrid.RowCount+1;
        for i:=1 to (QGrid.RowCount-1) do
        begin
          TableGrid.Cells[0,i]:=QGrid.Cells[1,i];
          TableGrid.Cells[1,i]:=QGrid.Cells[3,i];
        end;
        CreateFrqTableColumns;
        MainForm.AutosizeTabelColumn;
        CanCreateAChart:=True;
        TableIsQualitative:=True;
      end;
  end;
end;

procedure TTableForm.AutoCheckClick(Sender: TObject);
var i:Integer;
begin
  BaseCombo2.Enabled:=not(AutoCheck.Checked);
  if not(AutoCheck.Checked) then
  begin
    BaseCombo2.Items.Clear;
    MainForm.FillComboSpecific(BaseCombo2,3);
    if BaseCombo2.Items.Count=0 then
      AssignPanel.Enabled:=False
    else
    begin
      AssignPanel.Enabled:=True;
      ManualRadio.Checked:=True;
      ItemList.Items.Clear;
      SpanList.Items.Clear;
      LoadGridAsString(ItemList,GetGrid(BaseCombo2.Items.Strings[0]));
      for i:=1 to ItemList.Items.Count do
        SpanList.Items.Append('Not Assigned');
      SetLength(RangeTemp,ItemList.Items.Count);
      ItemList.ItemIndex:=0;
      SpanList.ItemIndex:=0;
    end;
    if SpanCombo.Items.Count=0 then
    begin
      SheetRadio.Enabled:=False;
      Spancombo.Enabled:=False;
    end;
    MainForm.FillComboSpecific(BaseCombo2,4);
    BaseCombo2.ItemIndex:=0;
  end
  else
  begin
    ItemList.Items.Clear;
    SpanList.Items.Clear;
    AssignPanel.Enabled:=False;
  end;
end;

procedure TTableForm.SheetRadioClick(Sender: TObject);
begin
  SpanCombo.Enabled:=SheetRadio.Checked;
  ItemList.Enabled:=not(SheetRadio.Checked);
  SpanList.Enabled:=not(SheetRadio.Checked);
  AssignRange.Enabled:=not(SheetRadio.Checked);
end;

procedure TTableForm.ManualRadioClick(Sender: TObject);
begin
  SpanCombo.Enabled:=not(ManualRadio.Checked);
  ItemList.Enabled:=ManualRadio.Checked;
  SpanList.Enabled:=ManualRadio.Checked;
  AssignRange.Enabled:=ManualRadio.Checked;
end;

procedure TTableForm.ItemListClick(Sender: TObject);
begin
  SpanList.ItemIndex:=ItemList.ItemIndex;
  if SpanList.Items.Strings[spanList.ItemIndex]='Not Assigned' then
    AssignRange.Caption:='Assign'
  else
    AssignRange.Caption:='Reassign';
    FromSpin.Value:=RangeTemp[ItemList.ItemIndex].LBound;
    ToSpin.Value:=RangeTemp[ItemList.ItemIndex].UBound;
end;

procedure TTableForm.SpanListClick(Sender: TObject);
begin
  ItemList.ItemIndex:=SpanList.ItemIndex;
  if SpanList.Items.Strings[spanList.ItemIndex]='Not Assigned' then
    AssignRange.Caption:='Assign'
  else
    AssignRange.Caption:='Reassign';
    FromSpin.Value:=RangeTemp[ItemList.ItemIndex].LBound;
    ToSpin.Value:=RangeTemp[ItemList.ItemIndex].UBound;
end;

procedure TTableForm.AssignRangeClick(Sender: TObject);
begin
  if ItemList.Items.Count=0 then
  begin
    ShowMessage('There is no item in the list.');
    Exit;
  end;  
  if FromSpin.Value>=ToSpin.Value then
  begin
    ShowMessage('The span lower bound value should be lower than the span upper bound value.');
    Exit;
  end;
  with RangeTemp[SpanList.ItemIndex] do
  begin
    LBound:=FromSpin.Value;
    UBound:=ToSpin.Value;
    SpanList.Items.Strings[SpanList.ItemIndex]:='['+IntToStr(LBound)+','+IntToStr(UBound)+']';
  end;  
end;

procedure TTableForm.BaseCombo2Change(Sender: TObject);
var Sheet:SheetSettings; i:Integer;
begin
  ItemList.Items.Clear;
  SpanList.Items.Clear;
  FromSpin.Value:=0;
  ToSpin.Value:=0;
  Sheet:=GetSheet(BaseCombo2.Items.Strings[BaseCombo2.ItemIndex]);
  if Sheet.TypeIndex=3 then
  begin
    AssignPanel.Enabled:=True;
    ManualRadio.Checked:=True;
    AssignRange.Caption:='Assign';
    LoadGridAsString(ItemList,GetGrid(BaseCombo2.Items.Strings[BaseCombo2.ItemIndex]));
    for i:=1 to ItemList.Items.Count do
      SpanList.Items.Append('Not Assigned');
    SetLength(RangeTemp,ItemList.Items.Count);
    for i:=0 to High(RangeTemp) do
      with RangeTemp[i] do
      begin
        LBound:=0;
        UBound:=0;
      end;
    SpanCombo.Enabled:=False;
  end
  else
  begin
    AssignPanel.Enabled:=False;
    SetLength(RangeTemp,0);
  end;      
end;

procedure TTableForm.FormHide(Sender: TObject);
begin
  SetLength(RangeTemp,0);
end;

procedure TTableForm.FromSpinKeyPress(Sender: TObject; var Key: Char);
begin
  if Key in ['a'..'z','A'..'Z',' ','`','~','!','@','#','$','%','^','&','*','(',')','_','+',']','[','}','{','"',';',':','?','>','<','/','.',',','\','|','='] then
    Key:=Chr(0);
end;

procedure TTableForm.FromSpinChange(Sender: TObject);
var S:String;
begin
  S:=(Sender as TSpinEdit).Text;
  if Length(S)=0 then
    S:='0';
  if (Length(S)>1) and (S[1]='0') then
    S:=Copy(S,2,Length(S)-1)
  else if (Length(S)>2) and (S[1]='-') and (S[2]='0') then
    S:='-'+Copy(S,3,Length(S)-2);
  if S='-' then S:='0';
  if S<>(Sender as TSpinEdit).Text then
    (Sender as TSpinEdit).Text:=S;
end;

procedure TTableForm.FormCreate(Sender: TObject);
begin
  QGrid.Cells[1,0]:='Title';
  QGrid.Cells[2,0]:='Value';
  QGrid.Cells[3,0]:='Frequency';
  Quantitative:=True;
//  ShowMessage(ColorToString(Image1.Picture.Bitmap.Canvas.Pixels[10,20]));
end;

procedure TTableForm.SortByValueClick(Sender: TObject);
begin
  if SortByValue.Checked then
  begin
    Sortbtn.Caption:='Sort By Value';
    CaseSensitive.Enabled:=True;
    CaseSensitive.Repaint;
  end;  
end;

procedure TTableForm.SortByTitleClick(Sender: TObject);
begin
  if SortByTitle.Checked then
  begin
    Sortbtn.Caption:='Sort By Title';
    CaseSensitive.Enabled:=True;
    CaseSensitive.Repaint;
  end;
end;

procedure TTableForm.SortByFreqClick(Sender: TObject);
begin
  if SortByFreq.Checked then
  begin
    Sortbtn.Caption:='Sort By Frequency';
    CaseSensitive.Enabled:=False;
    CaseSensitive.Repaint;
  end;
end;

procedure TTableForm.SortbtnDropDownClick(Sender: TObject);
begin
  Sortbtn.Caption:='Sort By ?[Select]';
end;

procedure TTableForm.SortPopupClose(Sender: TObject);
begin
  if Sortbtn.Caption='Sort By ?[Select]' then
  begin
    if SortByValue.Checked then
      Sortbtn.Caption:='Sort By Value'
    else if SortByTitle.Checked then
      Sortbtn.Caption:='Sort By Title'
    else if SortByFreq.Checked then
      Sortbtn.Caption:='Sort By Frequency';
  end;
end;

procedure TTableForm.FormActivate(Sender: TObject);
var B:Boolean;
begin
  if RebuildFreqTable then
  begin
    BaseCombo1.Enabled:=True;
    BaseCombo2.Enabled:=True;
    AssignPanel.Enabled:=True;
    BaseCombo1.Items.Clear;
    BaseCombo2.Items.Clear;
    SpanCombo.Items.Clear;
    ItemList.Items.Clear;
    SpanList.Items.Clear;
    AutoCheck.Checked:=True;
    AutoCheck.Enabled:=True;
    MainForm.FillComboSpecific(BaseCombo1,1);
    MainForm.FillComboSpecific(BaseCombo1,2);
    MainForm.FillComboSpecific(BaseCombo2,3);
    MainForm.FillComboSpecific(BaseCombo2,4);
    MainForm.FillComboSpecific(SpanCombo,4);
    if SpanCombo.Items.Count>0 then
      SpanCombo.ItemIndex:=0
    else
      SpanCombo.Enabled:=False;
    if BaseCombo1.Items.Count=0 then
    begin
//      ShowMessage('There is no data sheet with integer or decimal data type.To create a table, you should make a data sheet with integer or decimal data type to provide the table frequency values.');
      QuantitativePage.Enabled:=False;
    end
    else
      BaseCombo1.ItemIndex:=0;
    if BaseCombo2.Items.Count=0 then
    begin
      AutoCheck.Enabled:=False;
      BaseCombo2.Enabled:=False;
      AssignPanel.Enabled:=False;
    end
    else
      BaseCombo2.ItemIndex:=0;
    BaseCombo2.Enabled:=not(AutoCheck.Checked);
    AssignPanel.Enabled:=not(AutoCheck.Checked);
    QDataCombo.Items.Clear;
    MainForm.FillComboSpecific(QDataCombo,3);
    MainForm.FillComboSpecific(QDataCombo,1);
    MainForm.FillComboSpecific(QDataCombo,2);
    if QDataCombo.Items.Count>0 then
      QDataCombo.ItemIndex:=0
    else
    begin
      QualitativePage.Enabled:=False;
    end;

  FromSpin.Value:=0;
  ToSpin.Value:=0;
  if SpanCombo.Items.Count>0 then
  begin
    SheetRadio.Enabled:=True;
    SpanCombo.Enabled:=True;
    AssignLabel.Enabled:=True;
  end
  else
  begin
    SheetRadio.Enabled:=False;
    SpanCombo.Enabled:=False;
    AssignLabel.Enabled:=False;
  end;
  AssignPanel.Enabled:=False;
  if BaseCombo2.Items.Count=0 then
  begin
    AutoCheck.Enabled:=False;
    B:=False;
  end
  else
  begin
    AutoCheck.Enabled:=True;
    B:=True;
  end;
  ManualRadio.Checked:=B;
  ManualRadio.Enabled:=B;
  ItemList.Enabled:=B;
  SpanList.Enabled:=B;
  FromSpin.Enabled:=B;
  ToSpin.Enabled:=B;
  AssignRange.Enabled:=B;
  FromLabel.Enabled:=B;
  ToLabel.Enabled:=B;
  end;
  RebuildFreqTable:=False;
end;

procedure TTableForm.SortbtnClick(Sender: TObject);
var i,j,Col,Index:Integer;  S:String;
begin
  if QGrid.RowCount<3 then
    Exit;
  if SortByFreq.Checked then
  begin
    for i:=1 to (QGrid.RowCount-2) do
    begin
      Index:=i;
      for j:=(i+1) to (QGrid.RowCount-1) do
        if Ascending.Checked and (StrToInt(QGrid.Cells[3,j])<StrToInt(QGrid.Cells[3,Index])) then
          Index:=j
        else if Descending.Checked and (StrToInt(QGrid.Cells[3,j])>StrToInt(QGrid.Cells[3,Index])) then
          Index:=j;
      if Index<>i then
      begin
        for j:=1 to 3 do
        begin
          S:=QGrid.Cells[j,i];
          QGrid.Cells[j,i]:=QGrid.Cells[j,Index];
          QGrid.Cells[j,Index]:=S;
        end;
      end;
    end;
  end else if SortByValue.Checked or SortByTitle.Checked then
  begin
    if SortByTitle.Checked then
      Col:=1
    else
      Col:=2;
    for i:=1 to (QGrid.RowCount-2) do
    begin
      Index:=i;
      for j:=(i+1) to (QGrid.RowCount-1) do
        if CaseSensitive.Checked then
        begin
          if Ascending.Checked and (CompareStr(QGrid.Cells[Col,j],QGrid.Cells[Col,Index])<0) then
            Index:=j
          else if Descending.Checked and (CompareStr(QGrid.Cells[Col,j],QGrid.Cells[Col,Index])>0) then
            Index:=j;
        end else
        begin
          if Ascending.Checked and (CompareStr(LowerCase(QGrid.Cells[Col,j]),LowerCase(QGrid.Cells[Col,Index]))<0) then
            Index:=j
          else if Descending.Checked and (CompareStr(LowerCase(QGrid.Cells[Col,j]),LowerCase(QGrid.Cells[Col,Index]))>0) then
            Index:=j;
        end;
      if Index<>i then
      begin
        for j:=1 to 3 do
        begin
          S:=QGrid.Cells[j,i];
          QGrid.Cells[j,i]:=QGrid.Cells[j,Index];
          QGrid.Cells[j,Index]:=S;
        end;
      end;
    end;
  end;
end;

procedure TTableForm.CancelbtnClick(Sender: TObject);
begin
  TableForm.Hide;
end;

procedure TTableForm.PageControl1Change(Sender: TObject);
begin
  if PageControl1.ActivePageIndex=0 then
    Quantitative:=True
  else
    Quantitative:=False;
end;

procedure TTableForm.QFreqAutoClick(Sender: TObject);
begin
  ShowFreqbtn.Enabled:=not(QFreqAuto.Checked);
  ShowFreqbtn.Repaint;
end;

procedure TTableForm.QDataComboChange(Sender: TObject);
var TempGrid:TAdvStringGrid;
    QValues:array of String;
    i,j,k,Count:Integer;
    S:String;
    AddNew:Boolean;
begin
  if QFreqAuto.Checked then
  begin
    TempGrid:=GetGrid(QDataCombo.Items.Strings[QDataCombo.ItemIndex]);
    for i:=1 to (TempGrid.RowCount-2) do
      for j:=1 to (TempGrid.ColCount-1) do
        if Length(TempGrid.Cells[j,i])>0 then
        begin
          AddNew:=True;
          S:=TempGrid.Cells[j,i];
          if High(QValues)>=0 then
            for k:=0 to High(QValues) do
              if QValues[k]=S then
                AddNew:=False;
          if AddNew then
          begin
            SetLength(QValues,High(QValues)+2);
            QValues[High(QValues)]:=S;
          end;
        end;
    QGrid.RowCount:=High(QValues)+2;
    for k:=0 to High(QValues) do
    begin
      Count:=0;
      for i:=1 to (TempGrid.RowCount-1) do
        for j:=1 to (TempGrid.ColCount-1) do
          if TempGrid.Cells[j,i]=QValues[k] then
            Inc(Count);
      QGrid.Cells[2,k+1]:=QValues[k];
      QGrid.Cells[1,k+1]:=QValues[k];
      QGrid.Cells[3,k+1]:=IntToStr(Count);
    end;
    SetLength(QValues,0);
  end;
end;

procedure TTableForm.QGridMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var MoveCursor:Boolean;  i:Integer;
    R:TRect;
begin
  MoveCursor:=False;
  for i:=1 to (QGrid.RowCount-1) do
  begin
    R:=QGrid.CellRect(0,i);
    if (X>R.Left) and (X<R.Right) and (Y>R.Top) and (Y<R.Bottom) then
    begin
      MoveCursor:=True;
      Break;
    end;
  end;
  if MoveCursor then
  begin
    QGrid.Cursor:=crSizeNS;
    QGrid.Options:=QGrid.Options+[goRowMoving];
  end
  else
  begin
    QGrid.Cursor:=crDefault;
    QGrid.Options:=QGrid.Options-[goRowMoving];
  end
end;

procedure TTableForm.ShowFreqbtnClick(Sender: TObject);
begin
  QFreqAuto.Checked:=True;
  QDataCombo.OnChange(QDataCombo);
  QFreqAuto.Checked:=False;
end;

procedure TTableForm.QGridSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  if ACol<>1 then
    CanSelect:=False;
end;

procedure TTableForm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var A:Integer;
begin
  A:=ActiveSheet;
  ActiveSheet:=9;
  with MainForm do
  if Key=VK_F1 then
  begin
    if EnglishHelp.Checked then
    begin
      SetHelpTopic(lIDEngish);
      Application.HelpFile:=ExeDir+'STATMAN.HLP'
    end
    else if PersianHelp.Checked then
    begin
      SetHelpTopic(lIDPersian);
      Application.HelpFile:=ExeDir+'STATMANPERSIAN.HLP';
    end;
    if PromptForLan.Checked then
    begin
      if PromptAsList.Checked then
      begin
        HelpLangForm.Left:=Mouse.CursorPos.X-2;
        if (HelpLangForm.Width-(Screen.Width-Mouse.CursorPos.X))>0 then
          HelpLangForm.Left:=HelpLangForm.Left-(HelpLangForm.Width-(Screen.Width-Mouse.CursorPos.X));
        HelpLangForm.Top:=Mouse.CursorPos.Y;
        if (HelpLangForm.Height-(Screen.Height-Mouse.CursorPos.Y))>0 then
          HelpLangForm.Top:=HelpLangForm.Top-(HelpLangForm.Height-(Screen.Height-Mouse.CursorPos.Y));
        if EnglishHelp.Checked then
          HelpLangForm.LangList.ItemIndex:=0
        else if PersianHelp.Checked then
          HelpLangForm.LangList.ItemIndex:=1;
        HelpLangForm.Show;
        HelpLangForm.LangList.SetFocus;
      end
      else if PromptAsMenu.Checked then
        HelpLangMenu.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    end
    else
      Application.HelpContext(HelpTopic);
  end;
  ActiveSheet:=A;
end;

end.
