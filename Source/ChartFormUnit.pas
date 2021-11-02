unit ChartFormUnit;

interface

uses
  Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, TeeProcs, TeEngine, Chart, Series, ComCtrls,
  Menus, ExplBtn, Buttons, TitleButton, Windows, StrUtils, ToolWin;

type
  TChartForm = class(TForm)
    ToolBar6: TToolBar;
    SpeedButton2: TSpeedButton;
    ExplorerButton9: TExplorerButton;
    SaveAsPopup: TPopupMenu;
    SaveAsPicture1: TMenuItem;
    SaveAsMetafile1: TMenuItem;
    TitleButton1: TTitleButton;
    Chart1: TChart;
    Series1: TLineSeries;
    Series2: TBarSeries;
    Series3: THorizBarSeries;
    Series4: TAreaSeries;
    Series5: TPointSeries;
    Series6: TPieSeries;
    Series7: TFastLineSeries;
    SD1: TSaveDialog;
    PSD1: TPrinterSetupDialog;
    CopyPopup: TPopupMenu;
    CopyAsBitmap1: TMenuItem;
    CopyAsMetafile1: TMenuItem;
    ExplorerButton1: TExplorerButton;
    ChartBtn: TExplorerButton;
    ToolButton1: TToolButton;
    OPD1: TOpenDialog;
    ChartPopup: TPopupMenu;
    Print1: TMenuItem;
    N8: TMenuItem;
    SaveAsPicture2: TMenuItem;
    SaveAsMetafile2: TMenuItem;
    N9: TMenuItem;
    CopyAsPicture1: TMenuItem;
    CopyAsMetafile2: TMenuItem;
    Rollbtn: TTitleButton;
    procedure TitleButton1Mousedown(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SaveAsPicture1Click(Sender: TObject);
    procedure SaveAsMetafile1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure CopyAsBitmap1Click(Sender: TObject);
    procedure CopyAsMetafile1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure CopyAsPicture1Click(Sender: TObject);
    procedure CopyAsMetafile2Click(Sender: TObject);
    procedure Print1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormHide(Sender: TObject);
    procedure SaveAsPicture2Click(Sender: TObject);
    procedure SaveAsMetafile2Click(Sender: TObject);
    procedure SD1CanClose(Sender: TObject; var CanClose: Boolean);
    procedure RollbtnMousedown(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    { Private declarations }
    procedure WMSysCommand(var Message:TWMSysCommand);message WM_SYSCOMMAND;
  public
    { Public declarations }
  end;

var
  ChartForm: TChartForm;
  RollHeight:Integer;
  Rolled:Boolean=False;

implementation

uses MainUnit;

{$R *.dfm}


procedure TChartForm.WMSysCommand(var Message:TWMSysCommand);
begin
  if (Message.CmdType and $FFF0 = SC_MINIMIZE) then
    Application.Minimize
  else
    inherited;
end;

procedure TChartForm.TitleButton1Mousedown(Sender: TObject);
begin
  ChartForm.Visible:=False;
  MainForm.DetachBtn.Glyph.Assign(MainForm.Detachbmp.Picture.Bitmap);
  MainForm.DetachBtn.Hint:='Detach Chart';
end;

procedure TChartForm.FormCreate(Sender: TObject);
var
    hwndHandle:THANDLE;
    hMenuHandle:HMENU;
    iPos:Integer;
    hSysMenu:HMENU;
begin
  Chart1.View3DOptions.Orthogonal:=True;
  ChartForm.Chart1.View3DOptions.Orthogonal:=True;
  ChartForm.Chart1.Foot.Font.Color:=clRed;
  KeyPreview:=True;
end;

procedure TChartForm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key=VK_F4) and (ssAlt in Shift) then
    Key:=0;
end;

procedure TChartForm.SaveAsPicture1Click(Sender: TObject);
var S:String;
begin
  SD1.Filter:='Bitmap Files(*.BMP)|*.BMP|All Files|*.*';
  SD1.FilterIndex:=1;
  if SD1.Execute then
  begin
    S:=RightStr(ExtractFileName(SD1.FileName),4);
    if UpperCase(S)<>'.BMP' then
      S:='.bmp'
    else
      S:='';
    Chart1.SaveToBitmapFile(SD1.FileName+S);
  end;
end;

procedure TChartForm.SaveAsMetafile1Click(Sender: TObject);
var S:String;
begin
  SD1.Filter:='Metafiles(*.WMF)|*.WMF|All Files|*.*';
  SD1.FilterIndex:=1;
  if SD1.Execute then
  begin
    S:=RightStr(ExtractFileName(SD1.FileName),4);
    if UpperCase(S)<>'.WMF' then
      S:='.wmf'
    else
      S:='';
    Chart1.SaveToMetafile(SD1.FileName+S);
  end;
end;

procedure TChartForm.SpeedButton2Click(Sender: TObject);
begin
  if PSD1.Execute then
    Chart1.Print;
end;

procedure TChartForm.CopyAsBitmap1Click(Sender: TObject);
begin
  Chart1.CopyToClipboardBitmap;
end;

procedure TChartForm.CopyAsMetafile1Click(Sender: TObject);
begin
  Chart1.CopyToClipboardMetafile(False);
end;

procedure TChartForm.FormActivate(Sender: TObject);
begin
  ChartBtn.ExplorerPopup:=Mainform.ExplorerPopup1;
end;

procedure TChartForm.CopyAsPicture1Click(Sender: TObject);
begin
  Chart1.CopyToClipboardBitmap;
end;

procedure TChartForm.CopyAsMetafile2Click(Sender: TObject);
begin
  Chart1.CopyToClipboardMetafile(False);
end;

procedure TChartForm.Print1Click(Sender: TObject);
begin
  if PSD1.Execute then
    Chart1.Print;
end;

procedure TChartForm.FormShow(Sender: TObject);
var
    hwndHandle:THANDLE;
    hMenuHandle:HMENU;
    iPos:Integer;
    hSysMenu:HMENU;
begin
  MainForm.c1.Checked:=True;
  hwndHandle:=FindWindow(nil,PChar(Caption));
  if (hwndHandle<>0) then
  begin
    hMenuHandle:=GetSystemMenu(hwndHandle,FALSE);
    if (hMenuHandle<>0) then
    begin
      DeleteMenu(hMenuHandle,SC_CLOSE,MF_BYCOMMAND);
      iPos:=GetMenuItemCount(hMenuHandle);
      Dec(iPos);
      if iPos>-1 then
        DeleteMenu(hMenuHandle,iPos,MF_BYPOSITION);
    end;
  end;
  hSysMenu:=GetSystemMenu(Self.Handle,False);
  if hSysMenu<>0 then
  begin
    EnableMenuItem(hSysMenu,SC_CLOSE,MF_BYCOMMAND Or MF_GRAYED);
    DrawMenuBar(Self.Handle);
  end;
end;

procedure TChartForm.FormHide(Sender: TObject);
begin
  MainForm.c1.Checked:=False;
end;

procedure TChartForm.SaveAsPicture2Click(Sender: TObject);
begin
  SaveAsPicture1.Click;
end;

procedure TChartForm.SaveAsMetafile2Click(Sender: TObject);
begin
  SaveAsMetafile1.Click;
end;

procedure TChartForm.SD1CanClose(Sender: TObject; var CanClose: Boolean);
var S:String; R:Integer;
begin
  if SD1.Filter='Bitmap Files(*.BMP)|*.BMP|All Files|*.*' then
  begin
    S:=RightStr(ExtractFileName(SD1.FileName),4);
    if UpperCase(S)<>'.BMP' then
      S:='.bmp'
    else
      S:='';
    if FileExists(SD1.FileName+S) then
    begin
      r:=MessageDlg(SD1.FileName+S+' already exists.'+Chr(13)+'Do you want to replace it?',mtCustom,[mbYes,mbNo],0);
      if r=mrYes then
        CanClose:=True
      else
        CanClose:=False;
    end;
  end
  else if SD1.Filter='Metafiles(*.WMF)|*.WMF|All Files|*.*' then
  begin
    S:=RightStr(ExtractFileName(SD1.FileName),4);
    if UpperCase(S)<>'.WMF' then
      S:='.wmf'
    else
      S:='';
    if FileExists(SD1.FileName+S) then
    begin
      r:=MessageDlg(SD1.FileName+S+' already exists.'+Chr(13)+'Do you want to replace it?',mtCustom,[mbYes,mbNo],0);
      if r=mrYes then
        CanClose:=True
      else
        CanClose:=False;
    end;
  end;
end;

procedure TChartForm.RollbtnMousedown(Sender: TObject);
begin
  if not(Rolled) then
  begin
    Rolled:=True;
    Rollbtn.TipText:='Roll Down';
    RollHeight:=ChartForm.Height;
    ChartForm.Height:=0;
    ChartForm.BorderIcons:=chartForm.BorderIcons-[biMaximize];
  end
  else
  begin
    Rolled:=False;
    Rollbtn.TipText:='Roll Up';
    ChartForm.Height:=RollHeight;
    ChartForm.BorderIcons:=ChartForm.BorderIcons+[biMaximize];
    Application.ProcessMessages;
    ChartForm.Update;
 end;
end;

procedure TChartForm.FormResize(Sender: TObject);
begin
  if Rolled then
    Chartform.Height:=0
  else if not(Rolled) and (ChartForm.Width<201) then
    ChartForm.Width:=201;
  if Chartform.Width<198 then
    ChartForm.Height:=198;
end;

end.
