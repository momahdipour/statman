unit ChartPrevUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, Buttons, ComCtrls, ToolWin, TitleButton, ImgList,
  ExtCtrls, TeeProcs, TeEngine, Chart, Series;

type
  TChartPrevForm = class(TForm)
    TitleButton1: TTitleButton;
    ToolBar1: TToolBar;
    ToolButton2: TToolButton;
    SaveCopyPopup: TPopupMenu;
    SaveAsPicture1: TMenuItem;
    SaveAsMetafile1: TMenuItem;
    N1: TMenuItem;
    CopyAsPicture1: TMenuItem;
    CopyAsMetafile1: TMenuItem;
    ImageList1: TImageList;
    Chart1: TChart;
    Series1: TBarSeries;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ChartPrevForm: TChartPrevForm;

implementation

{$R *.dfm}


end.
